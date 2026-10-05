using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Net.Http.Headers;
using Azure.Core;
using Azure.AI.Projects;
using Azure.Identity;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Agents.AI;
using ModelContextProtocol.Client;
using UglyToad.PdfPig;

const string defaultEndpoint = "https://pocs-agents-petkosto-resource.services.ai.azure.com/api/projects/pocs-agents-petkosto";
const string defaultDeployment = "gpt-5.4-mini";

if (args.Length == 0 || args[0] is "--help" or "-h")
{
	Console.WriteLine("Usage: DocSummaryAgent <document-path> [--workiq] [--copy-to-onedrive]");
	return;
}

var filePath = Path.GetFullPath(args[0]);
if (!File.Exists(filePath))
{
	Console.Error.WriteLine($"Document not found: {filePath}");
	return;
}

var sources = ExtractSources(filePath);
if (sources.Count == 0)
{
	Console.Error.WriteLine("No extractable text was found. Scanned PDFs need OCR before they can be summarized.");
	return;
}

var endpoint = Environment.GetEnvironmentVariable("AZURE_FOUNDRY_PROJECT_ENDPOINT") ?? defaultEndpoint;
var deployment = Environment.GetEnvironmentVariable("AZURE_FOUNDRY_PROJECT_DEPLOYMENT_NAME") ?? defaultDeployment;
var tools = new List<Microsoft.Extensions.AI.AITool>();
McpClient? workIqClient = null;

try
{
	if (args.Contains("--workiq", StringComparer.OrdinalIgnoreCase))
	{
		workIqClient = await McpClient.CreateAsync(new StdioClientTransport(new()
		{
			Name = "Work IQ",
			Command = OperatingSystem.IsWindows() ? "npx.cmd" : "npx",
			Arguments = ["-y", "@microsoft/workiq", "mcp"]
		}));

		var availableTools = await workIqClient.ListToolsAsync();
		tools.AddRange(availableTools
			.Where(tool => IsReadOnlyWorkIqTool(tool.Name))
			.Cast<Microsoft.Extensions.AI.AITool>());
		Console.WriteLine($"Work IQ connected: {tools.Count} read-only tools available.");
	}

	var agent = new AIProjectClient(new Uri(endpoint), new DefaultAzureCredential())
		.AsAIAgent(
			model: deployment,
			name: "DocSummaryAgent",
			instructions: "Summarize the supplied document concisely. Identify important considerations, risks, decisions, and next steps. Every claim about the document must cite one or more source IDs exactly as [S001]. Use only the supplied source text for document claims. If Work IQ tools are available and the user asks for related workplace context, keep that context in a separate section and distinguish it from the document. Never claim a file was copied or a message was sent unless the application confirms that action.",
			tools: tools);

	var sourceText = string.Join("\n\n", sources.Select(source => $"[{source.Id}] {source.Location}\n{source.Text}"));
	var prompt = $"Analyze this document: {Path.GetFileName(filePath)}\n\nReturn: (1) a short summary, (2) considerations and risks, (3) decisions or next steps. Cite every document-based point using source IDs like [S001]. Do not invent details.\n\nDOCUMENT SOURCES\n{sourceText}";
	var response = await agent.RunAsync(prompt);
	var answer = response.ToString();

	foreach (Match match in Regex.Matches(answer, @"\[(S\d{3})\]"))
	{
		var source = sources.FirstOrDefault(item => item.Id == match.Groups[1].Value);
		if (source is null)
		{
			Console.Error.WriteLine($"Warning: response included unknown citation {match.Value}.");
			continue;
		}

		answer = answer.Replace(match.Value, $"[{source.Location}]", StringComparison.Ordinal);
	}

	Console.WriteLine($"\nDocument: {Path.GetFileName(filePath)}\n");
	Console.WriteLine(answer);

	if (args.Contains("--copy-to-onedrive", StringComparer.OrdinalIgnoreCase))
	{
		var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
		var clientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
		if (string.IsNullOrWhiteSpace(tenantId) || string.IsNullOrWhiteSpace(clientId))
		{
			throw new InvalidOperationException("Set AZURE_TENANT_ID and AZURE_CLIENT_ID for a public-client app registration with delegated Microsoft Graph Files.ReadWrite permission.");
		}

		Console.Write("OneDrive destination folder (blank for root): ");
		var folder = Console.ReadLine()?.Trim().Trim('/');
		if (folder is null || folder.Split('/').Any(segment => segment is "." or ".."))
		{
			Console.WriteLine("OneDrive copy canceled.");
			return;
		}

		var relativePath = string.IsNullOrEmpty(folder)
			? Path.GetFileName(filePath)
			: $"{folder}/{Path.GetFileName(filePath)}";
		var destination = string.Join('/', relativePath.Split('/').Select(Uri.EscapeDataString));
		Console.WriteLine($"OneDrive destination: /{relativePath}");
		Console.Write($"Copy '{filePath}' to this destination? [y/N]: ");
		if (!string.Equals(Console.ReadLine()?.Trim(), "y", StringComparison.OrdinalIgnoreCase))
		{
			Console.WriteLine("OneDrive copy canceled.");
			return;
		}

		await UploadToOneDriveAsync(filePath, destination, tenantId, clientId);
	}
}
catch (Exception exception)
{
	Console.Error.WriteLine($"Operation failed: {exception.Message}");
	Environment.ExitCode = 1;
}
finally
{
	if (workIqClient is not null)
	{
		await workIqClient.DisposeAsync();
	}
}

static bool IsReadOnlyWorkIqTool(string name)
{
	var normalized = name.Replace("_", string.Empty, StringComparison.Ordinal).ToLowerInvariant();
	return normalized.EndsWith("retrieve", StringComparison.Ordinal)
		|| normalized.EndsWith("ask", StringComparison.Ordinal)
		|| normalized.EndsWith("fetch", StringComparison.Ordinal)
		|| normalized.EndsWith("searchpaths", StringComparison.Ordinal)
		|| normalized.EndsWith("getschema", StringComparison.Ordinal)
		|| normalized.EndsWith("fetchblob", StringComparison.Ordinal)
		|| normalized.EndsWith("callfunction", StringComparison.Ordinal);
}

static async Task UploadToOneDriveAsync(string filePath, string destination, string tenantId, string clientId)
{
	const int chunkSize = 10 * 1024 * 1024;
	var options = new DeviceCodeCredentialOptions
	{
		TenantId = tenantId,
		ClientId = clientId,
		DeviceCodeCallback = (code, _) =>
		{
			Console.WriteLine(code.Message);
			return Task.CompletedTask;
		}
	};
	var credential = new DeviceCodeCredential(options);
	var token = await credential.GetTokenAsync(new TokenRequestContext(["Files.ReadWrite"]));
	var sessionUri = new Uri($"https://graph.microsoft.com/v1.0/me/drive/root:/{destination}:/createUploadSession");
	var payload = JsonSerializer.Serialize(new
	{
		item = new Dictionary<string, string>
		{
			["@microsoft.graph.conflictBehavior"] = "fail",
			["name"] = Path.GetFileName(filePath)
		}
	});

	using var client = new HttpClient();
	using var sessionRequest = new HttpRequestMessage(HttpMethod.Post, sessionUri);
	sessionRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
	sessionRequest.Content = new StringContent(payload, Encoding.UTF8, "application/json");
	using var sessionResponse = await client.SendAsync(sessionRequest);
	if (!sessionResponse.IsSuccessStatusCode)
	{
		var details = await sessionResponse.Content.ReadAsStringAsync();
		throw new InvalidOperationException($"OneDrive upload session failed ({(int)sessionResponse.StatusCode}): {details}");
	}

	using var sessionJson = await JsonDocument.ParseAsync(await sessionResponse.Content.ReadAsStreamAsync());
	var uploadUrl = sessionJson.RootElement.GetProperty("uploadUrl").GetString()
		?? throw new InvalidOperationException("OneDrive did not return an upload URL.");
	await using var fileStream = File.OpenRead(filePath);
	var totalLength = fileStream.Length;
	var offset = 0L;
	while (offset < totalLength)
	{
		var length = (int)Math.Min(chunkSize, totalLength - offset);
		var buffer = new byte[length];
		await fileStream.ReadExactlyAsync(buffer);

		using var chunkRequest = new HttpRequestMessage(HttpMethod.Put, uploadUrl);
		chunkRequest.Content = new ByteArrayContent(buffer);
		chunkRequest.Content.Headers.ContentRange = new ContentRangeHeaderValue(offset, offset + length - 1, totalLength);
		using var chunkResponse = await client.SendAsync(chunkRequest);
		if (!chunkResponse.IsSuccessStatusCode)
		{
			var details = await chunkResponse.Content.ReadAsStringAsync();
			throw new InvalidOperationException($"OneDrive file upload failed ({(int)chunkResponse.StatusCode}): {details}");
		}

		offset += length;
	}

	Console.WriteLine("OneDrive copy complete.");
}

static List<DocumentSource> ExtractSources(string path)
{
	var extension = Path.GetExtension(path).ToLowerInvariant();
	var result = new List<DocumentSource>();

	if (extension == ".pdf")
	{
		using var document = PdfDocument.Open(path);
		foreach (var page in document.GetPages())
		{
			AddSource(result, $"PDF page {page.Number}", page.Text);
		}
	}
	else if (extension == ".docx")
	{
		using var document = WordprocessingDocument.Open(path, false);
		var body = document.MainDocumentPart?.Document?.Body;
		var heading = "Document body";
		var paragraphNumber = 0;

		if (body is not null)
		{
			foreach (var paragraph in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
			{
				var text = paragraph.InnerText.Trim();
				if (text.Length == 0)
				{
					continue;
				}

				paragraphNumber++;
				var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? string.Empty;
				if (style.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
				{
					heading = text;
				}

				AddSource(result, $"DOCX section '{heading}', paragraph {paragraphNumber}", text);
			}
		}
	}
	else if (extension is ".txt" or ".md" or ".markdown")
	{
		var lines = File.ReadAllLines(path);
		var heading = "Document body";
		var paragraph = new StringBuilder();
		var startLine = 1;

		for (var index = 0; index <= lines.Length; index++)
		{
			var line = index < lines.Length ? lines[index] : string.Empty;
			if (line.StartsWith('#'))
			{
				FlushParagraph();
				heading = line.TrimStart('#', ' ').Trim();
				continue;
			}

			if (string.IsNullOrWhiteSpace(line))
			{
				FlushParagraph();
				continue;
			}

			if (paragraph.Length == 0)
			{
				startLine = index + 1;
			}
			else
			{
				paragraph.Append(' ');
			}
			paragraph.Append(line.Trim());
		}

		void FlushParagraph()
		{
			if (paragraph.Length == 0)
			{
				return;
			}

			AddSource(result, $"{Path.GetExtension(path).TrimStart('.').ToUpperInvariant()} section '{heading}', line {startLine}", paragraph.ToString());
			paragraph.Clear();
		}
	}
	else
	{
		throw new NotSupportedException("Supported formats are PDF, DOCX, TXT, and Markdown.");
	}

	return result;
}

static void AddSource(List<DocumentSource> sources, string location, string text)
{
	text = Regex.Replace(text, @"\s+", " ").Trim();
	if (text.Length == 0)
	{
		return;
	}

	var index = sources.Count + 1;
	sources.Add(new DocumentSource($"S{index:D3}", location, text));
}

sealed record DocumentSource(string Id, string Location, string Text);

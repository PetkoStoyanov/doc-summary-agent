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

var envFilePath = Path.Combine(Directory.GetCurrentDirectory(), ".env");
if (File.Exists(envFilePath))
{
	DotNetEnv.Env.NoClobber().Load(envFilePath);
}

if (args.Length == 0 || args[0] is "--help" or "-h")
{
	Console.WriteLine("Usage: DocSummaryAgent <document-path> [--workiq] [--copy-to-onedrive] [--send-email] [--send-teams]");
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

var tools = new List<Microsoft.Extensions.AI.AITool>();
McpClient? workIqClient = null;

try
{
	var endpoint = GetRequiredSetting("AZURE_FOUNDRY_PROJECT_ENDPOINT");
	var deployment = GetRequiredSetting("AZURE_FOUNDRY_PROJECT_DEPLOYMENT_NAME");
	Console.Error.WriteLine("Attempting Azure authentication using DefaultAzureCredential...");
	LogAzureAuthDiagnostics();
	var credential = CreateAzureCredential();

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

	var agent = new AIProjectClient(new Uri(endpoint), credential)
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

	var copyToOneDrive = args.Contains("--copy-to-onedrive", StringComparer.OrdinalIgnoreCase);
	var sendEmail = args.Contains("--send-email", StringComparer.OrdinalIgnoreCase);
	var sendTeams = args.Contains("--send-teams", StringComparer.OrdinalIgnoreCase);
	DeviceCodeCredential? graphCredential = null;
	if (copyToOneDrive || sendEmail || sendTeams)
	{
		graphCredential = CreateGraphCredential();
	}

	if (copyToOneDrive)
	{
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

		await UploadToOneDriveAsync(filePath, destination, graphCredential!);
	}

	if (sendEmail)
	{
		await PrepareAndSendEmailAsync(Path.GetFileName(filePath), answer, graphCredential!);
	}

	if (sendTeams)
	{
		await PrepareAndSendTeamsMessageAsync(Path.GetFileName(filePath), answer, graphCredential!);
	}
}
catch (AuthenticationFailedException exception)
{
	Console.Error.WriteLine("Azure authentication failed.");
	LogAzureAuthDiagnostics();
	DumpException(exception);
	Console.Error.WriteLine("Sign in with Azure CLI first: az login");
	Console.Error.WriteLine("If your tenant is not the default one, use: az login --tenant <tenant-id>");
	Environment.ExitCode = 1;
}
catch (Exception exception)
{
	Console.Error.WriteLine("Operation failed.");
	LogAzureAuthDiagnostics();
	DumpException(exception);
	Environment.ExitCode = 1;
}
finally
{
	if (workIqClient is not null)
	{
		await workIqClient.DisposeAsync();
	}
}

static void LogAzureAuthDiagnostics()
{
	var endpoint = Environment.GetEnvironmentVariable("AZURE_FOUNDRY_PROJECT_ENDPOINT");
	var deployment = Environment.GetEnvironmentVariable("AZURE_FOUNDRY_PROJECT_DEPLOYMENT_NAME");
	var tenant = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
	var clientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
	var subscriptionId = Environment.GetEnvironmentVariable("AZURE_SUBSCRIPTION_ID");

	Console.Error.WriteLine("=== Azure auth diagnostics ===");
	Console.Error.WriteLine($"Current directory: {Directory.GetCurrentDirectory()}");
	Console.Error.WriteLine($"Endpoint configured: {!string.IsNullOrWhiteSpace(endpoint)}");
	Console.Error.WriteLine($"Deployment configured: {!string.IsNullOrWhiteSpace(deployment)}");
	Console.Error.WriteLine($"Tenant configured: {!string.IsNullOrWhiteSpace(tenant)}");
	Console.Error.WriteLine($"Client ID configured: {!string.IsNullOrWhiteSpace(clientId)}");
	Console.Error.WriteLine($"Subscription configured: {!string.IsNullOrWhiteSpace(subscriptionId)}");
	if (!string.IsNullOrWhiteSpace(endpoint)) Console.Error.WriteLine($"Endpoint: {endpoint}");
	if (!string.IsNullOrWhiteSpace(deployment)) Console.Error.WriteLine($"Deployment: {deployment}");
	if (!string.IsNullOrWhiteSpace(tenant)) Console.Error.WriteLine($"Tenant: {tenant}");
	if (!string.IsNullOrWhiteSpace(clientId)) Console.Error.WriteLine($"Client ID: {clientId}");
	if (!string.IsNullOrWhiteSpace(subscriptionId)) Console.Error.WriteLine($"Subscription: {subscriptionId}");
	Console.Error.WriteLine("Credential chain: EnvironmentCredential -> ManagedIdentityCredential -> AzureCliCredential -> VisualStudioCodeCredential -> others");
	Console.Error.WriteLine("==============================");
}

static void DumpException(Exception exception)
{
	var current = exception;
	var depth = 0;
	while (current is not null)
	{
		Console.Error.WriteLine($"Exception[{depth}]: {current.GetType().Name}: {current.Message}");
		if (current is AggregateException aggregate)
		{
			foreach (var inner in aggregate.InnerExceptions)
			{
				Console.Error.WriteLine($"  Inner: {inner.GetType().Name}: {inner.Message}");
			}
		}
		current = current.InnerException;
		depth++;
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

static TokenCredential CreateAzureCredential()
{
	var tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
	var options = new DefaultAzureCredentialOptions
	{
		ExcludeVisualStudioCredential = true,
		ExcludeAzurePowerShellCredential = true,
		ExcludeInteractiveBrowserCredential = true
	};

	if (!string.IsNullOrWhiteSpace(tenantId))
	{
		options.TenantId = tenantId;
		options.VisualStudioTenantId = tenantId;
		options.SharedTokenCacheTenantId = tenantId;
	}

	return new DefaultAzureCredential(options);
}

static string GetRequiredSetting(string name)
{
	var value = Environment.GetEnvironmentVariable(name);
	if (string.IsNullOrWhiteSpace(value))
	{
		throw new InvalidOperationException($"Missing required setting '{name}'. Set it in .env or the process environment.");
	}

	return value;
}

static DeviceCodeCredential CreateGraphCredential()
{
	return new DeviceCodeCredential(new DeviceCodeCredentialOptions
	{
		TenantId = GetRequiredSetting("AZURE_TENANT_ID"),
		ClientId = GetRequiredSetting("AZURE_CLIENT_ID"),
		DeviceCodeCallback = (code, _) =>
		{
			Console.WriteLine(code.Message);
			return Task.CompletedTask;
		}
	});
}

static async Task PrepareAndSendEmailAsync(string fileName, string summary, TokenCredential credential)
{
	Console.Write("Email recipient address: ");
	var recipient = Console.ReadLine()?.Trim();
	if (string.IsNullOrWhiteSpace(recipient) || !string.Equals(new System.Net.Mail.MailAddress(recipient).Address, recipient, StringComparison.OrdinalIgnoreCase))
	{
		throw new InvalidOperationException("Enter one valid email address.");
	}

	Console.Write($"Email subject [Document summary: {fileName}]: ");
	var subject = Console.ReadLine()?.Trim();
	if (string.IsNullOrWhiteSpace(subject))
	{
		subject = $"Document summary: {fileName}";
	}

	Console.WriteLine($"\nEmail recipient: {recipient}\nSubject: {subject}\n\n{summary}");
	Console.Write("Send this email? [y/N]: ");
	if (!IsConfirmed())
	{
		Console.WriteLine("Email canceled.");
		return;
	}

	var payload = new
	{
		message = new
		{
			subject,
			body = new { contentType = "Text", content = summary },
			toRecipients = new[] { new { emailAddress = new { address = recipient } } }
		}
	};
	await SendGraphMessageAsync("https://graph.microsoft.com/v1.0/me/sendMail", "https://graph.microsoft.com/Mail.Send", payload, credential, "Email");
}

static async Task PrepareAndSendTeamsMessageAsync(string fileName, string summary, TokenCredential credential)
{
	Console.Write("Teams destination [1: existing chat, 2: channel]: ");
	var destinationType = Console.ReadLine()?.Trim();
	string path;
	string target;
	string scope;
	if (destinationType == "1")
	{
		Console.Write("Existing Teams chat ID: ");
		var chatId = Console.ReadLine()?.Trim();
		if (string.IsNullOrWhiteSpace(chatId))
		{
			throw new InvalidOperationException("A Teams chat ID is required.");
		}

		target = $"Chat ID: {chatId}";
		path = $"chats/{Uri.EscapeDataString(chatId)}/messages";
		scope = "https://graph.microsoft.com/ChatMessage.Send";
	}
	else if (destinationType == "2")
	{
		Console.Write("Teams team ID: ");
		var teamId = Console.ReadLine()?.Trim();
		Console.Write("Teams channel ID: ");
		var channelId = Console.ReadLine()?.Trim();
		if (string.IsNullOrWhiteSpace(teamId) || string.IsNullOrWhiteSpace(channelId))
		{
			throw new InvalidOperationException("Both a Teams team ID and channel ID are required.");
		}

		target = $"Team ID: {teamId}\nChannel ID: {channelId}";
		path = $"teams/{Uri.EscapeDataString(teamId)}/channels/{Uri.EscapeDataString(channelId)}/messages";
		scope = "https://graph.microsoft.com/ChannelMessage.Send";
	}
	else
	{
		throw new InvalidOperationException("Choose 1 for an existing chat or 2 for a channel.");
	}

	var message = $"Document summary: {fileName}\n\n{summary}";
	Console.WriteLine($"\nTeams destination:\n{target}\n\n{message}");
	Console.Write("Send this Teams message? [y/N]: ");
	if (!IsConfirmed())
	{
		Console.WriteLine("Teams message canceled.");
		return;
	}

	var payload = new { body = new { contentType = "text", content = message } };
	await SendGraphMessageAsync($"https://graph.microsoft.com/v1.0/{path}", scope, payload, credential, "Teams message");
}

static bool IsConfirmed() => string.Equals(Console.ReadLine()?.Trim(), "y", StringComparison.OrdinalIgnoreCase);

static async Task SendGraphMessageAsync(string requestUri, string scope, object payload, TokenCredential credential, string actionName)
{
	var token = await credential.GetTokenAsync(new TokenRequestContext([scope]), CancellationToken.None);
	using var client = new HttpClient();
	using var request = new HttpRequestMessage(HttpMethod.Post, requestUri);
	request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
	request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
	using var response = await client.SendAsync(request);
	if (!response.IsSuccessStatusCode)
	{
		var details = await response.Content.ReadAsStringAsync();
		throw new InvalidOperationException($"{actionName} failed ({(int)response.StatusCode}): {details}");
	}

	Console.WriteLine($"{actionName} accepted by Microsoft Graph.");
}

static async Task UploadToOneDriveAsync(string filePath, string destination, TokenCredential credential)
{
	const int chunkSize = 10 * 1024 * 1024;
	var token = await credential.GetTokenAsync(new TokenRequestContext(["https://graph.microsoft.com/Files.ReadWrite"]), CancellationToken.None);
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

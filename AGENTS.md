# Project Context

This is a local .NET 10 C# console application using Microsoft Agent Framework (MAF) with the existing Microsoft Foundry project. Do not provision Azure resources.

## Requirements

- Accept PDF, DOCX, TXT, and Markdown documents.
- Produce a concise summary and considerations, with citations to exact source pages, sections, paragraphs, or lines.
- Optionally use Work IQ for related Microsoft 365 context.
- Copy a document to OneDrive and prepare email or Teams messages only with explicit approval for each action.

## Existing Foundry Configuration

- Project endpoint: `https://pocs-agents-petkosto-resource.services.ai.azure.com/api/projects/pocs-agents-petkosto`
- Deployment: `gpt-5.4-mini`
- Azure resource group: `rg-pocs-agents-petkosto`
- Authenticate locally with Azure CLI credentials; do not put credentials in source files.

## Work IQ Safety

- Work IQ is started on demand with `npx -y @microsoft/workiq mcp`.
- Tenant sign-in and administrator consent may be required.
- Discover tools from the connected server; never assume a tool exists.
- Keep write tools unavailable to the model. Copy, email, and Teams operations must show the exact target and content and receive per-action confirmation before execution.
- Work IQ preview file upload is not available. Use an explicitly authorized Microsoft Graph upload path for local files if implemented.

## Current Implementation Status

- `Program.cs` currently extracts PDF/DOCX/TXT/Markdown text, asks the Foundry model for a cited summary, and has an optional read-only Work IQ MCP connection.
- OneDrive upload and email/Teams sending are not implemented yet.
- The latest build failed because `AIProjectClient` is unresolved. First try adding the `Azure.AI.Projects` namespace import appropriate to the installed MAF package, then rebuild and address remaining diagnostics before adding more features.
- Last verified command: `dotnet build DocSummaryAgent.csproj`; dependency-only build succeeded before the `Program.cs` implementation.

Read this file before continuing work in this project.
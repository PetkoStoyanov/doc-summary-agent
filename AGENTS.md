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

## Environment Files

- Copy `.env.example` to `.env` for local settings. `.env` is ignored by Git.
- Process environment variables take precedence over values from `.env`.
- Foundry project endpoint and deployment name are required settings.
- `AZURE_TENANT_ID` and `AZURE_CLIENT_ID` are required only when using `--copy-to-onedrive`, `--send-email`, or `--send-teams`.

## Work IQ Safety

- Work IQ is started on demand with `npx -y @microsoft/workiq mcp`.
- Tenant sign-in and administrator consent may be required.
- Discover tools from the connected server; never assume a tool exists.
- Keep write tools unavailable to the model. Copy, email, and Teams operations must show the exact target and content and receive per-action confirmation before execution.
- Work IQ preview file upload is not available. Use an explicitly authorized Microsoft Graph upload path for local files if implemented.

## OneDrive Copy

- Copying is opt-in with `--copy-to-onedrive` and requires a separate explicit `y` confirmation after showing the full source and OneDrive destination.
- Configure `AZURE_TENANT_ID` and `AZURE_CLIENT_ID` for a public-client app registration with delegated Microsoft Graph `Files.ReadWrite` permission and device-code flow enabled.
- Uploads use Microsoft Graph upload sessions with conflict behavior `fail`; existing files are not overwritten. Upload is chunked in 10 MiB ranges.
- No OneDrive sign-in or upload has been exercised in this environment; tenant consent and app registration configuration are still required.

## Email and Teams

- `--send-email` previews the recipient, subject, and exact text body, then requires a separate explicit `y` confirmation before sending through Microsoft Graph.
- `--send-teams` requires an existing chat ID or a team and channel ID, previews the target and exact message, then requires a separate explicit `y` confirmation.
- Configure only the delegated Graph permissions needed: `Mail.Send`, `ChatMessage.Send`, and/or `ChannelMessage.Send`, in addition to `Files.ReadWrite` if using OneDrive.
- Sending is performed by application code after confirmation; write tools are not exposed to the model.
- No email or Teams send has been exercised in this environment; tenant consent and app registration configuration are still required.

## Current Implementation Status

- `Program.cs` loads local settings from `.env` without overriding process environment variables, extracts PDF/DOCX/TXT/Markdown text, asks the Foundry model for a cited summary, supports an optional read-only Work IQ MCP connection, and can copy to OneDrive or send email/Teams messages after per-action confirmation.
- `AIProjectClient` resolves through the `Azure.AI.Projects` namespace, and the latest build completed without warnings or errors.
- Last verified commands: `dotnet build DocSummaryAgent.csproj` and `dotnet run --no-build -- --help`.

## Resume Point

- The OneDrive, email, and Teams Graph actions compile; each is opt-in and requires its own explicit confirmation. Teams content is sent as plain text.
- No OneDrive upload, email, or Teams message has been exercised against a tenant.
- Next: copy `.env.example` to `.env`, configure the public-client app registration and delegated Graph permissions, then test OneDrive first, email second, and Teams third. Never commit `.env` or send without reviewing and confirming the displayed target and content.

Read this file before continuing work in this project.
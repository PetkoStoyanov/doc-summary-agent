# Copilot instructions for DocSummaryAgent

## Project context
- This is a local .NET 10 console app for summarizing documents with Azure AI Foundry.
- Use .env for local config; never commit secrets or credentials.
- Do not add Azure provisioning logic to this project.
- Preserve the existing Microsoft Agent Framework + Foundry setup.

## Coding expectations
- Keep generated code minimal and focused.
- Avoid extra abstractions, helpers, or boilerplate unless required.
- Prefer small, direct fixes over broad refactors.
- Keep explanations concise and focused on the issue being solved.
- Ask before introducing optional future-facing features or new files.

## Security and auth
- Never hardcode tenant IDs, client IDs, keys, or tokens.
- Prefer Azure CLI sign-in for local auth.
- Use environment variables and `.env` only for local configuration.
- If Microsoft Graph write actions are involved, require explicit user confirmation before sending or uploading.

## Validation
- Validate with the smallest relevant command, usually `dotnet build` for this project.
- When debugging auth or Azure SDK issues, verify CLI login state first.

## Working style
- Respect AGENTS.md as the authoritative project context.
- Preserve the current app behavior unless the request requires a change.
- When handling document inputs, support PDF, DOCX, TXT, and Markdown files only.

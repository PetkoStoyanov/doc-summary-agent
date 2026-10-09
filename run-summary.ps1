param(
    [Parameter(Mandatory = $true)]
    [string]$File,

    [switch]$WorkIq,
    [switch]$CopyToOneDrive,
    [switch]$SendEmail,
    [switch]$SendTeams
)

$arguments = @($File)

if ($WorkIq) { $arguments += "--workiq" }
if ($CopyToOneDrive) { $arguments += "--copy-to-onedrive" }
if ($SendEmail) { $arguments += "--send-email" }
if ($SendTeams) { $arguments += "--send-teams" }

& dotnet run -- $arguments

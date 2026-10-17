# Forces the tenant-wide Teams setting that Graph reads for transcript APIs.
# The Teams admin center toggle can show On before Graph starts accepting requests.
# Requires the MicrosoftTeams module and a Teams administrator account.

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
  Write-Host "Installing MicrosoftTeams PowerShell module (CurrentUser)..."
  Install-Module MicrosoftTeams -Scope CurrentUser -Force
}

Import-Module MicrosoftTeams
Connect-MicrosoftTeams

Write-Host "Current Global meeting configuration:"
Get-CsTeamsMeetingConfiguration -Identity Global |
  Select-Object Identity, EnableGraphTranscriptAccess, EnableAttributedTranscripts |
  Format-List

Set-CsTeamsMeetingConfiguration `
  -Identity Global `
  -EnableGraphTranscriptAccess $true `
  -EnableAttributedTranscripts $true

Write-Host "Updated Global meeting configuration:"
Get-CsTeamsMeetingConfiguration -Identity Global |
  Select-Object Identity, EnableGraphTranscriptAccess, EnableAttributedTranscripts |
  Format-List

Write-Host "Graph can take up to several hours to honor this. Click Fetch Data again after it has had time to apply."

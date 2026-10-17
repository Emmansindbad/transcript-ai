# Graph application permission IDs (Microsoft Graph resource appId)
# 00000003-0000-0000-c000-000000000000

$envPath = Resolve-Path (Join-Path (Split-Path $PSScriptRoot -Parent) ".env.local")
$envVars = @{}
Get-Content $envPath | ForEach-Object {
  if ($_ -match '^\s*#' -or $_ -notmatch '=') { return }
  $parts = $_ -split '=', 2
  $envVars[$parts[0].Trim()] = $parts[1].Trim()
}

$clientId = $envVars["AZURE_CLIENT_ID"]
if (-not $clientId) {
  throw "AZURE_CLIENT_ID not found in .env.local"
}

Write-Host "Using app client ID from .env.local: $clientId"
$graphAppId = "00000003-0000-0000-c000-000000000000"

$requiredRoles = @(
  @{ Id = "798ee544-9d2d-430c-a058-570e29e34338"; Name = "Calendars.Read" },
  @{ Id = "c1684f21-1984-47fa-9d61-2dc8c296bb70"; Name = "OnlineMeetings.Read.All" },
  @{ Id = "a4a80d8d-d283-4bd8-8504-555ec3870630"; Name = "OnlineMeetingTranscript.Read.All" },
  @{ Id = "a4a08342-c95d-476b-b943-97e100569c8d"; Name = "OnlineMeetingRecording.Read.All" },
  @{ Id = "df021288-bdef-4463-88db-98f22de89214"; Name = "User.Read.All" },
  @{ Id = "b633e1c5-b582-4048-a93e-9f11b44c7e96"; Name = "Mail.Send" },
  @{ Id = "6b7d71aa-70aa-4810-a8d9-5d9fb2830017"; Name = "Chat.Read.All" },
  @{ Id = "b9bb2381-47a4-46cd-aafb-00cb12f68504"; Name = "ChatMessage.Read.All" },
  @{ Id = "44e666d1-d276-445b-a5fc-8815eeb81d55"; Name = "Tasks.ReadWrite.All" }
)

Import-Module Microsoft.Graph.Applications

Connect-MgGraph -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All" -UseDeviceAuthentication -NoWelcome

$app = Get-MgApplication -Filter "appId eq '$clientId'" -Property "id,appId,displayName,requiredResourceAccess"
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Property "id,appId,appRoles"
$appSp = Get-MgServicePrincipal -Filter "appId eq '$clientId'" -Property "id,appId,displayName"

Write-Host "App: $($app.displayName)"
Write-Host "App Object ID: $($app.id)"
Write-Host "Service Principal ID: $($appSp.id)"

$resourceAccess = @(
  foreach ($role in $requiredRoles) {
    @{
      id   = $role.Id
      type = "Role"
    }
  }
)

$requiredResourceAccess = @(
  @{
    resourceAppId  = $graphAppId
    resourceAccess = $resourceAccess
  }
)

Update-MgApplication -ApplicationId $app.id -BodyParameter @{
  requiredResourceAccess = $requiredResourceAccess
}

Write-Host "Updated app registration requiredResourceAccess."

foreach ($role in $requiredRoles) {
  $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSp.id |
    Where-Object { $_.AppRoleId -eq $role.Id }

  if ($existing) {
    Write-Host "Already granted: $($role.Name)"
    continue
  }

  New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSp.id -BodyParameter @{
    principalId = $appSp.id
    resourceId  = $graphSp.id
    appRoleId   = $role.Id
  } | Out-Null

  Write-Host "Granted admin consent: $($role.Name)"
}

Write-Host "--- Verification ---"
Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSp.id |
  ForEach-Object {
    $appRole = $graphSp.AppRoles | Where-Object { $_.Id -eq $_.AppRoleId }
    [PSCustomObject]@{
      Permission = $appRole.Value
      AppRoleId  = $_.AppRoleId
    }
  } | Format-Table -AutoSize

Write-Host "Done. Restart npm run dev and retry the app."

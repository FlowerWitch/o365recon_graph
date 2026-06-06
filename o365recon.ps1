# GraphRecon - Modern retrieval of Entra ID / O365 tenant information
# Updated to use the modern Microsoft.Graph PowerShell SDK
#
# Requirements: 
# Open PowerShell window as Admin
# Run: Install-Module Microsoft.Graph -Scope CurrentUser

param(
    [Parameter(Mandatory=$false)]
    [String]$ProjectName
)

######################################################################################
# CONNECTING TO MICROSOFT GRAPH

Write-Host "Connecting to Microsoft Graph Services..." -ForegroundColor Cyan

Write-Host -NoNewline "`tChecking for Microsoft.Graph Module ... "
if (Get-Module -ListAvailable -Name Microsoft.Graph) {
    Write-Host "DONE" -ForegroundColor Green
} else {
    Write-Host "FAILED" -ForegroundColor Red
    Write-Host "`tPlease install the Microsoft.Graph Module using:`n`t`tInstall-Module Microsoft.Graph -Scope CurrentUser"
    exit
}

# Define all required permissions for a comprehensive tenant audit
$scopes = @(
    "Organization.Read.All",
    "Domain.Read.All",
    "User.Read.All",
    "Group.Read.All",
    "Directory.Read.All",
    "Device.Read.All",
    "Application.Read.All",
    "RoleManagement.Read.Directory",
    "Policy.Read.All"
)

try {
    Write-Host "`tPrompting for Microsoft Graph Authentication..." -ForegroundColor Yellow
    Connect-MgGraph -UseDeviceAuthentication
    #Connect-MgGraph -UseDeviceAuthentication
    Write-Host "`tConnected successfully to Microsoft Graph!" -ForegroundColor Green
} catch {
    Write-Host "Could not connect to Microsoft Graph." -ForegroundColor Red
    throw $_
    exit
}

######################################################################################
# Setting up working directory

if (-not $ProjectName) {
    $ProjectName = Read-Host -Prompt "Please enter a project name"
}
$ProjectName = $ProjectName -replace '[^a-zA-Z0-9]',''

$pathIsOK = $false
while ($pathIsOK -eq $false) {
    if (-not (Test-Path $ProjectName)) {
        try {
            New-Item -ItemType Directory -Path $ProjectName > $null
            $CURRENTJOB = "./$ProjectName/$ProjectName"
            $pathIsOK = $true
        } catch {
            Write-Host "Whoops, failed to create folder." -ForegroundColor Red
        }
    } else {
        $ProjectName = Read-Host -Prompt "Folder exists. Please enter a different project name"
        $ProjectName = $ProjectName -replace '[^a-zA-Z0-9]',''
    }
}

######################################################################################
# GET COMPANY AND DOMAIN INFO

Write-Host -NoNewline "`tRetrieving Company & Sync Info ... "
$org = Get-MgOrganization | Select-Object -First 1
$syncConfig = Get-MgDirectoryOnPremiseSynchronization | Select-Object -First 1

$org | Format-List -Property * | Out-File -FilePath .\${CURRENTJOB}.OrgInfo.txt
Write-Host "DONE" -ForegroundColor Green

# Generate Report Header
$reportPath = ".\${CURRENTJOB}.Report.txt"
@"
------------------------------------------------------------------------------------
Overview
Company Name: $($org.DisplayName)
Tenant ID: $($org.Id)
Technical Contacts: $($org.TechnicalNotificationMails -join ', ')
------------------------------------------------------------------------------------
Directory Sync Configuration (Replaced SyncConfiguration)
Directory Synchronization Enabled: $($org.OnPremisesSyncEnabled)
Last Dir Sync Time: $($syncConfig.Configuration.SynchronizationInterval)
Accidental Deletion Prevention: $($syncConfig.Configuration.AccidentalDeletionPrevention.DeletionPreventionStatus)
Threshold Count: $($syncConfig.Configuration.AccidentalDeletionPrevention.ThresholdCount)
------------------------------------------------------------------------------------
"@ | Tee-Object -FilePath $reportPath

# Get Domain Info
Write-Host -NoNewline "`tRetrieving Domain Information ... "
Get-MgDomain | Select-Object Id, IsDefault, IsInitial, AuthenticationType | Format-Table -AutoSize | Out-File -FilePath .\${CURRENTJOB}.Domains.txt
Write-Host "DONE" -ForegroundColor Green

######################################################################################
# USER INFO

Write-Host -NoNewline "`tRetrieving User List (fetching required fields explicitly) ... "
# Note: Graph requires explicit selection of certain properties to avoid empty returns
$userSelect = @("id","userPrincipalName","displayName","department","jobTitle","businessPhones","officeLocation","onPremisesLastSyncDateTime","accountEnabled")
$userlist = Get-MgUser -All -Property $userSelect
Write-Host "DONE ($($userlist.Count) users found)" -ForegroundColor Green

# Simple User List
$userlist.UserPrincipalName | Out-File -FilePath .\${CURRENTJOB}.Users_Simple.txt

# Detailed CSV
Write-Host -NoNewline "`tCreating Detailed User Export ... "
$userlist | Where-Object { $_.UserPrincipalName -notlike "HealthMailbox*" } | Select-Object UserPrincipalName, DisplayName, Department, JobTitle, OfficeLocation, OnPremisesLastSyncDateTime, AccountEnabled | Export-Csv -Path .\${CURRENTJOB}.Users_Detailed.csv -NoTypeInformation
Write-Host "DONE" -ForegroundColor Green

######################################################################################
# GROUP & MEMBERSHIP INFO

Write-Host -NoNewline "`tRetrieving Group & Membership Profiles ... "
$grouplist = Get-MgGroup -All -Property "id","displayName","description","groupTypes"
$grouplist | Select-Object DisplayName, Description, GroupTypes | Format-Table -AutoSize | Out-File -FilePath .\${CURRENTJOB}.Groups_All.txt

# Iterate memberships and hunt targeted keywords (admin, vpn, etc.)
foreach ($group in $grouplist) {
    $members = Get-MgGroupMember -GroupId $group.Id -All 2>$null
    if ($members) {
        foreach ($member in $members) {
            # Map upn if user, or id if object type differs
            $memberUPN = if ($member.AdditionalProperties.containsKey("userPrincipalName")) { $member.AdditionalProperties["userPrincipalName"] } else { $member.Id }
            $logLine = "$($group.DisplayName):$memberUPN"
            
            # General output
            $logLine | Out-File -Append -FilePath .\${CURRENTJOB}.GroupMembership_All.txt
            
            # Targeted keywords filter
            if ($group.DisplayName -match "admin|dirsync|aad") {
                $logLine | Out-File -Append -FilePath .\${CURRENTJOB}.GroupMembership_AdminGroups.txt
            }
            if ($group.DisplayName -match "vpn|cisco|globalprotect|palo") {
                $logLine | Out-File -Append -FilePath .\${CURRENTJOB}.GroupMembership_VPNGroups.txt
            }
        }
    }
}
Write-Host "DONE" -ForegroundColor Green

######################################################################################
# ROLE MEMBERSHIP (ADMINS)

Write-Host -NoNewline "`tIterating Active Directory Directory Roles ... "
$directoryRoles = Get-MgDirectoryRole -All 2>$null
foreach ($role in $directoryRoles) {
    $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All 2>$null
    foreach ($member in $roleMembers) {
        $memberUPN = if ($member.AdditionalProperties.containsKey("userPrincipalName")) { $member.AdditionalProperties["userPrincipalName"] } else { $member.Id }
        "$($role.DisplayName):$memberUPN" | Out-File -Append -FilePath .\${CURRENTJOB}.Roles_Admins.txt
    }
}
Write-Host "DONE" -ForegroundColor Green

######################################################################################
# DEVICE INFO

Write-Host -NoNewline "`tRetrieving Device Inventories ... "
$devicelist = Get-MgDevice -All -Property "id","displayName","operatingSystem","operatingSystemVersion","trustType","isManaged"
if ($devicelist) {
    $devicelist | Select-Object DisplayName, OperatingSystem, OperatingSystemVersion, TrustType, IsManaged | Export-Csv -Path .\${CURRENTJOB}.Devices_Detailed.csv -NoTypeInformation
}
Write-Host "DONE" -ForegroundColor Green

######################################################################################
# APPLICATION SECURITY & AUTHORIZATION POLICIES

Write-Host -NoNewline "`tAnalyzing Applications & Tenant Authorization Policies ... "
$apps = Get-MgApplication -All
if ($apps) {
    $apps | Format-List -Property * | Out-File -FilePath .\${CURRENTJOB}.Applications_All.txt
}

# Check Authorization Policies (Tenant Settings)
$authPolicy = Get-MgPolicyAuthorizationPolicy | Select-Object -First 1

$policyReport = @(
    "`n------------------------------------------------------------------------------------",
    "Tenant Authorization Policies & Security posture:",
    "Allowed To Create Applications: $($authPolicy.DefaultUserRolePermissions.AllowedToCreateApps)",
    "Allowed To Create Security Groups: $($authPolicy.DefaultUserRolePermissions.AllowedToCreateSecurityGroups)",
    "Block Msol PowerShell Access: $($authPolicy.BlockMsolPowerShell)"
)
$policyReport | Out-File -Append -FilePath $reportPath
Write-Host "DONE" -ForegroundColor Green

######################################################################################
# STATS SUMMARY APPEND

@"
------------------------------------------------------------------------------------
Environment Statistics
Total Environment Users: $($userlist.Count)
Total Environment Groups: $($grouplist.Count)
Total Tracked Devices: $($devicelist.Count)
Total Integrated Applications: $($apps.Count)
------------------------------------------------------------------------------------
"@ | Out-File -Append -FilePath $reportPath

Write-Host "`n[+] RECON JOB COMPLETE. Data collected under directory: ./$ProjectName" -ForegroundColor Green
Get-ChildItem -Path .\$ProjectName

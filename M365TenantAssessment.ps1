#region Description
<#     
       .NOTES
       ==============================================================================
       Created on:         2025/03/03 
       Created by:         Drago Petrovic
       Organization:       MSB365.blog
       Filename:           M365TenantAssessment.ps1
       Current version:    V1.0     

       Find us on:
             * Website:         https://www.msb365.blog
             * Technet:         https://social.technet.microsoft.com/Profile/MSB365
             * LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
             * MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
       ==============================================================================

       .DESCRIPTION
       PowerShell script that generates an HTML report that can be used as Microsoft 365 Tenant assessment.           
       

       .NOTES






       .EXAMPLE
       .\M365TenantAssessment.ps1
             

       .COPYRIGHT
       Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
       to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
       and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
       WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
       ===========================================================================
       .CHANGE LOG
             V1.00, 2025/03/03 - DrPe - Initial version

             
			 




--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>
#endregion
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "Microsoft 365 assessment Script"
$RKEY = "MSB365_M365TenantAssessment"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2025 Drago Petrovic\par
$Scriptname \par
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################
write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
Start-Sleep -s 2
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
###############################################################################


#----------------------------------------------------------------------------------------
$MaximumFunctionCount = 10000
#Requires -Modules ExchangeOnlineManagement, MicrosoftTeams, Microsoft.Graph

# Function to check and install required modules
function Ensure-ModuleInstalled {
    param (
        [string]$ModuleName
    )
    
    Write-Host "Checking for $ModuleName module..."
    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName module not found. Installing..."
        try {
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser
            Write-Host "$ModuleName module installed successfully."
        } catch {
            Write-Warning "Failed to install $ModuleName module. Error: $_"
            return $false
        }
    } else {
        Write-Host "$ModuleName module is already installed."
    }
    return $true
}

# Check and install required modules
$requiredModules = @("Microsoft.Graph", "ExchangeOnlineManagement", "MicrosoftTeams")
$allModulesInstalled = $true
foreach ($module in $requiredModules) {
    if (!(Ensure-ModuleInstalled -ModuleName $module)) {
        $allModulesInstalled = $false
    }
}

if (!$allModulesInstalled) {
    Write-Error "Not all required modules could be installed. Please check the errors and try again."
    exit
}

# Import required modules
Write-Host "Importing required modules..."
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Identity.SignIns
Import-Module ExchangeOnlineManagement
Import-Module MicrosoftTeams
Import-Module Microsoft.Graph.Reports
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Users
try {
    Import-Module Microsoft.Graph.DeviceManagement.Enrolment -ErrorAction Stop
    $intuneModuleAvailable = $true
} catch {
    Write-Warning "Failed to import Microsoft.Graph.DeviceManagement.Enrolment module. Some Intune features may not be available."
    $intuneModuleAvailable = $false
}
Write-Host "Modules imported successfully."

# Function to connect to Microsoft Graph with required permissions
function Connect-ToMicrosoftGraph {
    Write-Host "Connecting to Microsoft Graph..."
    $requiredScopes = @(
        "Directory.Read.All", "Organization.Read.All", "User.Read.All",
        "Group.Read.All", "Application.Read.All", "Policy.Read.All",
        "Reports.Read.All", "Sites.Read.All", "TeamSettings.Read.All",
        "MailboxSettings.Read", "AuditLog.Read.All"
    )
    if ($intuneModuleAvailable) {
        $requiredScopes += "DeviceManagementManagedDevices.Read.All"
    }

    Connect-MgGraph -Scopes $requiredScopes
    Write-Host "Connected to Microsoft Graph successfully."
}

# Function to get tenant configuration
function Get-TenantConfiguration {
    Write-Host "Retrieving tenant configuration..."
    $tenantInfo = Get-MgOrganization
    Write-Host "Tenant configuration retrieved."
    return $tenantInfo
}

# Function to get Exchange Online configuration
function Get-ExchangeOnlineConfiguration {
    Write-Host "Connecting to Exchange Online..."
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
            throw "Failed to connect to Exchange Online or retrieve mailbox information."
        }
        Write-Host "Connected to Exchange Online. Retrieving mailbox information..."
        
        $userMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
        $roomMailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited
        $equipmentMailboxes = Get-Mailbox -RecipientTypeDetails EquipmentMailbox -ResultSize Unlimited

        $mailboxCategories = @{
            UserMailboxes = $userMailboxes
            SharedMailboxes = $sharedMailboxes
            RoomMailboxes = $roomMailboxes
            EquipmentMailboxes = $equipmentMailboxes
        }

        $mailboxStats = @{}
        foreach ($category in $mailboxCategories.Keys) {
            $totalSize = 0
            $mailboxCategories[$category] | ForEach-Object {
                $stats = Get-MailboxStatistics -Identity $_.UserPrincipalName
                $totalSize += [long]($stats.TotalItemSize.Value -replace '[^\d]')
            }
            $mailboxStats[$category] = @{
                Count = $mailboxCategories[$category].Count
                TotalSize = $totalSize
            }
        }

        $top5Mailboxes = Get-Mailbox -ResultSize Unlimited | 
            Get-MailboxStatistics | 
            Sort-Object TotalItemSize -Descending | 
            Select-Object -First 5 DisplayName, @{Name="TotalSizeGB";Expression={[math]::Round(($_.TotalItemSize.Value.ToString() -replace '.*$$(.+) bytes$$.*', '$1') -as [long] / 1GB, 2)}}

        $transportRules = Get-TransportRule
        $domains = Get-AcceptedDomain

        Write-Host "Exchange Online data retrieved. Disconnecting..."
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected from Exchange Online."

        return @{
            MailboxCategories = $mailboxCategories
            MailboxStats = $mailboxStats
            Top5Mailboxes = $top5Mailboxes
            TransportRules = $transportRules
            Domains = $domains
        }
    }
    catch {
        Write-Warning "Failed to connect to Exchange Online: $_"
        return $null
    }
}

# Function to get Microsoft Teams configuration
function Get-TeamsConfiguration {
    Write-Host "Connecting to Microsoft Teams..."
    try {
        Connect-MicrosoftTeams -ErrorAction Stop
        Write-Host "Connected to Microsoft Teams. Retrieving teams and policies..."
        $teams = Get-Team
        $policies = Get-CsTeamsClientConfiguration
        Write-Host "Teams data retrieved. Disconnecting..."
        Disconnect-MicrosoftTeams
        Write-Host "Disconnected from Microsoft Teams."
        return @{
            Teams = $teams
            Policies = $policies
        }
    }
    catch {
        Write-Warning "Failed to connect to Microsoft Teams: $_"
        return $null
    }
}

# Function to get Entra ID configuration
function Get-EntraIDConfiguration {
    Write-Host "Retrieving Entra ID configuration..."
    $users = Get-MgUser -All
    $groups = Get-MgGroup -All
    $applications = Get-MgApplication -All
    Write-Host "Entra ID configuration retrieved."
    return @{
        Users = $users
        Groups = $groups
        Applications = $applications
    }
}

# Function to get access authorizations and guest access
function Get-AccessAuthorizations {
    Write-Host "Retrieving access authorizations and guest access information..."
    $guestUsers = Get-MgUser -Filter "userType eq 'Guest'"
    $roleAssignments = Get-MgRoleManagementDirectoryRoleAssignment
    Write-Host "Access authorizations and guest access information retrieved."
    return @{
        GuestUsers = $guestUsers
        RoleAssignments = $roleAssignments
    }
}

# Function to get security policies
function Get-SecurityPolicies {
    Write-Host "Retrieving security policies..."
    $conditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy
    $authenticationMethodsPolicies = Get-MgPolicyAuthenticationMethodPolicy

    $policiesWithExcludedUsers = foreach ($policy in $conditionalAccessPolicies) {
        $excludedUserIds = $policy.Conditions.Users.ExcludeUsers
        $excludedUsers = @()
        if ($excludedUserIds) {
            $excludedUsers = foreach ($userId in $excludedUserIds) {
                (Get-MgUser -UserId $userId).UserPrincipalName
            }
        }
        [PSCustomObject]@{
            PolicyName = $policy.DisplayName
            ExcludedUsers = $excludedUsers -join ', '
        }
    }

    Write-Host "Security policies retrieved."
    return @{
        ConditionalAccessPolicies = $policiesWithExcludedUsers
        AuthenticationMethodsPolicies = $authenticationMethodsPolicies
    }
}

# Function to get Intune configuration
function Get-IntuneConfiguration {
    Write-Host "Retrieving Intune configuration..."
    try {
        if ($intuneModuleAvailable) {
            $intuneDevices = Get-MgDeviceManagementManagedDevice -All | Select-Object DeviceName, OperatingSystem, LastSyncDateTime, ComplianceState
            $intuneOverview = Get-MgDeviceManagementManagedDeviceOverview
        } else {
            Write-Warning "Intune module not available. Using alternative method to retrieve device information."
            $intuneDevices = Get-MgDevice -All | Select-Object DisplayName, OperatingSystem, ApproximateLastSignInDateTime
            $intuneOverview = @{ ManagedDeviceCount = ($intuneDevices | Measure-Object).Count }
        }
        Write-Host "Intune configuration retrieved."
        return @{
            Devices = $intuneDevices
            Overview = $intuneOverview
        }
    } catch {
        Write-Warning "Failed to retrieve Intune configuration: $_"
        return $null
    }
}

# Function to get MFA status
function Get-MFAStatus {
    Write-Host "Retrieving MFA status..."
    try {
        $mfaStatus = Get-MgReportAuthenticationMethodUserRegistrationDetail
        Write-Host "MFA status retrieved."
        return $mfaStatus
    } catch {
        Write-Warning "Failed to retrieve MFA status: $_"
        return $null
    }
}

# Function to get license information
function Get-LicenseInfo {
    Write-Host "Retrieving license information..."
    try {
        $licenses = Get-MgSubscribedSku
        Write-Host "License information retrieved."
        return $licenses
    } catch {
        Write-Warning "Failed to retrieve license information: $_"
        return $null
    }
}

# Function to get Teams Phone configuration
function Get-TeamsPhoneConfig {
    Write-Host "Retrieving Teams Phone configuration..."
    try {
        $teamsModule = Get-Module -Name MicrosoftTeams -ListAvailable
        if (-not $teamsModule) {
            Write-Warning "MicrosoftTeams module is not installed. Installing..."
            Install-Module -Name MicrosoftTeams -Force -AllowClobber
        }
        Import-Module MicrosoftTeams
        Connect-MicrosoftTeams -ErrorAction Stop
        $teamsPhoneConfig = Get-CsOnlineVoiceRoutingPolicy -ErrorAction Stop
        if ($teamsPhoneConfig) {
            $phoneNumbers = Get-CsOnlineUser | Where-Object { $_.EnterpriseVoiceEnabled -eq $true } | 
                Select-Object UserPrincipalName, LineUri, @{N='BusinessHours';E={$_.TenantDialPlan}}
            Write-Host "Teams Phone configuration retrieved."
            Disconnect-MicrosoftTeams
            return @{
                Configured = $true
                PhoneNumbers = $phoneNumbers
            }
        } else {
            Write-Host "Teams Phone is not configured."
            Disconnect-MicrosoftTeams
            return @{
                Configured = $false
                PhoneNumbers = $null
            }
        }
    } catch {
        Write-Warning "Failed to retrieve Teams Phone configuration: $_"
        return @{
            Configured = $false
            PhoneNumbers = $null
        }
    }
}

# Function to get registered domains
function Get-RegisteredDomains {
    Write-Host "Retrieving registered domains..."
    try {
        $domains = Get-MgDomain
        Write-Host "Registered domains retrieved successfully."
        return $domains
    } catch {
        Write-Warning "Failed to retrieve registered domains: $_"
        return $null
    }
}

# Function to get public DNS settings
function Get-PublicDNSSettings {
    Write-Host "Retrieving public DNS settings for domains..."
    try {
        $domains = Get-MgDomain
        $dnsSettings = foreach ($domain in $domains) {
            $dnsRecords = Resolve-DnsName -Name $domain.Id -Type ALL -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                DomainName = $domain.Id
                DNSRecords = $dnsRecords | Select-Object @{Name='RecordType';Expression={$_.RecordType}}, 
                                                        @{Name='TTL';Expression={$_.TTL}}, 
                                                        @{Name='Records';Expression={$_.Strings -join ', '}}
            }
        }
        Write-Host "Public DNS settings retrieved successfully."
        return $dnsSettings
    } catch {
        Write-Warning "Failed to retrieve public DNS settings: $_"
        return $null
    }
}

# Function to get all Microsoft Teams teams
function Get-AllTeams {
    Write-Host "Retrieving all Microsoft Teams teams..."
    try {
        $teams = Get-MgTeam
        Write-Host "Microsoft Teams teams retrieved successfully."
        return $teams
    } catch {
        Write-Warning "Failed to retrieve Microsoft Teams teams: $_"
        return $null
    }
}

# Function to get SharePoint sites
function Get-SharePointSites {
    Write-Host "Retrieving SharePoint sites..."
    try {
        $sites = Get-MgSite -All
        Write-Host "SharePoint sites retrieved successfully."
        return $sites
    } catch {
        Write-Warning "Failed to retrieve SharePoint sites: $_"
        return $null
    }
}

# Function to get SharePoint sites with external sharing enabled
function Get-ExternalSharingEnabledSites {
    Write-Host "Retrieving SharePoint sites with external sharing enabled..."
    try {
        $sites = Get-MgSite -All
        $externalSharingSites = $sites | Where-Object { $_.SharingCapability -ne 'Disabled' }
        Write-Host "SharePoint sites with external sharing retrieved successfully."
        return $externalSharingSites
    } catch {
        Write-Warning "Failed to retrieve SharePoint sites with external sharing: $_"
        return $null
    }
}

# Function to generate HTML report
function Generate-HTMLReport {
    param (
        $TenantConfig,
        $ExchangeConfig,
        $TeamsConfig,
        $EntraIDConfig,
        $AccessConfig,
        $SecurityPolicies,
        $IntuneConfig,
        $MFAStatus,
        $LicenseInfo,
        $TeamsPhoneConfig,
        $RegisteredDomains,
        $PublicDNSSettings,
        $AllTeams,
        $SharePointSites,
        $ExternalSharingSites
    )

    # Set default values for null parameters
    $TenantConfig = if ($null -eq $TenantConfig) { @{DisplayName='N/A'; Id='N/A'; VerifiedDomains=@()} } else { $TenantConfig }
    $ExchangeConfig = if ($null -eq $ExchangeConfig) { @{MailboxStats=@{}; Top5Mailboxes=@()} } else { $ExchangeConfig }
    $TeamsConfig = if ($null -eq $TeamsConfig) { @{} } else { $TeamsConfig }
    $EntraIDConfig = if ($null -eq $EntraIDConfig) { @{Users=@(); Groups=@(); Applications=@()} } else { $EntraIDConfig }
    $AccessConfig = if ($null -eq $AccessConfig) { @{} } else { $AccessConfig }
    $SecurityPolicies = if ($null -eq $SecurityPolicies) { @{ConditionalAccessPolicies=@()} } else { $SecurityPolicies }
    $IntuneConfig = if ($null -eq $IntuneConfig) { @{Overview=@{ManagedDeviceCount=0}; Devices=@()} } else { $IntuneConfig }
    $MFAStatus = if ($null -eq $MFAStatus) { @() } else { $MFAStatus }
    $LicenseInfo = if ($null -eq $LicenseInfo) { @() } else { $LicenseInfo }
    $TeamsPhoneConfig = if ($null -eq $TeamsPhoneConfig) { @{Configured=$false; PhoneNumbers=@()} } else { $TeamsPhoneConfig }
    $RegisteredDomains = if ($null -eq $RegisteredDomains) { @() } else { $RegisteredDomains }
    $PublicDNSSettings = if ($null -eq $PublicDNSSettings) { @() } else { $PublicDNSSettings }
    $AllTeams = if ($null -eq $AllTeams) { @() } else { $AllTeams }
    $SharePointSites = if ($null -eq $SharePointSites) { @() } else { $SharePointSites }
    $ExternalSharingSites = if ($null -eq $ExternalSharingSites) { @() } else { $ExternalSharingSites }

    Write-Host "Starting HTML report generation..."

    try {
        $outputPath = "C:\Temp\M365TenantAssessment.html"
        Write-Host "Output path set to: $outputPath"

        # Ensure the directory exists
        $directory = Split-Path -Path $outputPath -Parent
        if (-not (Test-Path -Path $directory)) {
            Write-Host "Creating directory: $directory"
            New-Item -ItemType Directory -Force -Path $directory
            Write-Host "Directory created successfully"
        } else {
            Write-Host "Directory already exists: $directory"
        }

        $mfaEnabledCount = ($MFAStatus | Where-Object { $_.IsMfaRegistered -eq $true }).Count
        $mfaNotEnabledUsers = $MFAStatus | Where-Object { $_.IsMfaRegistered -eq $false } | Select-Object UserPrincipalName
        $mfaEnabledPercentage = if (($mfaEnabledCount + $mfaNotEnabledUsers.Count) -gt 0) {
            [math]::Round(($mfaEnabledCount / ($mfaEnabledCount + $mfaNotEnabledUsers.Count)) * 360, 2)
        } else { 0 }

        $adminRoles = Get-MgDirectoryRole | Where-Object { $_.DisplayName -like "*admin*" }
        $adminAccounts = @()
        foreach ($role in $adminRoles) {
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id
            $adminAccounts += $members | Where-Object { $_.AdditionalProperties.userPrincipalName -ne $null } |
                Select-Object @{N='UPN';E={$_.AdditionalProperties.userPrincipalName}}, 
                              @{N='Role';E={$role.DisplayName}}
        }
        $adminAccountsWithLicenses = @()
        foreach ($admin in $adminAccounts) {
            $user = Get-MgUser -UserId $admin.UPN -ErrorAction SilentlyContinue
            if ($user) {
                $userLicenses = $user.AssignedLicenses | ForEach-Object { 
                    ($LicenseInfo | Where-Object { $_.SkuId -eq $_.SkuId }).SkuPartNumber 
                }
                $adminAccountsWithLicenses += [PSCustomObject]@{
                    UPN = $admin.UPN
                    Role = $admin.Role
                    Licenses = ($userLicenses | Select-Object -Unique) -join ', '
                }
            }
        }

        $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Tenant Assessment Report</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        h1, h2 { color: #0078D4; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .pie-chart { width: 200px; height: 200px; border-radius: 50%; background: conic-gradient(#0078d4 0deg ${mfaEnabledPercentage}deg, #83c5be ${mfaEnabledPercentage}deg 360deg); margin: 20px auto; }
        .legend { display: flex; justify-content: center; }
        .legend-item { margin: 0 10px; }
        .teams-phone-configured { background-color: #83f28f; padding: 10px; }
        .teams-phone-not-configured { background-color: #ee6b6e; padding: 10px; }
    </style>
</head>
<body>
    <div class="container">
        <img src="https://www.dotcloud.expert/wp-content/uploads/2025/02/Logo_Long_Plus_B2-2.png" alt="dotCloud Logo" style="max-width: 300px; display: block; margin: 0 auto;">
        <h1>Microsoft 365 Tenant Assessment Report</h1>
        
        <h2>Tenant Configuration</h2>
        <table>
            <tr><th>Display Name</th><td>$($TenantConfig.DisplayName)</td></tr>
            <tr><th>Tenant ID</th><td>$($TenantConfig.Id)</td></tr>
        </table>

        <h2>Registered Domains</h2>
        <table>
            <tr><th>Domain Name</th><th>Status</th><th>Authentication Type</th></tr>
            $(foreach ($domain in $RegisteredDomains) {
                "<tr><td>$($domain.Id)</td><td>$($domain.IsVerified)</td><td>$($domain.AuthenticationType)</td></tr>"
            })
        </table>

        <h2>Exchange Online Configuration</h2>
        <table>
            <tr><th>Mailbox Type</th><th>Count</th><th>Total Size (GB)</th></tr>
            $(foreach ($key in $ExchangeConfig.MailboxStats.Keys) {
                "<tr><td>$key</td><td>$($ExchangeConfig.MailboxStats[$key].Count)</td><td>$([math]::Round($ExchangeConfig.MailboxStats[$key].TotalSize / 1GB, 2))</td></tr>"
            })
        </table>

        <h2>Teams Configuration</h2>
        <p>Total Teams: $($AllTeams.Count)</p>

        <h2>Entra ID Configuration</h2>
        <table>
            <tr><th>Users</th><th>Groups</th><th>Applications</th></tr>
            <tr>
                <td>$($EntraIDConfig.Users.Count)</td>
                <td>$($EntraIDConfig.Groups.Count)</td>
                <td>$($EntraIDConfig.Applications.Count)</td>
            </tr>
        </table>

        <h2>Multi-Factor Authentication (MFA) Status</h2>
        <div class="pie-chart"></div>
        <div class="legend">
            <div class="legend-item">MFA Enabled: $mfaEnabledCount</div>
            <div class="legend-item">MFA Not Enabled: $($mfaNotEnabledUsers.Count)</div>
        </div>
        <h3>Users without MFA:</h3>
        <ul>
            $(foreach ($user in $mfaNotEnabledUsers) {
                "<li>$($user.UserPrincipalName)</li>"
            })
        </ul>

        <h2>License Information</h2>
        <table>
            <tr><th>License</th><th>Assigned</th><th>Total</th></tr>
            $(foreach ($license in $LicenseInfo) {
                "<tr><td>$($license.SkuPartNumber)</td><td>$($license.ConsumedUnits)</td><td>$($license.PrepaidUnits.Enabled)</td></tr>"
            })
        </table>

        <h2>Admin Accounts with Licenses</h2>
        <table>
            <tr><th>UPN</th><th>Role</th><th>Licenses</th></tr>
            $(foreach ($admin in $adminAccountsWithLicenses) {
                "<tr><td>$($admin.UPN)</td><td>$($admin.Role)</td><td>$($admin.Licenses)</td></tr>"
            })
        </table>

        <h2>Conditional Access Policies</h2>
        <table>
            <tr><th>Policy Name</th><th>Excluded Users</th></tr>
            $(foreach ($policy in $SecurityPolicies.ConditionalAccessPolicies) {
                "<tr><td>$($policy.PolicyName)</td><td>$($policy.ExcludedUsers)</td></tr>"
            })
        </table>

        <h2>Intune Configuration</h2>
        <p>Total Managed Devices: $($IntuneConfig.Overview.ManagedDeviceCount)</p>
        <table>
            <tr><th>Device Name</th><th>Operating System</th><th>Compliance State</th></tr>
            $(foreach ($device in $IntuneConfig.Devices) {
                "<tr><td>$($device.DeviceName)</td><td>$($device.OperatingSystem)</td><td>$($device.ComplianceState)</td></tr>"
            })
        </table>

        <h2>Teams Phone Configuration</h2>
        $(if ($TeamsPhoneConfig.Configured) {
            @"
            <div class="teams-phone-configured">
                <p>Teams Phone is configured.</p>
                <table>
                    <tr><th>UPN</th><th>Phone Number</th><th>Business Hours</th></tr>
                    $(foreach ($phone in $TeamsPhoneConfig.PhoneNumbers) {
                        "<tr><td>$($phone.UserPrincipalName)</td><td>$($phone.LineUri)</td><td>$($phone.BusinessHours)</td></tr>"
                    })
                </table>
            </div>
"@
        } else {
            @"
            <div class="teams-phone-not-configured">
                <p>Teams Phone is not configured.</p>
            </div>
"@
        })

        <h2>SharePoint Sites with External Sharing</h2>
        <table>
            <tr><th>Site Name</th><th>URL</th><th>Sharing Capability</th></tr>
            $(foreach ($site in $ExternalSharingSites) {
                "<tr><td>$($site.DisplayName)</td><td>$($site.WebUrl)</td><td>$($site.SharingCapability)</td></tr>"
            })
        </table>
    </div>
</body>
</html>
"@

        Write-Host "HTML content generated successfully"

        Write-Host "Attempting to save HTML report..."
        $htmlContent | Out-File -FilePath $outputPath -Encoding UTF8 -Force
        
        if (Test-Path $outputPath) {
            Write-Host "HTML report saved successfully to $outputPath"
            Write-Host "Attempting to open the report in default browser..."
            Start-Process $outputPath
            Write-Host "Report should now be open in your default browser"
        } else {
            throw "File not found after saving attempt"
        }
    }
    catch {
        Write-Error "Error in Generate-HTMLReport: $_"
        Write-Error "Stack Trace: $($_.ScriptStackTrace)"
    }
}

# Main script execution
try {
    Write-Host "Starting Microsoft 365 Tenant Assessment..."
    
    Connect-ToMicrosoftGraph

    Write-Host "Collecting tenant configuration..."
    $tenantConfig = Get-TenantConfiguration

    Write-Host "Collecting Exchange Online configuration..."
    $exchangeConfig = Get-ExchangeOnlineConfiguration

    Write-Host "Collecting Microsoft Teams configuration..."
    $teamsConfig = Get-TeamsConfiguration

    Write-Host "Collecting Entra ID configuration..."
    $entraIDConfig = Get-EntraIDConfiguration

    Write-Host "Collecting access authorizations and guest access information..."
    $accessConfig = Get-AccessAuthorizations

    Write-Host "Collecting security policies..."
    $securityPolicies = Get-SecurityPolicies

    Write-Host "Collecting Intune configuration..."
    $intuneConfig = Get-IntuneConfiguration

    Write-Host "Collecting MFA status..."
    $mfaStatus = Get-MFAStatus

    Write-Host "Collecting license information..."
    $licenseInfo = Get-LicenseInfo

    Write-Host "Collecting license information..."
    $licenseInfo = Get-LicenseInfo

    Write-Host "Collecting Teams Phone configuration..."
    $teamsPhoneConfig = Get-TeamsPhoneConfig

    Write-Host "Collecting registered domains..."
    $registeredDomains = Get-RegisteredDomains

    Write-Host "Collecting public DNS settings..."
    $publicDNSSettings = Get-PublicDNSSettings

    Write-Host "Collecting all Microsoft Teams teams..."
    $allTeams = Get-AllTeams

    Write-Host "Collecting SharePoint sites..."
    $sharePointSites = Get-SharePointSites

    Write-Host "Collecting SharePoint sites with external sharing..."
    $externalSharingSites = Get-ExternalSharingEnabledSites

    Write-Host "Calling Generate-HTMLReport function..."
    Generate-HTMLReport -TenantConfig $tenantConfig `
                        -ExchangeConfig $exchangeConfig `
                        -TeamsConfig $teamsConfig `
                        -EntraIDConfig $entraIDConfig `
                        -AccessConfig $accessConfig `
                        -SecurityPolicies $securityPolicies `
                        -IntuneConfig $intuneConfig `
                        -MFAStatus $mfaStatus `
                        -LicenseInfo $licenseInfo `
                        -TeamsPhoneConfig $teamsPhoneConfig `
                        -RegisteredDomains $registeredDomains `
                        -PublicDNSSettings $publicDNSSettings `
                        -AllTeams $allTeams `
                        -SharePointSites $sharePointSites `
                        -ExternalSharingSites $externalSharingSites

    Write-Host "Generate-HTMLReport function completed"
}
catch {
    Write-Error "An error occurred during the assessment: $_"
    Write-Error "Stack Trace: $($_.ScriptStack)"
}
finally {
    Write-Host "Disconnecting from Microsoft Graph..."
    Disconnect-MgGraph
}

<#
.SYNOPSIS
    Commvault M365 backup onboarding - GUI to create Azure App Registrations for Commvault
    Microsoft 365 backup (Exchange, OneDrive, Teams, SharePoint), including API permissions
    and admin consent.

.DESCRIPTION
    - Connect to / disconnect from a tenant (Microsoft Graph, interactive sign-in)
    - One tab per workload, each with its own App Registration name and permission set
      (per Commvault documentation 11.40):
        Exchange   : application_permissions_for_azure_app_for_exchange_online.html
        OneDrive   : register_azure_app_for_onedrive_for_business_with_azure_ad.html
        Teams      : register_azure_app_for_teams_with_azure_ad.html      (+ Web redirect URI)
        SharePoint : register_azure_app_for_sharepoint_online_with_azure_ad.html (+ certificate)
    - Grants admin consent (app role assignments + oauth2 permission grants)
    - Optional client secret, shared via a one-time link on https://password.previder.com
      (valid 24 hours, can be opened once). The secret is encrypted locally (AES-256-CTR,
      same scheme as the website); only the ciphertext is sent, the key stays in the link.
    - Output: Application (client) ID, Directory (tenant) ID, Object ID, secret link,
      certificate thumbprint

.NOTES
    Requires : PowerShell 5.1 or 7.x on Windows, module Microsoft.Graph.Authentication
               (you will be offered to install it if it is missing).
    Account  : Global Administrator or Privileged Role Administrator
               (needed to grant admin consent for Graph application permissions).
    If an app with the same name already exists, you can choose to update it.
#>
#Requires -Version 5.1

# WinForms requires an STA thread
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

#region Configuration --------------------------------------------------------------------

$AppTitle   = 'Commvault M365 backup onboarding'
$GraphAppId = '00000003-0000-0000-c000-000000000000'   # Microsoft Graph
$ExoAppId   = '00000002-0000-0ff1-ce00-000000000000'   # Office 365 Exchange Online (EWS)
$SpoAppId   = '00000003-0000-0ff1-ce00-000000000000'   # Office 365 SharePoint Online

$PreviderUrl      = 'https://password.previder.com/'
$PreviderValidity = 86400                              # seconds (24 hours = shortest option)

$DocBase = 'https://documentation.commvault.com/11.40/software/'
$RedirectTemplate = 'https://<hostname>/commandcenter/processAzureAuthToken.do'

function New-Perm {
    # Required = shown as required; Checked = default state (defaults to Required)
    param([string]$Api, [string]$Type, [string]$Value, [bool]$Required = $true, [string]$Note = '', $Checked = $null)
    $resId = switch ($Api) { 'Graph' { $GraphAppId } 'EXO' { $ExoAppId } 'SPO' { $SpoAppId } }
    [pscustomobject]@{
        Api           = $Api
        ResourceAppId = $resId
        Type          = $Type
        Value         = $Value
        Required      = $Required
        Checked       = $(if ($null -eq $Checked) { $Required } else { [bool]$Checked })
        Note          = $Note
    }
}

$Workloads = [ordered]@{
    Exchange   = [pscustomobject]@{
        Key = 'Exchange'; Title = 'Exchange'
        DocUrl = $DocBase + 'application_permissions_for_azure_app_for_exchange_online.html'
        Permissions = @(
            New-Perm 'Graph' 'Application' 'Directory.Read.All'           $true  'Discover all users and user groups'
            New-Perm 'Graph' 'Application' 'Group.ReadWrite.All'          $true  'Discover all user groups'
            New-Perm 'Graph' 'Application' 'Policy.Read.All'              $true  "Read your organization's policies"
            New-Perm 'Graph' 'Application' 'Group.Read.All'               $true  'Discover all groups'
            New-Perm 'Graph' 'Application' 'MailboxSettings.Read'         $true  'Discover all user mailbox settings'
            New-Perm 'Graph' 'Application' 'User.Read.All'                $true  'Discover full profiles of all users'
            New-Perm 'Graph' 'Application' 'MailboxItem.ImportExport.All' $true  'Back up and restore the mailboxes'
            New-Perm 'Graph' 'Application' 'MailboxItem.Read.All'         $false 'Read user mailbox items'
            New-Perm 'Graph' 'Application' 'MailboxFolder.Read.All'       $false 'Read user mailbox folders'
            New-Perm 'Graph' 'Application' 'Application.ReadWrite.All'    $false 'Metallic only (reply URL / secret auto-creation)'
            New-Perm 'Graph' 'Delegated'   'Directory.AccessAsUser.All'   $false 'Access the directory as the signed-in user'
            New-Perm 'EXO'   'Application' 'full_access_as_app'           $true  'Back up and restore the mailboxes (EWS)'
        )
    }
    OneDrive   = [pscustomobject]@{
        Key = 'OneDrive'; Title = 'OneDrive'
        DocUrl = $DocBase + 'register_azure_app_for_onedrive_for_business_with_azure_ad.html'
        Permissions = @(
            New-Perm 'Graph' 'Application' 'Directory.Read.All'
            New-Perm 'Graph' 'Application' 'Files.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'User.Read.All'
            New-Perm 'Graph' 'Application' 'Notes.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Application.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Reports.Read.All'
        )
    }
    Teams      = [pscustomobject]@{
        Key = 'Teams'; Title = 'Teams'
        DocUrl = $DocBase + 'register_azure_app_for_teams_with_azure_ad.html'
        Permissions = @(
            New-Perm 'Graph' 'Application' 'Application.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Channel.Create'
            New-Perm 'Graph' 'Application' 'Channel.ReadBasic.All'
            New-Perm 'Graph' 'Application' 'ChannelMember.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'ChannelMessage.Read.All'
            New-Perm 'Graph' 'Application' 'ChannelSettings.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Chat.Read.All'
            New-Perm 'Graph' 'Application' 'Directory.Read.All'
            New-Perm 'Graph' 'Application' 'Files.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Group.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Notes.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Reports.Read.All'
            New-Perm 'Graph' 'Application' 'Sites.FullControl.All'
            New-Perm 'Graph' 'Application' 'Team.ReadBasic.All'
            New-Perm 'Graph' 'Application' 'TeamsAppInstallation.ReadWriteForTeam.All'
            New-Perm 'Graph' 'Application' 'TeamMember.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'TeamworkTag.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Tasks.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'User.Read.All'
            New-Perm 'Graph' 'Delegated'   'ChannelMessage.Read.All'
            New-Perm 'Graph' 'Delegated'   'ChannelMessage.Send'
            New-Perm 'Graph' 'Delegated'   'Directory.AccessAsUser.All'
            New-Perm 'Graph' 'Delegated'   'Group.ReadWrite.All'
            New-Perm 'Graph' 'Delegated'   'Notes.ReadWrite.All'        $false 'Only needed to restore OneNote' $true
            New-Perm 'Graph' 'Delegated'   'offline_access'
            New-Perm 'Graph' 'Delegated'   'openid'
            New-Perm 'EXO'   'Application' 'full_access_as_app'
        )
    }
    SharePoint = [pscustomobject]@{
        Key = 'SharePoint'; Title = 'SharePoint'
        DocUrl = $DocBase + 'register_azure_app_for_sharepoint_online_with_azure_ad.html'
        Permissions = @(
            New-Perm 'SPO'   'Application' 'Sites.FullControl.All'
            New-Perm 'Graph' 'Application' 'Sites.FullControl.All'
            New-Perm 'Graph' 'Application' 'Group.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Directory.Read.All'
            New-Perm 'Graph' 'Application' 'Application.ReadWrite.All'
            New-Perm 'Graph' 'Application' 'Reports.Read.All'           $false 'Only needed for reports'
        )
    }
}

# Delegated scopes for the admin running this script
$ConnectScopes = @(
    'User.Read'
    'Application.ReadWrite.All'
    'AppRoleAssignment.ReadWrite.All'
    'DelegatedPermissionGrant.ReadWrite.All'
)

$script:TenantName = $null
$script:Account    = $null
$script:Results    = @{}   # per workload: last result (for future extensions, export etc.)
$Tabs              = @{}   # per workload: controls + state

#endregion

#region Helpers --------------------------------------------------------------------------

function Write-Log {
    param([string]$Message, [ValidateSet('Info', 'Warn', 'Error', 'Success')][string]$Level = 'Info')
    $color = switch ($Level) {
        'Error'   { [System.Drawing.Color]::Firebrick }
        'Warn'    { [System.Drawing.Color]::DarkOrange }
        'Success' { [System.Drawing.Color]::ForestGreen }
        default   { [System.Drawing.Color]::Black }
    }
    $txtLog.SelectionStart  = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor  = $color
    $txtLog.AppendText(("[{0}] {1}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $Message))
    $txtLog.SelectionColor  = $txtLog.ForeColor
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Wait-Ui {
    param([int]$Seconds)
    $until = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $until) {
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Get-ErrText {
    param($Err)
    $msg = $Err.Exception.Message
    if ($Err.ErrorDetails -and $Err.ErrorDetails.Message) {
        try {
            $j = $Err.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop
            if ($j.error.message) { $msg = "$($j.error.code): $($j.error.message)" }
        } catch { $msg = $Err.ErrorDetails.Message }
    }
    return $msg
}

function Invoke-Graph {
    param(
        [string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Uri,
        $Body
    )
    $p = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($null -ne $Body) {
        $p.Body        = ($Body | ConvertTo-Json -Depth 10)
        $p.ContentType = 'application/json'
    }
    Invoke-MgGraphRequest @p
}

function Invoke-WithRetry {
    # New apps/service principals sometimes need a moment to replicate
    param([scriptblock]$Action, [string]$What, [int]$Tries = 6, [int]$DelaySec = 5)
    for ($i = 1; $i -le $Tries; $i++) {
        try {
            return (& $Action)
        } catch {
            if ($i -eq $Tries) { throw }
            Write-Log "$What failed (attempt $i/$Tries), retrying in $DelaySec s: $(Get-ErrText $_)" 'Warn'
            Wait-Ui $DelaySec
        }
    }
}

function Add-Access {
    param($Map, [string]$ResourceAppId, [string]$Id, [string]$Type)
    if (-not $ResourceAppId -or -not $Id) { return }
    if (-not $Map.Contains($ResourceAppId)) {
        $Map[$ResourceAppId] = New-Object System.Collections.Generic.List[object]
    }
    $dupe = $Map[$ResourceAppId] | Where-Object { $_.id -eq $Id -and $_.type -eq $Type }
    if (-not $dupe) { $Map[$ResourceAppId].Add(@{ id = $Id; type = $Type }) }
}

function Test-Connected {
    if (-not (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) { return $false }
    return [bool](Get-MgContext)
}

function Initialize-GraphModule {
    if (Get-Module Microsoft.Graph.Authentication) { return $true }
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "The module Microsoft.Graph.Authentication is not installed.`r`nInstall it now for the current user?",
            'Module missing', 'YesNo', 'Question')
        if ($answer -ne 'Yes') { return $false }
        Write-Log 'Installing Microsoft.Graph.Authentication (CurrentUser)...'
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Write-Log "Microsoft.Graph.Authentication $((Get-Module Microsoft.Graph.Authentication).Version) loaded."
    return $true
}

#endregion

#region Previder one-time secret link ----------------------------------------------------

function ConvertTo-PreviderCipher {
    # AES-256-CTR, initial counter 5 - identical to aes-js "new aesjs.Counter(5)" used by the site
    param([byte[]]$Key, [string]$PlainText)
    $data = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Mode    = [System.Security.Cryptography.CipherMode]::ECB
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::None
    $aes.Key     = $Key
    $enc = $aes.CreateEncryptor()
    $counter = New-Object byte[] 16
    $counter[15] = 5
    $out = New-Object byte[] $data.Length
    $ks  = New-Object byte[] 16
    for ($i = 0; $i -lt $data.Length; $i += 16) {
        [void]$enc.TransformBlock($counter, 0, 16, $ks, 0)
        for ($j = 0; $j -lt 16 -and ($i + $j) -lt $data.Length; $j++) {
            $out[$i + $j] = [byte]($data[$i + $j] -bxor $ks[$j])
        }
        for ($k = 15; $k -ge 0; $k--) {
            if ($counter[$k] -eq 255) { $counter[$k] = 0 } else { $counter[$k]++; break }
        }
    }
    $enc.Dispose(); $aes.Dispose()
    return (($out | ForEach-Object { $_.ToString('x2') }) -join '')
}

function New-PreviderLink {
    # Returns @{ Url; Expires }. Only the ciphertext is sent; the key stays in the URL fragment.
    param([string]$Secret, [int]$ValiditySeconds = $PreviderValidity)
    $key = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($key); $rng.Dispose()
    $keyHex = ($key | ForEach-Object { $_.ToString('x2') }) -join ''
    $cipher = ConvertTo-PreviderCipher -Key $key -PlainText $Secret

    $resp = Invoke-RestMethod -Method Post -Uri ($PreviderUrl + 'password/') -ErrorAction Stop -Body @{
        linkValidity      = $ValiditySeconds
        encryptedPassword = $cipher
    }
    if (-not $resp.id) { throw 'password.previder.com returned no id.' }
    $expires = $null
    try {
        $e = $resp.expireDate
        $dt = if ("$e" -match '^\d+$') { [DateTimeOffset]::FromUnixTimeMilliseconds([int64]$e).LocalDateTime } else { ([datetime]$e).ToLocalTime() }
        $expires = $dt.ToString('yyyy-MM-dd HH:mm')
    } catch { $expires = "$($resp.expireDate)" }
    return [pscustomobject]@{ Url = "$PreviderUrl#$($resp.id)/$keyHex"; Expires = $expires }
}

#endregion

#region Certificate helpers --------------------------------------------------------------

function New-DialogForm($Title, $W, $H) {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title; $f.ClientSize = New-Object System.Drawing.Size($W, $H)
    $f.StartPosition = 'CenterParent'; $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false; $f.Font = $fontMain
    $f.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $f.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    return $f
}

function Show-PasswordDialog {
    param([string]$Title, [string]$Prompt)
    $f = New-DialogForm $Title 380 120
    $lbl = New-Label $Prompt 12 12 350
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(12, 38); $tb.Size = New-Object System.Drawing.Size(355, 23)
    $tb.UseSystemPasswordChar = $true
    $ok = New-Button 'OK' 196 80 80 27; $ok.DialogResult = 'OK'
    $cancel = New-Button 'Cancel' 287 80 80 27; $cancel.DialogResult = 'Cancel'
    $f.AcceptButton = $ok; $f.CancelButton = $cancel
    $f.Controls.AddRange(@($lbl, $tb, $ok, $cancel))
    $r = $f.ShowDialog($form)
    $val = $tb.Text
    $f.Dispose()
    if ($r -eq 'OK') { return $val } else { return $null }
}

function Show-CertGenerateDialog {
    param([string]$DefaultSubject)
    $f = New-DialogForm 'Generate self-signed certificate' 430 200
    $tbSubj = New-Object System.Windows.Forms.TextBox
    $tbSubj.Location = New-Object System.Drawing.Point(130, 12); $tbSubj.Size = New-Object System.Drawing.Size(285, 23); $tbSubj.Text = $DefaultSubject
    $numY = New-Object System.Windows.Forms.NumericUpDown
    $numY.Location = New-Object System.Drawing.Point(130, 42); $numY.Size = New-Object System.Drawing.Size(55, 23)
    $numY.Minimum = 1; $numY.Maximum = 10; $numY.Value = 2
    $tbPw1 = New-Object System.Windows.Forms.TextBox
    $tbPw1.Location = New-Object System.Drawing.Point(130, 72); $tbPw1.Size = New-Object System.Drawing.Size(285, 23); $tbPw1.UseSystemPasswordChar = $true
    $tbPw2 = New-Object System.Windows.Forms.TextBox
    $tbPw2.Location = New-Object System.Drawing.Point(130, 102); $tbPw2.Size = New-Object System.Drawing.Size(285, 23); $tbPw2.UseSystemPasswordChar = $true
    $ok = New-Button 'OK' 244 160 80 27
    $cancel = New-Button 'Cancel' 335 160 80 27; $cancel.DialogResult = 'Cancel'
    $ok.Add_Click({
            if ($tbSubj.Text.Trim() -notmatch '^CN=.+') { [System.Windows.Forms.MessageBox]::Show('Subject must start with CN=', 'Invalid subject') | Out-Null; return }
            if (-not $tbPw1.Text) { [System.Windows.Forms.MessageBox]::Show('Please enter a PFX password.', 'Password missing') | Out-Null; return }
            if ($tbPw1.Text -ne $tbPw2.Text) { [System.Windows.Forms.MessageBox]::Show('Passwords do not match.', 'Password mismatch') | Out-Null; return }
            $f.DialogResult = 'OK'
        })
    $f.AcceptButton = $ok; $f.CancelButton = $cancel
    $f.Controls.AddRange(@((New-Label 'Subject:' 12 15 110), $tbSubj, (New-Label 'Valid (years):' 12 45 110), $numY,
            (New-Label 'PFX password:' 12 75 110), $tbPw1, (New-Label 'Confirm:' 12 105 110), $tbPw2, $ok, $cancel))
    $r = $f.ShowDialog($form)
    $res = [pscustomobject]@{ Subject = $tbSubj.Text.Trim(); Years = [int]$numY.Value; Password = $tbPw1.Text }
    $f.Dispose()
    if ($r -eq 'OK') { return $res } else { return $null }
}

function Set-SelectedCert {
    param($t, $Cert, $PfxPath)
    $t.Cert = $Cert
    $t.PfxPath = $PfxPath
    if ($Cert) {
        $t.TxtCert.Text = '{0} | {1} | valid until {2}' -f $Cert.Subject, $Cert.Thumbprint, $Cert.NotAfter.ToString('yyyy-MM-dd')
        Write-Log "Certificate selected: $($Cert.Subject) ($($Cert.Thumbprint))."
        if ($Cert.NotAfter -lt (Get-Date)) { Write-Log 'Warning: this certificate has expired.' 'Warn' }
    } else {
        $t.TxtCert.Text = ''
    }
}

function Select-CertificateFile {
    param($t)
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title  = 'Select certificate'
    $ofd.Filter = 'Certificate files (*.cer;*.crt;*.pfx;*.p12)|*.cer;*.crt;*.pfx;*.p12|All files (*.*)|*.*'
    if ($ofd.ShowDialog($form) -ne 'OK') { return }
    $path = $ofd.FileName
    if ($path -match '\.(pfx|p12)$') {
        $pw = Show-PasswordDialog 'PFX password' "Password for $([System.IO.Path]::GetFileName($path)):"
        if ($null -eq $pw) { return }
        $full = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($path, $pw)
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, $full.RawData)
        $full.Reset()
        Set-SelectedCert $t $cert $path
    } else {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($path)
        Set-SelectedCert $t $cert $null
        Write-Log 'Note: Commvault needs the matching .pfx (with private key) and its password.' 'Warn'
    }
}

function New-SelfSignedCertFiles {
    param($t)
    $default = if ($t.TxtName.Text.Trim()) { 'CN=' + $t.TxtName.Text.Trim() } else { 'CN=Commvault M365 backup SharePoint' }
    $d = Show-CertGenerateDialog $default
    if (-not $d) { return }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title    = 'Save PFX (the .cer is saved next to it)'
    $sfd.Filter   = 'PFX file (*.pfx)|*.pfx'
    $sfd.FileName = (($d.Subject -replace '^CN=', '') -replace '[^\w\-\. ]', '_') + '.pfx'
    if ($sfd.ShowDialog($form) -ne 'OK') { return }
    $pfxPath = $sfd.FileName
    $cerPath = [System.IO.Path]::ChangeExtension($pfxPath, '.cer')

    Write-Log "Generating self-signed certificate $($d.Subject) ($($d.Years) year(s))..."
    $tmp = New-SelfSignedCertificate -Subject $d.Subject -CertStoreLocation 'Cert:\CurrentUser\My' `
        -KeyExportPolicy Exportable -KeySpec Signature -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears($d.Years) -ErrorAction Stop
    try {
        $sec = ConvertTo-SecureString $d.Password -AsPlainText -Force
        Export-PfxCertificate -Cert $tmp -FilePath $pfxPath -Password $sec -ErrorAction Stop | Out-Null
        Export-Certificate -Cert $tmp -FilePath $cerPath -Type CERT -ErrorAction Stop | Out-Null
    } finally {
        # Private key only lives in the PFX; remove it from the local store
        try { Remove-Item -Path "Cert:\CurrentUser\My\$($tmp.Thumbprint)" -DeleteKey -ErrorAction Stop } catch {
            Remove-Item -Path "Cert:\CurrentUser\My\$($tmp.Thumbprint)" -ErrorAction SilentlyContinue
        }
    }
    Write-Log "Saved: $pfxPath (+ .cer). Upload the .pfx and its password in Commvault Command Center." 'Success'
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cerPath)
    Set-SelectedCert $t $cert $pfxPath
}

function Get-ThumbprintBase64([string]$Thumbprint) {
    $bytes = [byte[]]@(for ($i = 0; $i -lt $Thumbprint.Length; $i += 2) { [Convert]::ToByte($Thumbprint.Substring($i, 2), 16) })
    return [Convert]::ToBase64String($bytes)
}

#endregion

#region UI state -------------------------------------------------------------------------

function Get-CurrentTab { return $Tabs[$tabControl.SelectedTab.Tag] }

function Update-UiState {
    $c = Test-Connected
    $btnConnect.Enabled    = -not $c
    $txtTenant.Enabled     = -not $c
    $btnDisconnect.Enabled = $c
    $btnCreate.Enabled     = $c
    if ($c) {
        $ctx = Get-MgContext
        $name = if ($script:TenantName) { $script:TenantName } else { 'tenant' }
        $lblStatus.Text      = "Connected to $name ($($ctx.TenantId)) as $($script:Account)"
        $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
    } else {
        $lblStatus.Text      = 'Not connected'
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
    }
}

function Set-Busy {
    param([bool]$Busy)
    foreach ($b in @($btnConnect, $btnDisconnect, $btnCreate)) { $b.Enabled = -not $Busy }
    $tabControl.Enabled = -not $Busy
    $form.UseWaitCursor = $Busy
    [System.Windows.Forms.Application]::DoEvents()
    if (-not $Busy) { Update-UiState }
}

function Show-Result {
    param([string]$Key)
    $r = $script:Results[$Key]
    $grpOut.Text = "3. Result - $($Workloads[$Key].Title)"
    $txtOutAppId.Text      = if ($r) { $r.ApplicationId } else { '' }
    $txtOutTenantId.Text   = if ($r) { $r.DirectoryId } else { '' }
    $txtOutObjectId.Text   = if ($r) { $r.ObjectId } else { '' }
    $txtOutSecret.Text     = if ($r) { if ($r.SecretLink) { $r.SecretLink } else { $r.ClientSecret } } else { '' }
    $txtOutSecretEnd.Text  = if ($r) { $r.SecretInfo } else { '' }
    $txtOutThumb.Text      = if ($r) { $r.CertificateThumbprint } else { '' }
    $lblOutSecret.Text     = if ($r -and -not $r.SecretLink -and $r.ClientSecret) { 'Client secret:' } else { 'Secret link:' }
}

#endregion

#region GUI ------------------------------------------------------------------------------

$fontMain = New-Object System.Drawing.Font('Segoe UI', 9)
$fontMono = New-Object System.Drawing.Font('Consolas', 9)

$form = New-Object System.Windows.Forms.Form
$form.Text            = $AppTitle
# Layout is designed at 96 DPI; WinForms scales it to the actual DPI (125%, 150%, ...)
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$form.AutoScaleMode       = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.ClientSize      = New-Object System.Drawing.Size(1255, 518)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox     = $false
$form.AutoScroll      = $true
$form.Font            = $fontMain

function New-Label($Text, $X, $Y, $W = 170) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.Location = New-Object System.Drawing.Point($X, $Y); $l.Size = New-Object System.Drawing.Size($W, 20)
    return $l
}
function New-Button($Text, $X, $Y, $W = 130, $H = 28) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Location = New-Object System.Drawing.Point($X, $Y); $b.Size = New-Object System.Drawing.Size($W, $H)
    return $b
}
function New-TextBox($X, $Y, $W) {
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point($X, $Y); $tb.Size = New-Object System.Drawing.Size($W, 23)
    return $tb
}

# --- 1. Connection ---
$grpConn = New-Object System.Windows.Forms.GroupBox
$grpConn.Text = '1. Tenant connection'
$grpConn.Location = New-Object System.Drawing.Point(10, 10)
$grpConn.Size = New-Object System.Drawing.Size(765, 92)

$txtTenant     = New-TextBox 170 25 285
$btnConnect    = New-Button 'Connect' 470 23
$btnDisconnect = New-Button 'Disconnect' 610 23

$lblTenantHint = New-Label 'Domain (e.g. customer.onmicrosoft.com) or tenant ID. Empty = home tenant of your account.' 170 50 580
$lblTenantHint.ForeColor = [System.Drawing.Color]::Gray

$lblStatus = New-Label 'Not connected' 15 68 735
$lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$grpConn.Controls.AddRange(@((New-Label 'Tenant (optional):' 15 28 150), $txtTenant, $btnConnect, $btnDisconnect, $lblTenantHint, $lblStatus))

# --- 2. App Registration (tabs) ---
$grpApp = New-Object System.Windows.Forms.GroupBox
$grpApp.Text = '2. App Registration'
$grpApp.Location = New-Object System.Drawing.Point(10, 108)
$grpApp.Size = New-Object System.Drawing.Size(765, 400)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 22)
$tabControl.Size = New-Object System.Drawing.Size(745, 330)

function New-WorkloadTab {
    param($W)
    $page = New-Object System.Windows.Forms.TabPage
    $page.Text = $W.Title
    $page.Tag  = $W.Key
    $page.UseVisualStyleBackColor = $true
    $page.AutoScroll = $true

    $t = @{ Key = $W.Key; Workload = $W; Page = $page; Cert = $null; PfxPath = $null; TxtRedirect = $null; TxtCert = $null }

    $t.TxtName = New-TextBox 160 10 345
    $t.TxtName.MaxLength = 120

    $link = New-Object System.Windows.Forms.LinkLabel
    $link.Text = 'Commvault documentation'
    $link.Location = New-Object System.Drawing.Point(560, 13); $link.Size = New-Object System.Drawing.Size(165, 20)
    $link.TextAlign = 'TopRight'
    $link.Tag = $W.DocUrl
    $link.Add_LinkClicked({ param($s, $e) Start-Process $s.Tag })

    $t.ChkSecret = New-Object System.Windows.Forms.CheckBox
    $t.ChkSecret.Text = 'Create client secret, valid for'
    $t.ChkSecret.Location = New-Object System.Drawing.Point(10, 41)
    $t.ChkSecret.Size = New-Object System.Drawing.Size(200, 22)
    $t.ChkSecret.Checked = $true

    $t.NumMonths = New-Object System.Windows.Forms.NumericUpDown
    $t.NumMonths.Location = New-Object System.Drawing.Point(210, 41)
    $t.NumMonths.Size = New-Object System.Drawing.Size(55, 23)
    $t.NumMonths.Minimum = 1; $t.NumMonths.Maximum = 24; $t.NumMonths.Value = 24

    $t.ChkShare = New-Object System.Windows.Forms.CheckBox
    $t.ChkShare.Text = 'Share via password.previder.com (24h, one-time)'
    $t.ChkShare.Location = New-Object System.Drawing.Point(345, 41)
    $t.ChkShare.Size = New-Object System.Drawing.Size(380, 22)
    $t.ChkShare.Checked = $true

    $numRef = $t.NumMonths; $shareRef = $t.ChkShare
    $t.ChkSecret.Tag = @($numRef, $shareRef)
    $t.ChkSecret.Add_CheckedChanged({ param($s, $e) foreach ($c in $s.Tag) { $c.Enabled = $s.Checked } })

    $t.Clb = New-Object System.Windows.Forms.CheckedListBox
    $t.Clb.Location = New-Object System.Drawing.Point(10, 92)
    $t.Clb.Size = New-Object System.Drawing.Size(715, 150)
    $t.Clb.CheckOnClick = $true
    $t.Clb.HorizontalScrollbar = $true
    $t.Clb.Font = $fontMono
    foreach ($p in $W.Permissions) {
        $req  = if ($p.Required) { 'required' } else { 'optional' }
        $text = '{0,-5} {1,-11} {2,-41} {3,-8}' -f $p.Api, $p.Type, $p.Value, $req
        if ($p.Note) { $text += " - $($p.Note)" }
        [void]$t.Clb.Items.Add($text, $p.Checked)
    }

    $page.Controls.AddRange(@((New-Label 'App Registration name:' 10 13 150), $t.TxtName, $link,
            $t.ChkSecret, $t.NumMonths, (New-Label 'months' 270 44 60), $t.ChkShare,
            (New-Label 'API permissions:' 10 72 400), $t.Clb))

    if ($W.Key -eq 'Teams') {
        $t.TxtRedirect = New-TextBox 160 252 565
        $t.TxtRedirect.Text = $RedirectTemplate
        $hint = New-Label 'Web platform. Replace <hostname> with the Command Center host. Separate multiple URIs with ;' 160 278 565
        $hint.ForeColor = [System.Drawing.Color]::Gray
        $page.Controls.AddRange(@((New-Label 'Redirect URI (Web):' 10 255 150), $t.TxtRedirect, $hint))
    }

    if ($W.Key -eq 'SharePoint') {
        $t.TxtCert = New-TextBox 160 252 360
        $t.TxtCert.ReadOnly = $true
        $btnBrowse = New-Button 'Browse...' 527 251 95 25
        $btnGen    = New-Button 'Generate...' 630 251 95 25
        $btnClr    = New-Button 'Clear' 630 277 95 23
        foreach ($b in @($btnBrowse, $btnGen, $btnClr)) { $b.Tag = $W.Key }
        $btnBrowse.Add_Click({ param($s, $e) try { Select-CertificateFile $Tabs[$s.Tag] } catch { Write-Log "Certificate: $($_.Exception.Message)" 'Error' } })
        $btnGen.Add_Click({ param($s, $e) try { New-SelfSignedCertFiles $Tabs[$s.Tag] } catch { Write-Log "Certificate: $($_.Exception.Message)" 'Error' } })
        $btnClr.Add_Click({ param($s, $e) Set-SelectedCert $Tabs[$s.Tag] $null $null })
        $hint = New-Label 'Public key (.cer) is uploaded to the app. Commvault needs the .pfx + password.' 160 279 465
        $hint.ForeColor = [System.Drawing.Color]::Gray
        $page.Controls.AddRange(@((New-Label 'Certificate (optional):' 10 255 150), $t.TxtCert, $btnBrowse, $btnGen, $btnClr, $hint))
    }

    return $t
}

foreach ($W in $Workloads.Values) {
    $t = New-WorkloadTab $W
    $Tabs[$W.Key] = $t
    [void]$tabControl.TabPages.Add($t.Page)
}

$btnCreate = New-Button 'Create App Registration + consent' 10 360 300 30
$grpApp.Controls.AddRange(@($tabControl, $btnCreate))

# --- 3. Result ---
$grpOut = New-Object System.Windows.Forms.GroupBox
$grpOut.Text = '3. Result'
$grpOut.Location = New-Object System.Drawing.Point(785, 10)
$grpOut.Size = New-Object System.Drawing.Size(460, 235)

function New-OutputRow($Caption, $Y) {
    $tb = New-TextBox 135 $Y 255
    $tb.ReadOnly = $true
    $tb.Font = $fontMono
    $btn = New-Button 'Copy' 395 ($Y - 1) 52 25
    $btn.Tag = $tb
    $btn.Add_Click({
            param($s, $e)
            if ($s.Tag.Text) {
                [System.Windows.Forms.Clipboard]::SetText($s.Tag.Text)
                Write-Log 'Copied to clipboard.'
            }
        })
    $lbl = New-Label $Caption 12 ($Y + 3) 122
    $grpOut.Controls.AddRange(@($lbl, $tb, $btn))
    return @($tb, $lbl)
}

$txtOutAppId, $null            = New-OutputRow 'Application ID:'   25
$txtOutTenantId, $null         = New-OutputRow 'Directory ID:'     55
$txtOutObjectId, $null         = New-OutputRow 'Object ID:'        85
$txtOutSecret, $lblOutSecret   = New-OutputRow 'Secret link:'      115
$txtOutSecretEnd, $null        = New-OutputRow 'Secret expires:'   145
$txtOutThumb, $null            = New-OutputRow 'Cert thumbprint:'  175

$btnCopyAll = New-Button 'Copy all' 287 203 160 25
$grpOut.Controls.Add($btnCopyAll)

# --- Log ---
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(785, 272)
$txtLog.Size = New-Object System.Drawing.Size(460, 236)
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Font = $fontMono

$form.Controls.AddRange(@($grpConn, $grpApp, $grpOut, (New-Label 'Log:' 787 252), $txtLog))

#endregion

#region Create App Registration ----------------------------------------------------------

function Invoke-CreateApp {
    param($t)
    $W    = $t.Workload
    $name = $t.TxtName.Text.Trim()
    if (-not $name) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a name for the $($W.Title) App Registration first.", 'Name missing', 'OK', 'Warning') | Out-Null
        [void]$t.TxtName.Focus(); return
    }
    $selected = @(for ($i = 0; $i -lt $t.Clb.Items.Count; $i++) { if ($t.Clb.GetItemChecked($i)) { $W.Permissions[$i] } })
    if ($selected.Count -eq 0) { Write-Log 'No permissions selected.' 'Warn'; return }

    # Redirect URIs (Teams)
    $redirects = @()
    if ($t.TxtRedirect) {
        $rawRedirect = $t.TxtRedirect.Text.Trim()
        if ($rawRedirect -eq $RedirectTemplate) { $rawRedirect = '' }
        foreach ($u in @($rawRedirect -split '[;,\s]+' | Where-Object { $_ })) {
            [uri]$parsed = $null
            $valid = ($u -notmatch '[<>]') -and [uri]::TryCreate($u, [UriKind]::Absolute, [ref]$parsed) -and
                     ($parsed.Scheme -eq 'https' -or $parsed.Host -eq 'localhost')
            if (-not $valid) {
                [System.Windows.Forms.MessageBox]::Show("Invalid redirect URI:`r`n$u`r`n`r`nExpected e.g. https://commvault.customer.com/commandcenter/processAzureAuthToken.do",
                    'Invalid redirect URI', 'OK', 'Warning') | Out-Null
                return
            }
            $redirects += $u
        }
        if ($redirects.Count -eq 0) { Write-Log 'No redirect URI entered; you can add it later.' 'Warn' }
    }

    Set-Busy $true
    $ctx = Get-MgContext
    Write-Log "=== Start $($W.Title): '$name' in tenant $($ctx.TenantId) ==="

    # 1. Resolve permissions on the resource service principals
    Write-Log 'Resolving permissions...'
    $resSps = @{}
    foreach ($resAppId in @($selected.ResourceAppId | Select-Object -Unique)) {
        $resSps[$resAppId] = Invoke-Graph -Uri "v1.0/servicePrincipals(appId='$resAppId')?`$select=id,appId,displayName,appRoles,oauth2PermissionScopes"
    }
    $resolved = @(foreach ($p in $selected) {
            $rsp = $resSps[$p.ResourceAppId]
            if ($p.Type -eq 'Application') {
                $def = @($rsp.appRoles) | Where-Object { $_.value -eq $p.Value } | Select-Object -First 1
                $accessType = 'Role'
            } else {
                $def = @($rsp.oauth2PermissionScopes) | Where-Object { $_.value -eq $p.Value } | Select-Object -First 1
                $accessType = 'Scope'
            }
            if (-not $def) { Write-Log "  ! $($p.Type) permission '$($p.Value)' not found on $($rsp.displayName) - skipped" 'Warn'; continue }
            [pscustomobject]@{ Perm = $p; Id = $def.id; AccessType = $accessType; ResourceSpId = $rsp.id; ResourceName = $rsp.displayName }
        })

    # 2. Create App Registration or update the existing one
    $filter   = [uri]::EscapeDataString("displayName eq '$($name.Replace("'", "''"))'")
    $existing = @((Invoke-Graph -Uri "v1.0/applications?`$filter=$filter&`$select=id,appId,displayName,requiredResourceAccess,web,keyCredentials").value | Where-Object { $_ })
    $app = $null
    if ($existing.Count -gt 1) {
        throw "There are already $($existing.Count) apps named '$name'. Please choose a unique name."
    } elseif ($existing.Count -eq 1) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "An App Registration named '$name' already exists (App ID $($existing[0].appId)).`r`n`r`nUpdate this existing app?",
            'App already exists', 'YesNo', 'Question')
        if ($answer -ne 'Yes') { Write-Log 'Cancelled by user.' 'Warn'; return }
        $app = $existing[0]
    }

    $rra = [ordered]@{}
    if ($app) {
        foreach ($r in @($app.requiredResourceAccess | Where-Object { $_ })) {
            foreach ($a in @($r.resourceAccess | Where-Object { $_ })) { Add-Access $rra $r.resourceAppId $a.id $a.type }
        }
    }
    foreach ($r in $resolved) { Add-Access $rra $r.Perm.ResourceAppId $r.Id $r.AccessType }
    $rraBody = @(foreach ($k in $rra.Keys) { @{ resourceAppId = $k; resourceAccess = $rra[$k].ToArray() } })

    $allRedirects = @()
    if ($app -and $app.web -and $app.web.redirectUris) { $allRedirects += @($app.web.redirectUris) }
    $allRedirects = [string[]]@($allRedirects + $redirects | Where-Object { $_ } | Select-Object -Unique)

    if ($app) {
        $patch = @{ requiredResourceAccess = $rraBody }
        if ($redirects.Count) { $patch.web = @{ redirectUris = $allRedirects } }
        Invoke-Graph -Method PATCH -Uri "v1.0/applications/$($app.id)" -Body $patch | Out-Null
        Write-Log "Existing App Registration updated (App ID $($app.appId))." 'Success'
    } else {
        $body = @{
            displayName            = $name
            signInAudience         = 'AzureADMyOrg'
            requiredResourceAccess = $rraBody
        }
        if ($redirects.Count) { $body.web = @{ redirectUris = $allRedirects } }
        $app = Invoke-Graph -Method POST -Uri 'v1.0/applications' -Body $body
        Write-Log "App Registration created (App ID $($app.appId))." 'Success'
    }
    if ($redirects.Count) { Write-Log "  Redirect URI(s): $($allRedirects -join ', ')" 'Success' }

    # 3. Certificate (SharePoint)
    $thumb = $null
    if ($t.Cert) {
        $thumb    = $t.Cert.Thumbprint
        $thumbB64 = Get-ThumbprintBase64 $thumb
        $keys     = @($app.keyCredentials | Where-Object { $_ })
        $already  = $keys | Where-Object { $_.customKeyIdentifier -eq $thumbB64 -or $_.customKeyIdentifier -eq $thumb }
        $upload   = $true
        if ($already) {
            Write-Log "  = Certificate $thumb already present on the app."
            $upload = $false
        } elseif ($keys.Count -gt 0) {
            $answer = [System.Windows.Forms.MessageBox]::Show(
                "The app already has $($keys.Count) certificate(s). Uploading via Graph replaces the existing certificate(s).`r`n`r`nReplace them with the selected certificate?",
                'Existing certificates', 'YesNo', 'Warning')
            if ($answer -ne 'Yes') { Write-Log '  Certificate upload skipped.' 'Warn'; $upload = $false }
        }
        if ($upload) {
            Invoke-Graph -Method PATCH -Uri "v1.0/applications/$($app.id)" -Body @{
                keyCredentials = @(@{
                        type        = 'AsymmetricX509Cert'
                        usage       = 'Verify'
                        key         = [Convert]::ToBase64String($t.Cert.RawData)
                        displayName = $t.Cert.Subject
                    })
            } | Out-Null
            Write-Log "  + Certificate uploaded ($thumb, valid until $($t.Cert.NotAfter.ToString('yyyy-MM-dd')))." 'Success'
        }
    }

    # 4. Service principal (Enterprise application)
    $spFilter = [uri]::EscapeDataString("appId eq '$($app.appId)'")
    $sp = @((Invoke-Graph -Uri "v1.0/servicePrincipals?`$filter=$spFilter&`$select=id,appId").value | Where-Object { $_ }) | Select-Object -First 1
    if (-not $sp) {
        $sp = Invoke-WithRetry -What 'Creating service principal' -Action {
            Invoke-Graph -Method POST -Uri 'v1.0/servicePrincipals' -Body @{
                appId = $app.appId
                tags  = @('WindowsAzureActiveDirectoryIntegratedApp')
            }
        }
        Write-Log 'Service principal (Enterprise application) created.' 'Success'
    } else {
        Write-Log 'Service principal already exists.'
    }

    # 5. Admin consent - application permissions
    Write-Log 'Granting admin consent for application permissions...'
    $assigned = @((Invoke-WithRetry -What 'Retrieving existing assignments' -Action {
                Invoke-Graph -Uri "v1.0/servicePrincipals/$($sp.id)/appRoleAssignments"
            }).value | Where-Object { $_ })
    foreach ($r in @($resolved | Where-Object { $_.AccessType -eq 'Role' })) {
        if ($assigned | Where-Object { $_.appRoleId -eq $r.Id -and $_.resourceId -eq $r.ResourceSpId }) {
            Write-Log "  = $($r.Perm.Value) ($($r.ResourceName)) already granted"
            continue
        }
        Invoke-WithRetry -What "Consent $($r.Perm.Value)" -Action {
            Invoke-Graph -Method POST -Uri "v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" -Body @{
                principalId = $sp.id
                resourceId  = $r.ResourceSpId
                appRoleId   = $r.Id
            }
        } | Out-Null
        Write-Log "  + $($r.Perm.Value) ($($r.ResourceName))" 'Success'
    }

    # 6. Admin consent - delegated permissions (tenant-wide)
    foreach ($g in @($resolved | Where-Object { $_.AccessType -eq 'Scope' } | Group-Object ResourceSpId)) {
        $resId  = $g.Name
        $scopes = @($g.Group | ForEach-Object { $_.Perm.Value })
        $gf     = [uri]::EscapeDataString("clientId eq '$($sp.id)' and resourceId eq '$resId'")
        $grant  = @((Invoke-Graph -Uri "v1.0/oauth2PermissionGrants?`$filter=$gf").value |
                Where-Object { $_ -and $_.consentType -eq 'AllPrincipals' }) | Select-Object -First 1
        if ($grant) {
            $newScope = (@(($grant.scope -split ' ') + $scopes) | Where-Object { $_ } | Select-Object -Unique) -join ' '
            Invoke-Graph -Method PATCH -Uri "v1.0/oauth2PermissionGrants/$($grant.id)" -Body @{ scope = $newScope } | Out-Null
        } else {
            Invoke-WithRetry -What 'Delegated consent' -Action {
                Invoke-Graph -Method POST -Uri 'v1.0/oauth2PermissionGrants' -Body @{
                    clientId    = $sp.id
                    consentType = 'AllPrincipals'
                    resourceId  = $resId
                    scope       = ($scopes -join ' ')
                }
            } | Out-Null
        }
        Write-Log "  + delegated: $($scopes -join ', ')" 'Success'
    }

    # 7. Client secret (+ one-time link)
    $secretText = $null; $secretLink = $null; $secretInfo = $null
    if ($t.ChkSecret.Checked) {
        $end = (Get-Date).ToUniversalTime().AddMonths([int]$t.NumMonths.Value)
        $pw  = Invoke-Graph -Method POST -Uri "v1.0/applications/$($app.id)/addPassword" -Body @{
            passwordCredential = @{
                displayName = "Commvault M365 backup $($W.Title) $(Get-Date -Format 'yyyy-MM-dd')"
                endDateTime = $end.ToString('yyyy-MM-ddTHH:mm:ssZ')
            }
        }
        $secretEnd  = $end.ToLocalTime().ToString('yyyy-MM-dd')
        $secretInfo = "secret valid until $secretEnd"
        Write-Log "Client secret created (valid until $secretEnd)."

        if ($t.ChkShare.Checked) {
            try {
                $linkObj    = New-PreviderLink -Secret $pw.secretText
                $secretLink = $linkObj.Url
                $secretInfo = "link expires $($linkObj.Expires) (one-time) | $secretInfo"
                Write-Log "One-time secret link created on password.previder.com (expires $($linkObj.Expires))." 'Success'
                Write-Log 'Do not open the link yourself: it can be retrieved only once.' 'Warn'
            } catch {
                # Fallback: show the secret itself, otherwise it is lost
                $secretText = $pw.secretText
                Write-Log "Creating the secret link failed: $(Get-ErrText $_). The secret value is shown instead." 'Error'
            }
        } else {
            $secretText = $pw.secretText
            Write-Log 'Store the client secret now: it cannot be retrieved later!' 'Warn'
        }
    }

    # 8. Result
    $script:Results[$W.Key] = [pscustomobject]@{
        Workload              = $W.Title
        DisplayName           = $name
        ApplicationId         = $app.appId
        DirectoryId           = $ctx.TenantId
        TenantName            = $script:TenantName
        ObjectId              = $app.id
        ServicePrincipalId    = $sp.id
        ClientSecret          = $secretText
        SecretLink            = $secretLink
        SecretInfo            = $secretInfo
        RedirectUris          = $allRedirects
        CertificateThumbprint = $thumb
        PfxPath               = $t.PfxPath
        Permissions           = @($resolved | ForEach-Object { "$($_.Perm.Api)/$($_.Perm.Type)/$($_.Perm.Value)" })
    }
    Show-Result $W.Key
    Write-Log "=== Done: $($W.Title) ===" 'Success'
}

#endregion

#region Events ---------------------------------------------------------------------------

$tabControl.Add_SelectedIndexChanged({ Show-Result $tabControl.SelectedTab.Tag })

$btnConnect.Add_Click({
        Set-Busy $true
        try {
            if (-not (Initialize-GraphModule)) { Write-Log 'Connect cancelled: module missing.' 'Warn'; return }

            $params = @{ Scopes = $ConnectScopes; ContextScope = 'Process'; ErrorAction = 'Stop' }
            if ((Get-Command Connect-MgGraph).Parameters.ContainsKey('NoWelcome')) { $params.NoWelcome = $true }
            $tenant = $txtTenant.Text.Trim()
            if ($tenant) { $params.TenantId = $tenant }

            Write-Log ("Connecting to {0}... (the sign-in window may open behind this window)" -f $(if ($tenant) { $tenant } else { 'home tenant' }))
            try {
                Connect-MgGraph @params | Out-Null
            } catch {
                # WAM sign-in sometimes fails from a GUI; fall back to browser sign-in
                if (Get-Command Set-MgGraphOption -ErrorAction SilentlyContinue) {
                    Write-Log "Sign-in failed ($($_.Exception.Message)). Retrying via browser..." 'Warn'
                    Set-MgGraphOption -DisableLoginByWAM $true
                    Connect-MgGraph @params | Out-Null
                } else { throw }
            }

            $ctx = Get-MgContext
            $missing = @($ConnectScopes | Where-Object { $ctx.Scopes -notcontains $_ })
            if ($missing.Count) { Write-Log "Warning, not all scopes were granted: $($missing -join ', ')" 'Warn' }

            try {
                $org = @((Invoke-Graph -Uri 'v1.0/organization?$select=id,displayName').value)[0]
                $script:TenantName = $org.displayName
            } catch { $script:TenantName = $null }

            $script:Account = $ctx.Account
            if (-not $script:Account) {
                try { $script:Account = (Invoke-Graph -Uri 'v1.0/me?$select=userPrincipalName').userPrincipalName } catch { }
            }
            Write-Log "Connected to $($script:TenantName) ($($ctx.TenantId)) as $($script:Account)." 'Success'
        } catch {
            Write-Log "Connect failed: $(Get-ErrText $_)" 'Error'
        } finally {
            Set-Busy $false
        }
    })

$btnDisconnect.Add_Click({
        Set-Busy $true
        try {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            $script:TenantName = $null
            Write-Log 'Disconnected.' 'Success'
        } catch {
            Write-Log "Disconnect failed: $(Get-ErrText $_)" 'Error'
        } finally {
            Set-Busy $false
        }
    })

$btnCreate.Add_Click({
        try {
            Invoke-CreateApp (Get-CurrentTab)
        } catch {
            $msg = Get-ErrText $_
            Write-Log "Error: $msg" 'Error'
            if ($msg -match 'Authorization_RequestDenied|Insufficient privileges|Forbidden') {
                Write-Log 'Tip: granting admin consent requires Global Administrator or Privileged Role Administrator.' 'Warn'
            }
        } finally {
            Set-Busy $false
        }
    })

$btnCopyAll.Add_Click({
        $r = $script:Results[$tabControl.SelectedTab.Tag]
        if (-not $r) { return }
        $lines = @(
            "Workload                : $($r.Workload)"
            "App Registration        : $($r.DisplayName)"
            "Tenant                  : $($r.TenantName)"
            "Application (client) ID : $($r.ApplicationId)"
            "Directory (tenant) ID   : $($r.DirectoryId)"
            "Object ID               : $($r.ObjectId)"
        )
        if ($r.SecretLink)            { $lines += "Secret link (one-time)  : $($r.SecretLink)" }
        if ($r.ClientSecret)          { $lines += "Client secret           : $($r.ClientSecret)" }
        if ($r.SecretInfo)            { $lines += "Secret                  : $($r.SecretInfo)" }
        if ($r.RedirectUris)          { $lines += "Redirect URI(s)         : $($r.RedirectUris -join ', ')" }
        if ($r.CertificateThumbprint) { $lines += "Certificate thumbprint  : $($r.CertificateThumbprint)" }
        [System.Windows.Forms.Clipboard]::SetText(($lines -join "`r`n"))
        Write-Log 'All details copied to clipboard.'
    })

$form.Add_Shown({
        if (Get-Module -ListAvailable Microsoft.Graph.Authentication) {
            Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
        } else {
            Write-Log 'Microsoft.Graph.Authentication is not installed; you will be offered to install it on Connect.' 'Warn'
        }
        $wa = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
        if ($form.Width -gt $wa.Width)   { $form.Width = $wa.Width;   $form.Left = $wa.Left }
        if ($form.Height -gt $wa.Height) { $form.Height = $wa.Height; $form.Top = $wa.Top }
        Update-UiState
        Show-Result $tabControl.SelectedTab.Tag
        Write-Log 'Ready. Step 1: connect to the tenant.'
        $form.Activate()
    })

$form.Add_FormClosing({
        if (Test-Connected) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    })

#endregion

[void]$form.ShowDialog()
$form.Dispose()
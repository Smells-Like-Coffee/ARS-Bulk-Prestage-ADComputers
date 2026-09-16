<#
.SYNOPSIS
    ARS Bulk Computer Prestage Utility

.DESCRIPTION
    This PowerShell script provides a graphical interface for bulk prestaging
    computer accounts in Active Directory through One Identity Active Roles.

    The utility allows computer names to be imported from a CSV file and
    created in a selected Organizational Unit.

    All directory browsing, validation, computer creation, and computer
    modification operations are performed through One Identity Active Roles.

    The Microsoft ActiveDirectory PowerShell module / RSAT is NOT required.

    The script connects to One Identity Active Roles using the current user's
    Windows credentials through an Active Roles Proxy connection.

    Configuration values are stored in a config.ini file located in the same
    directory as the script.

    Settings include:

        - Active Roles Server
        - Domain Distinguished Name
        - Computer CSV file
        - Destination Organizational Unit
        - Report output folder

    Active Roles / Active Directory attributes are configurable through the
    [Attributes] section of config.ini.

    The following attribute is included by default:

        edsvaCHSServer=FALSE
        edsaJoinComputerToDomain=US\Domain Users

    Additional attributes may be added to the [Attributes] section as needed.

    Example:

        [Attributes]
        edsvaCHSServer=FALSE
        edsaJoinComputerToDomain=US\Domain Users
        description=Bulk Prestaged Computer
        SomeOtherAttribute=TRUE

    TRUE and FALSE values are converted to PowerShell Boolean values.
    Numeric values are converted to numbers when possible.
    Other values are treated as strings.

    Existing computer objects are NOT moved to the selected destination OU.
    If an existing computer is found, the configured attributes are updated.

    The script generates a CSV results report showing whether each computer
    account was created, updated, or failed.

.ACTIVE ROLES DEPENDENCIES
    Requires 64-bit Windows PowerShell 5.1 and three Active Roles components:
    ADSI Provider, SDK, and PowerShell Module (ActiveRolesManagementShell).

    Tools > Active Roles Prerequisites shows each component's status and Install
    button. Use an individual Install button or Install All to install
    only missing components, in ADSI / SDK / Shell order, with one Windows UAC
    elevation or administrator credential prompt for the entire batch.

    Paths are relative to this script's Prerequisites folder:
      ActiveRoles ADSI Provider - For exporting objects and reports\x64\ADSI\_x64.msi
      ActiveRoles ADSI Provider - For exporting objects and reports\x64\SDK\_x64.msi
      ActiveRoles Powershell Module\x64\Shell\_x64.msi

    Installation uses /qn /norestart. Failure stops later components; already
    installed components are not rolled back. Reboot-required results block
    directory operations until Windows has been restarted and the tool reopened.
    Logs and staged MSI media are retained in the displayed temporary folder.
    Credentials are handled by Windows, never collected or saved by this script.

    Preserves the v1.6 SearchRoot / SearchScope Base DN fix and asynchronous
    OU-loading window, including progress, cancellation and loaded OU count.

.USAGE
    1. Launch ARS-Bulk-Prestage-ADComputers.ps1.

    2. Open Tools > Active Roles Prerequisites and use Install All or Install.

    3. If Windows requests administrative credentials, enter an authorized
       administrator account.

    4. Confirm or enter the Active Roles server.

    5. Confirm or enter the Domain Distinguished Name.

    6. Select a CSV containing a ComputerName column.

    7. Select the destination Organizational Unit.

    8. Select the report output directory.

    9. Review configured attributes.

    10. Click Start Import.

.REQUIREMENTS
    - Windows PowerShell
    - One Identity ActiveRolesManagementShell module
    - Network connectivity to the Active Roles server
    - Appropriate Active Roles / Active Directory permissions
    - Read access to the CSV
    - Write access to the report folder

    Microsoft RSAT / ActiveDirectory PowerShell module is NOT required.

.NOTES
    Author:
        CJ Micklitsch

    Created with assistance from:
        Karen (OpenAI Codex)

    Version:
        2.1
#>

# ============================================================
# WINDOWS FORMS
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
# SCRIPT / CONFIG PATH
# ============================================================

$ScriptDirectory = if ($PSScriptRoot) {
    $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -LiteralPath $MyInvocation.MyCommand.Path -Parent
}
else {
    (Get-Location).Path
}

$ConfigPath = Join-Path $ScriptDirectory "config.ini"

# ============================================================
# ACTIVE ROLES PREREQUISITES
# Detection is read-only: never use Win32_Product (it can repair products).
# Keep these three entries in installation order.
# ============================================================
$script:Prerequisites = @(
    [pscustomobject]@{ Id='ADSI'; Name='ADSI Provider'; RelativePath='ActiveRoles ADSI Provider - For exporting objects and reports\x64\ADSI\_x64.msi'; Pattern='(?i)Active\s*Roles.*ADSI.*Provider' },
    [pscustomobject]@{ Id='SDK'; Name='SDK'; RelativePath='ActiveRoles ADSI Provider - For exporting objects and reports\x64\SDK\_x64.msi'; Pattern='(?i)Active\s*Roles.*(SDK|Software Development Kit)' },
    [pscustomobject]@{ Id='Shell'; Name='PowerShell Module'; RelativePath='ActiveRoles Powershell Module\x64\Shell\_x64.msi'; Pattern='(?i)Active\s*Roles.*(Management Shell|PowerShell)' }
)
$PrerequisiteRoot = Join-Path $ScriptDirectory 'Prerequisites'
$script:PrerequisiteResults = @{}
$script:PrerequisiteRestartRequired = $false
$script:ActiveRolesReady = $false
$script:ModuleLoadError = ''

function Release-InstallerComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-PrerequisiteMsiMetadata {
    param([Parameter(Mandatory)][string]$Path)
    $Installer = $Database = $View = $Record = $null
    try {
        $Installer = New-Object -ComObject WindowsInstaller.Installer
        $Database = $Installer.OpenDatabase($Path, 0)
        $Properties = @{}
        foreach ($Key in @('ProductCode','UpgradeCode','ProductName','ProductVersion')) {
            $View = $Database.OpenView("SELECT ``Value`` FROM ``Property`` WHERE ``Property``='$Key'")
            $View.Execute()
            $Record = $View.Fetch()
            if ($Record) { $Properties[$Key] = [string]$Record.StringData(1) }
            Release-InstallerComObject $Record
            $Record = $null
            $View.Close()
            Release-InstallerComObject $View
            $View = $null
        }
        if ($Properties.ProductCode -notmatch '^\{[0-9A-Fa-f-]{36}\}$') {
            throw 'The MSI does not contain a valid ProductCode.'
        }
        return [pscustomobject]$Properties
    }
    finally {
        Release-InstallerComObject $Record
        Release-InstallerComObject $View
        Release-InstallerComObject $Database
        Release-InstallerComObject $Installer
    }
}

function Get-ActiveRolesInstalledProducts {
    # Explicit registry views also work from a 32-bit launcher. Only machine
    # installations count; installing for a different admin user is insufficient.
    foreach ($RegistryView in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
        $BaseKey = $UninstallKey = $null
        try {
            $BaseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $RegistryView)
            $UninstallKey = $BaseKey.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
            if (-not $UninstallKey) { continue }
            foreach ($KeyName in $UninstallKey.GetSubKeyNames()) {
                $ProductKey = $null
                try {
                    $ProductKey = $UninstallKey.OpenSubKey($KeyName)
                    if (-not $ProductKey) { continue }
                    $Name = [string]$ProductKey.GetValue('DisplayName')
                    if ($Name -match '(?i)Active\s*Roles') {
                        [pscustomobject]@{
                            ProductCode=$KeyName; Name=$Name
                            Version=[string]$ProductKey.GetValue('DisplayVersion')
                            IsMsi=($ProductKey.GetValue('WindowsInstaller') -eq 1)
                            View=$RegistryView.ToString()
                        }
                    }
                }
                finally { if ($ProductKey) { $ProductKey.Dispose() } }
            }
        }
        finally {
            if ($UninstallKey) { $UninstallKey.Dispose() }
            if ($BaseKey) { $BaseKey.Dispose() }
        }
    }
}

function Get-ActiveRolesPrerequisiteStatus {
    $Products = @(Get-ActiveRolesInstalledProducts)
    foreach ($Component in $script:Prerequisites) {
        $Path = Join-Path $PrerequisiteRoot $Component.RelativePath
        $Metadata = $null
        $MediaError = ''
        $MediaExists = Test-Path -LiteralPath $Path -PathType Leaf
        if ($MediaExists) {
            try { $Metadata = Get-PrerequisiteMsiMetadata -Path $Path }
            catch { $MediaError = $_.Exception.Message }
        }
        # Prefer exact package identity; recognize installed releases by component
        # name if the bundle is missing or a different release is already present.
        $MatchesInstalled = @($Products | Where-Object {
            $_.IsMsi -and $_.View -eq 'Registry64' -and
            (($Metadata -and $_.ProductCode -eq $Metadata.ProductCode) -or $_.Name -match $Component.Pattern)
        })
        $Installed = $MatchesInstalled.Count -gt 0
        $Status = 'Missing'
        $Detail = 'Ready to install from the bundled MSI.'
        if ($Installed) {
            $Status = 'Installed'
            $Detail = ($MatchesInstalled | ForEach-Object { "$($_.Name) $($_.Version)" }) -join '; '
        }
        elseif (-not $MediaExists) { $Status = 'Missing - installer not found'; $Detail = $Path }
        elseif ($MediaError) { $Status = 'Detection error'; $Detail = $MediaError }
        if ($Metadata -and $Metadata.ProductName -notmatch $Component.Pattern) {
            $MediaError = "Unexpected MSI product '$($Metadata.ProductName)' in the $($Component.Name) folder."
            if (-not $Installed) { $Status = 'Wrong installer'; $Detail = $MediaError }
        }
        [pscustomobject]@{
            Id=$Component.Id; Name=$Component.Name; Path=$Path; Installed=$Installed
            CanInstall=($MediaExists -and -not $MediaError -and -not $Installed)
            Metadata=$Metadata; Status=$Status; Detail=$Detail
        }
    }
}

function Test-ActiveRolesManagementShell {
    return [bool](Get-Module -ListAvailable -Name ActiveRolesManagementShell -ErrorAction SilentlyContinue)
}

function Refresh-PowerShellModulePath {
    $Paths = New-Object System.Collections.Generic.List[string]
    foreach ($Scope in @('Process','Machine','User')) {
        $Value = [Environment]::GetEnvironmentVariable('PSModulePath', $Scope)
        foreach ($Entry in ($Value -split ';')) {
            if (-not [string]::IsNullOrWhiteSpace($Entry) -and -not $Paths.Contains($Entry)) { $Paths.Add($Entry) }
        }
    }
    $env:PSModulePath = $Paths -join ';'
}

function Update-ActiveRolesModuleReadiness {
    Refresh-PowerShellModulePath
    $script:ActiveRolesReady = $false
    $script:ModuleLoadError = ''
    if (-not [Environment]::Is64BitProcess -or $PSVersionTable.PSEdition -eq 'Core') {
        $script:ModuleLoadError = 'Open this utility in 64-bit Windows PowerShell 5.1.'
        return
    }
    if (-not (Test-ActiveRolesManagementShell)) {
        $script:ModuleLoadError = 'ActiveRolesManagementShell is not available in this session.'
        return
    }
    try {
        Import-Module ActiveRolesManagementShell -Global -ErrorAction Stop
        $Required = @('Connect-QADService','Disconnect-QADService','Get-QADObject','Get-QADComputer','New-QADComputer','Set-QADComputer')
        $Missing = @($Required | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
        if ($Missing.Count) { throw "Required commands unavailable: $($Missing -join ', ')" }
        $script:ActiveRolesReady = $true
    }
    catch { $script:ModuleLoadError = $_.Exception.Message }
}

function Update-ActiveRolesPrerequisiteGui {
    # Called while Show-ImportConfiguration owns the controls (including events).
    $script:PrerequisiteStatus = @(Get-ActiveRolesPrerequisiteStatus)
    Update-ActiveRolesModuleReadiness
    foreach ($Component in $script:PrerequisiteStatus) {
        $Text = $Component.Status
        $Detail = $Component.Detail
        if ($script:PrerequisiteResults.ContainsKey($Component.Id)) {
            $Last = $script:PrerequisiteResults[$Component.Id]
            $Text += " | $($Last.Status)"
            $Detail += "`r`n$($Last.Detail)"
        }
        $Label = $PrerequisiteLabels[$Component.Id]
        $Label.Text = "$($Component.Name): $Text"
        $Label.ForeColor = if ($Component.Installed) { [Drawing.Color]::ForestGreen } else { [Drawing.Color]::Firebrick }
        $PrerequisiteToolTip.SetToolTip($Label, $Detail)
        $PrerequisiteInstallButtons[$Component.Id].Enabled = -not $Component.Installed -and -not $script:PrerequisiteRestartRequired
    }
    $Ready = $script:ActiveRolesReady -and
        @($script:PrerequisiteStatus | Where-Object { -not $_.Installed }).Count -eq 0 -and
        -not $script:PrerequisiteRestartRequired
    $PrerequisiteHint.Text = if ($script:PrerequisiteRestartRequired) {
        'Windows restart required. Restart before using Active Roles.'
    } elseif ($Ready) { 'All prerequisites ready.' }
    elseif ($script:ModuleLoadError) { $script:ModuleLoadError }
    else { 'Install missing prerequisites. Hover over a status for details.' }
    foreach ($Control in @($StartButton,$TestARConnectionMenuItem,$ValidateOUMenuItem)) { $Control.Enabled = $Ready }
    $InstallPrerequisitesButton.Enabled = @($script:PrerequisiteStatus | Where-Object { -not $_.Installed }).Count -gt 0 -and -not $script:PrerequisiteRestartRequired
}

# Runs in ONE elevated Windows PowerShell process. The original GUI and all QAD
# operations remain under the original user's identity. No credentials are saved.
$script:PrerequisiteBatchWorker = {
    param($Plan)
    $ErrorActionPreference = 'Stop'
    function Write-BatchState {
        $Json = $Batch | ConvertTo-Json -Depth 6
        [IO.File]::WriteAllText($Plan.StatusPath, $Json, [Text.Encoding]::UTF8)
    }
    $Batch = [ordered]@{ Complete=$false; RestartRequired=$false; Results=@() }
    foreach ($Item in $Plan.Items) {
        $Batch.Results += [pscustomobject]@{ Id=$Item.Id; Status='Queued'; Detail=''; ExitCode=$null }
    }
    $Failed = $false
    try {
        Write-BatchState
        for ($Index=0; $Index -lt $Plan.Items.Count; $Index++) {
            $Item = $Plan.Items[$Index]
            $Result = $Batch.Results[$Index]
            if ($Failed) { $Result.Status='Not run'; $Result.Detail='An earlier component failed.'; continue }
            $Installer = $null
            try {
                # Recheck identity immediately before installing. ProductState 5
                # includes machine installs available to this elevated account.
                $Installer = New-Object -ComObject WindowsInstaller.Installer
                if ($Installer.ProductState($Item.ProductCode) -eq 5) {
                    $Result.Status='Already installed'; $Result.Detail='Skipped after elevated recheck.'
                    Write-BatchState
                    continue
                }
                if ((Get-FileHash -LiteralPath $Item.Path -Algorithm SHA256).Hash -ne $Item.Hash) {
                    throw 'The staged MSI changed after preparation. Installation stopped.'
                }
                $Result.Status='Installing'; $Result.Detail="Log: $($Item.LogPath)"
                Write-BatchState
                $Arguments = '/i "{0}" /qn /norestart REBOOT=ReallySuppress ALLUSERS=1 /L*v "{1}"' -f $Item.Path,$Item.LogPath
                $Process = Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList $Arguments -Wait -PassThru -WindowStyle Hidden
                $Code = $Process.ExitCode
                $Process.Dispose()
                $Result.ExitCode = $Code
                if ($Code -notin @(0,3010,1641)) { throw "MSI failed with exit code $Code. See $($Item.LogPath)" }
                if ($Code -in @(3010,1641)) {
                    $Batch.RestartRequired = $true
                    $Result.Status = 'Installed - restart required'
                } else { $Result.Status = 'Installed' }
                if ($Installer.ProductState($Item.ProductCode) -ne 5) {
                    throw "MSI returned $Code, but installed product verification failed. See $($Item.LogPath)"
                }
                # 1641 means Windows is restarting; do not start another MSI.
                if ($Code -eq 1641) {
                    for ($Next=$Index+1; $Next -lt $Plan.Items.Count; $Next++) {
                        $Batch.Results[$Next].Status='Not run - restart required'
                    }
                    break
                }
            }
            catch { $Result.Status='Failed'; $Result.Detail=$_.Exception.Message; $Failed=$true }
            finally { if ($Installer -and [Runtime.InteropServices.Marshal]::IsComObject($Installer)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Installer) } }
            Write-BatchState
        }
    }
    finally { $Batch.Complete=$true; Write-BatchState }
    if ($Failed) { exit 1 }
    if ($Batch.RestartRequired) { exit 3010 }
    exit 0
}

function Install-ActiveRolesPrerequisites {
    param(
        [System.Windows.Forms.Form]$Owner,
        [ValidateSet('ADSI','SDK','Shell')][string]$ComponentId
    )
    $Status = @(Get-ActiveRolesPrerequisiteStatus)
    $Missing = @($Status | Where-Object { -not $_.Installed -and (-not $ComponentId -or $_.Id -eq $ComponentId) })
    if (-not $Missing.Count) { return }
    foreach ($Component in $Missing) { $script:PrerequisiteResults.Remove($Component.Id) }
    $Blocked = @($Missing | Where-Object { -not $_.CanInstall })
    if ($Blocked.Count) {
        $Message = ($Blocked | ForEach-Object { "$($_.Name): $($_.Status)`r`n$($_.Detail)" }) -join "`r`n`r`n"
        [void][Windows.Forms.MessageBox]::Show($Owner, $Message, 'Prerequisite media required', 'OK', 'Warning')
        return
    }
    $ProgressForm = New-Object Windows.Forms.Form
    $ProgressForm.Text = 'Install Active Roles Prerequisites'
    $ProgressForm.ClientSize = New-Object Drawing.Size(720,240)
    $ProgressForm.StartPosition = 'CenterParent'
    $ProgressForm.FormBorderStyle = 'FixedDialog'
    $ProgressForm.ControlBox = $false
    $ProgressForm.ShowInTaskbar = $false
    $ProgressLabel = New-Object Windows.Forms.Label
    $ProgressLabel.SetBounds(20,20,680,170)
    $ProgressLabel.Text = 'Preparing the missing prerequisite installers...'
    $ProgressForm.Controls.Add($ProgressLabel)
    $ProgressBar = New-Object Windows.Forms.ProgressBar
    $ProgressBar.SetBounds(20,200,680,22)
    $ProgressBar.Style = 'Marquee'
    $ProgressForm.Controls.Add($ProgressBar)
    $Owner.Enabled = $false
    $ProgressForm.Show($Owner)
    $Worker = $null
    $BatchDirectory = $null
    $LastState = $null
    try {
        [Windows.Forms.Application]::DoEvents()
        # Stage each complete component folder (including any external CABs) to a
        # local folder accessible to both the launching user and elevated admin.
        $BatchDirectory = Join-Path ([IO.Path]::GetTempPath()) ('ARS-Prerequisites-' + [guid]::NewGuid().ToString('N'))
        [void](New-Item -ItemType Directory -Path $BatchDirectory -ErrorAction Stop)
        $Acl = New-Object Security.AccessControl.DirectorySecurity
        $Acl.SetAccessRuleProtection($true,$false)
        $UserSid = [Security.Principal.WindowsIdentity]::GetCurrent().User
        foreach ($Sid in @($UserSid, [Security.Principal.SecurityIdentifier]'S-1-5-32-544', [Security.Principal.SecurityIdentifier]'S-1-5-18')) {
            $Rule = New-Object Security.AccessControl.FileSystemAccessRule($Sid,'FullControl','ContainerInherit,ObjectInherit','None','Allow')
            $Acl.AddAccessRule($Rule)
        }
        Set-Acl -LiteralPath $BatchDirectory -AclObject $Acl -ErrorAction Stop
        $Items = @()
        foreach ($Component in $Missing) {
            $ProgressLabel.Text = "Preparing $($Component.Name)..."
            [Windows.Forms.Application]::DoEvents()
            $Destination = Join-Path $BatchDirectory $Component.Id
            Copy-Item -LiteralPath (Split-Path $Component.Path -Parent) -Destination $Destination -Recurse -ErrorAction Stop
            $StagedPath = Join-Path $Destination '_x64.msi'
            $Items += [pscustomobject]@{
                Id=$Component.Id; Path=$StagedPath; ProductCode=$Component.Metadata.ProductCode
                Hash=(Get-FileHash -LiteralPath $StagedPath -Algorithm SHA256).Hash
                LogPath=(Join-Path $BatchDirectory ($Component.Id + '.log'))
            }
        }
        $Plan = [pscustomobject]@{ Items=$Items; StatusPath=(Join-Path $BatchDirectory 'status.json') }
        # The plan is data, not interpolated PowerShell. The worker code is passed
        # directly to Windows PowerShell; no user-writable elevated .ps1 is loaded.
        $PlanJson = $Plan | ConvertTo-Json -Depth 5 -Compress
        $Plan64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($PlanJson))
        $Command = '$Plan = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(''' + $Plan64 + ''')) | ConvertFrom-Json; & {' + $script:PrerequisiteBatchWorker.ToString() + '} $Plan'
        $Encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Command))
        if ($Encoded.Length -gt 30000) { throw 'The installation command is too long. Move the package to a shorter local path.' }
        $PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        if (-not [Environment]::Is64BitProcess) { $PowerShellExe = "$env:SystemRoot\Sysnative\WindowsPowerShell\v1.0\powershell.exe" }
        $ProgressLabel.Text = 'Approve the Windows administrator prompt to install the missing components.'
        [Windows.Forms.Application]::DoEvents()
        # This is the only elevation request for the whole batch.
        $Worker = Start-Process -FilePath $PowerShellExe -Verb RunAs -WindowStyle Hidden -PassThru -ErrorAction Stop -ArgumentList "-NoProfile -NonInteractive -EncodedCommand $Encoded"
        do {
            if (Test-Path -LiteralPath $Plan.StatusPath) {
                try {
                    $Candidate = Get-Content -LiteralPath $Plan.StatusPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    if ($Candidate.Results) { $LastState = $Candidate }
                } catch { } # The writer may currently have the JSON file open.
            }
            if ($LastState) {
                $ProgressLabel.Text = ($LastState.Results | ForEach-Object { "$($_.Id): $($_.Status)" }) -join "`r`n"
                foreach ($Result in $LastState.Results) { $script:PrerequisiteResults[$Result.Id] = $Result }
            }
            [Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 150
            $Worker.Refresh()
        } while (-not $Worker.HasExited)
        # Read the final state after process exit; do not rely on the last poll.
        if (Test-Path -LiteralPath $Plan.StatusPath) {
            $LastState = Get-Content -LiteralPath $Plan.StatusPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        }
        if (-not $LastState -or -not $LastState.Complete) { throw "The elevated batch did not return a complete result (exit $($Worker.ExitCode))." }
        foreach ($Result in $LastState.Results) { $script:PrerequisiteResults[$Result.Id] = $Result }
        $script:PrerequisiteRestartRequired = [bool]$LastState.RestartRequired
        $Summary = ($LastState.Results | ForEach-Object { "$($_.Id): $($_.Status)`r`n$($_.Detail)" }) -join "`r`n`r`n"
        if ($script:PrerequisiteRestartRequired) { $Summary += "`r`n`r`nRestart Windows before using Active Roles." }
        $Summary += "`r`n`r`nInstaller logs: $BatchDirectory"
        [void][Windows.Forms.MessageBox]::Show($ProgressForm, $Summary, 'Prerequisite installation results', 'OK', 'Information')
    }
    catch {
        $Detail = $_.Exception.Message
        foreach ($Component in $Missing) {
            if (-not $script:PrerequisiteResults.ContainsKey($Component.Id) -or $script:PrerequisiteResults[$Component.Id].Status -in @('Queued','Installing')) {
                $script:PrerequisiteResults[$Component.Id] = [pscustomobject]@{ Status='Not completed'; Detail=$Detail }
            }
        }
        [void][Windows.Forms.MessageBox]::Show($ProgressForm, "Installation was cancelled or could not complete.`r`n`r`n$Detail`r`n`r`nLogs: $BatchDirectory", 'Prerequisite installation', 'OK', 'Warning')
    }
    finally {
        if ($Worker) { $Worker.Dispose() }
        $ProgressForm.Close()
        $ProgressForm.Dispose()
        $Owner.Enabled = $true
        $Owner.Activate()
        # Preserve logs and staged media for diagnostics/MSI maintenance.
        Update-ActiveRolesPrerequisiteGui
    }
}


# ============================================================
# DEFAULT ATTRIBUTES
# ============================================================

$DefaultAttributeStrings = [ordered]@{
    edsvaCHSServer           = "FALSE"
    edsaJoinComputerToDomain = "US\Domain Users"
}

$script:ConfigAttributeStrings =
    [ordered]@{}

$script:ComputerObjectAttributes =
    @{}

# ============================================================
# CONFIG FUNCTIONS
# ============================================================

function Get-IniConfiguration {

    param (
        [Parameter(Mandatory)]
        [string]$Path,
        [switch]$StrictRead
    )

    $Result = [PSCustomObject]@{
        Settings   = [ordered]@{}
        Attributes = [ordered]@{}
        OUs = [ordered]@{}
    }

    if (-not (Test-Path $Path -PathType Leaf)) {
        return $Result
    }

    try {

        $Lines =
            Get-Content `
                -Path $Path `
                -ErrorAction Stop
    }
    catch {
        if ($StrictRead) { throw }

        return $Result
    }

    $CurrentSection = ""

    foreach ($RawLine in $Lines) {

        $Line = $RawLine.Trim()

        if ([string]::IsNullOrWhiteSpace($Line)) {
            continue
        }

        if (
            $Line.StartsWith(";") -or
            $Line.StartsWith("#")
        ) {
            continue
        }

        if ($Line -match '^\[(.+)\]$') {

            $CurrentSection =
                $Matches[1].Trim()

            continue
        }

        if (
            $Line -match
            '^\s*([^=]+?)\s*=\s*(.*)\s*$'
        ) {

            $Key =
                $Matches[1].Trim()

            $Value =
                $Matches[2].Trim()

            switch ($CurrentSection.ToLower()) {

                "settings" {

                    $Result.Settings[$Key] =
                        $Value
                }

                "ou" { $Result.OUs[$Key] = $Value }

                "attributes" {

                    $Result.Attributes[$Key] =
                        $Value
                }
            }
        }
    }

    return $Result
}

function Convert-IniAttributeValue {

    param (
        [AllowNull()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $TrimmedValue =
        $Value.Trim()

    if ($TrimmedValue -match '^(?i:true)$') {
        return $true
    }

    if ($TrimmedValue -match '^(?i:false)$') {
        return $false
    }

    if ($TrimmedValue -match '^(?i:null)$') {
        return $null
    }

    $IntegerValue = 0L

    if (
        [Int64]::TryParse(
            $TrimmedValue,
            [ref]$IntegerValue
        )
    ) {

        return $IntegerValue
    }

    return $TrimmedValue
}

function Update-ComputerObjectAttributes {

    $script:ComputerObjectAttributes =
        @{}

    foreach (
        $AttributeName in
        $script:ConfigAttributeStrings.Keys
    ) {

        $ConvertedValue =
            Convert-IniAttributeValue `
                -Value `
                $script:ConfigAttributeStrings[$AttributeName]

        $script:ComputerObjectAttributes[$AttributeName] =
            $ConvertedValue
    }
}

function Update-ImportPreview {
    if (-not $ImportPreviewBox -or $ImportPreviewBox.IsDisposed) { return }
    $Lines = New-Object System.Collections.Generic.List[string]
    $Lines.Add('Destination OU (new computers):')
    if ($OUComboBox.SelectedItem) {
        $Lines.Add([string]$OUComboBox.SelectedItem.FriendlyName)
        $Lines.Add([string]$OUComboBox.SelectedItem.DistinguishedName)
    } else { $Lines.Add('No destination OU selected.') }
    $Lines.Add('')
    $Lines.Add('Configured attributes to apply:')
    if ($script:AttributeReadError) {
        $Lines.Add('Unavailable: config.ini could not be read.')
    } elseif ($script:ConfigAttributeStrings.Count -eq 0) {
        $Lines.Add('(None)')
    } else {
        foreach ($Name in $script:ConfigAttributeStrings.Keys) {
            $Lines.Add("$Name = $($script:ConfigAttributeStrings[$Name])")
        }
    }
    $Lines.Add('')
    $Lines.Add('Existing computers receive these attributes but are not moved.')
    $PreviewText = $Lines -join "`r`n"
    # Avoid resetting the user's scroll position on every timer tick.
    if ($ImportPreviewBox.Text -ne $PreviewText) { $ImportPreviewBox.Text = $PreviewText }
}

function Refresh-ConfiguredAttributes {
    # Existing files are authoritative: removed keys must not reappear as defaults.
    # Only attributes are refreshed; unsaved GUI server/path fields stay intact.
    try {
        $Next = [ordered]@{}
        $Exists = Test-Path -LiteralPath $ConfigPath -PathType Leaf -ErrorAction Stop
        if ($Exists) {
            $Latest = Get-IniConfiguration -Path $ConfigPath -StrictRead
            foreach ($Key in $Latest.Attributes.Keys) { $Next[$Key] = $Latest.Attributes[$Key] }
        }
        else {
            foreach ($Key in $DefaultAttributeStrings.Keys) { $Next[$Key] = $DefaultAttributeStrings[$Key] }
        }
        $Converted = @{}
        foreach ($Key in $Next.Keys) { $Converted[$Key] = Convert-IniAttributeValue -Value $Next[$Key] }
        $script:ConfigAttributeStrings = $Next
        $script:ComputerObjectAttributes = $Converted
        $script:AttributeReadError = ''
        if ($AttributeLabel -and -not $AttributeLabel.IsDisposed) {
            $AttributeLabel.Text = "Configured Attributes: $($Converted.Count)"
            if (-not $Exists) { $AttributeLabel.Text += ' (defaults - config.ini not found)' }
            $AttributeLabel.ForeColor = [Drawing.SystemColors]::ControlText
        }
        Update-ImportPreview
        return $true
    }
    catch {
        # Keep the last complete snapshot during file locks/read errors, but
        # never label it as current or allow a save/import to use it silently.
        $script:AttributeReadError = $_.Exception.Message
        if ($AttributeLabel -and -not $AttributeLabel.IsDisposed) {
            $AttributeLabel.Text = 'Configured Attributes: config.ini could not be read'
            $AttributeLabel.ForeColor = [Drawing.Color]::Firebrick
        }
        Update-ImportPreview
        return $false
    }
}

function Refresh-ConfiguredOUs {
    try {
        $Latest = Get-IniConfiguration -Path $ConfigPath -StrictRead
        $Entries = @($Latest.OUs.Keys | Where-Object { -not [string]::IsNullOrWhiteSpace($Latest.OUs[$_]) } | ForEach-Object {
            [pscustomobject]@{ FriendlyName=[string]$_; DistinguishedName=[string]$Latest.OUs[$_] }
        })
        $Signature = ConvertTo-Json -InputObject $Entries -Compress
        if ($OUComboBox.Tag -ne $Signature) {
            $Previous = $OUComboBox.SelectedItem
            $OUComboBox.BeginUpdate()
            try {
                $OUComboBox.Items.Clear()
                foreach ($Entry in $Entries) { [void]$OUComboBox.Items.Add($Entry) }
                $OUComboBox.SelectedIndex = -1
                for ($Index=0; $Index -lt $Entries.Count; $Index++) {
                    if (($Previous -and $Entries[$Index].FriendlyName -eq $Previous.FriendlyName -and $Entries[$Index].DistinguishedName -eq $Previous.DistinguishedName) -or
                        ($null -eq $OUComboBox.Tag -and $Entries[$Index].DistinguishedName -eq $DefaultDestinationOU)) {
                        $OUComboBox.SelectedIndex = $Index
                        break
                    }
                }
                $OUComboBox.Tag = $Signature
            } finally { $OUComboBox.EndUpdate() }
        }
        $OUConfigHint.Text = if ($Entries.Count) { '' } else { 'Add FriendlyName=OU designation entries to [OU] in config.ini.' }
        Update-ImportPreview
        return $true
    }
    catch {
        $OUComboBox.Items.Clear()
        $OUComboBox.Tag = ''
        $OUConfigHint.Text = 'Unable to read [OU] from config.ini. Check the file and try again.'
        Update-ImportPreview
        return $false
    }
}

function Get-SelectedOuDesignation {
    if (-not (Refresh-ConfiguredOUs)) { return '' }
    if ($OUComboBox.SelectedItem) { return [string]$OUComboBox.SelectedItem.DistinguishedName }
    return ''
}

function Save-IniSettings {

    param (
        [Parameter(Mandatory)]
        [string]$Path,

        [string]$ARServer = "",

        [string]$DomainDN = "",

        [string]$CSVPath = "",

        [string]$DestinationOU = "",

        [string]$ReportPath = ""
    )

    if (-not (Refresh-ConfiguredAttributes)) {
        throw "Cannot save while config.ini attributes are unreadable: $script:AttributeReadError"
    }

    $Lines =
        New-Object System.Collections.Generic.List[string]

    $Lines.Add("[Settings]")
    $Lines.Add("ARServer=$ARServer")
    $Lines.Add("DomainDN=$DomainDN")
    $Lines.Add("CSVPath=$CSVPath")
    $Lines.Add("DestinationOU=$DestinationOU")
    $Lines.Add("ReportPath=$ReportPath")

    $Lines.Add("")
    $Lines.Add("[Attributes]")

    foreach (
        $AttributeName in
        $script:ConfigAttributeStrings.Keys
    ) {

        $AttributeValue =
            $script:ConfigAttributeStrings[$AttributeName]

        $Lines.Add(
            "$AttributeName=$AttributeValue"
        )
    }

    # Preserve the latest friendly-name mappings when saving GUI settings.
    $LatestOUConfig = Get-IniConfiguration -Path $Path -StrictRead
    $Lines.Add('')
    $Lines.Add('[OU]')
    foreach ($Name in $LatestOUConfig.OUs.Keys) {
        $Lines.Add("$Name=$($LatestOUConfig.OUs[$Name])")
    }

    Set-Content `
        -Path $Path `
        -Value $Lines `
        -Encoding UTF8 `
        -Force `
        -ErrorAction Stop
}

# ============================================================
# LOAD CONFIGURATION
# ============================================================

$Config =
    Get-IniConfiguration `
        -Path $ConfigPath

[void](Refresh-ConfiguredAttributes)

# ============================================================
# LOAD GUI DEFAULTS
# ============================================================

$DefaultARServer = if (
    $Config.Settings.Contains("ARServer")
) {

    $Config.Settings["ARServer"]
}
else {
    ""
}

$DefaultDomainDN = if (
    $Config.Settings.Contains("DomainDN")
) {

    $Config.Settings["DomainDN"]
}
else {
    ""
}

$DefaultCSVPath = if (
    $Config.Settings.Contains("CSVPath")
) {

    $Config.Settings["CSVPath"]
}
else {
    ""
}

$DefaultDestinationOU = if (
    $Config.Settings.Contains("DestinationOU")
) {

    $Config.Settings["DestinationOU"]
}
else {
    ""
}

$DefaultReportPath = if (
    $Config.Settings.Contains("ReportPath")
) {

    $Config.Settings["ReportPath"]
}
else {
    ""
}

# ============================================================
# ATTRIBUTE DISPLAY
# ============================================================

function Get-AttributeDisplayText {

    $Lines = @()

    foreach (
        $AttributeName in
        $script:ConfigAttributeStrings.Keys
    ) {

        $Value =
            $script:ConfigAttributeStrings[$AttributeName]

        $Lines +=
            "$AttributeName = $Value"
    }

    return (
        $Lines -join "`r`n"
    )
}

# ============================================================
# DISTINGUISHED NAME HELPER
# ============================================================

function Get-ParentDistinguishedName {

    param (
        [Parameter(Mandatory)]
        [string]$DistinguishedName
    )

    $Escaped = $false

    for (
        $i = 0;
        $i -lt $DistinguishedName.Length;
        $i++
    ) {

        $Character =
            $DistinguishedName[$i]

        if ($Character -eq '\') {

            $Escaped =
                -not $Escaped

            continue
        }

        if (
            $Character -eq ',' -and
            -not $Escaped
        ) {

            return $DistinguishedName.Substring(
                $i + 1
            )
        }

        $Escaped = $false
    }

    return ""
}

# ============================================================
# ACTIVE ROLES OBJECT LOOKUP
# ============================================================

function Get-ActiveRolesObjectByDN {

    param (
        [Parameter(Mandatory)]
        [string]$ARServer,

        [Parameter(Mandatory)]
        [string]$DistinguishedName
    )

    $Connection = $null

    try {

        $Connection =
            Connect-QADService `
                -Service $ARServer `
                -Proxy `
                -ErrorAction Stop

        $Object =
            Get-QADObject `
                -SearchRoot $DistinguishedName `
                -SearchScope Base `
                -Connection $Connection `
                -ErrorAction Stop

        return $Object
    }
    finally {

        if ($Connection) {

            Disconnect-QADService `
                -Connection $Connection `
                -ErrorAction SilentlyContinue
        }
    }
}

# ============================================================
# ACTIVE ROLES OU VALIDATION
# ============================================================

function Test-ActiveRolesOU {

    param (
        [Parameter(Mandatory)]
        [string]$ARServer,

        [Parameter(Mandatory)]
        [string]$DistinguishedName
    )

    try {

        $Object =
            Get-ActiveRolesObjectByDN `
                -ARServer $ARServer `
                -DistinguishedName $DistinguishedName

        if (-not $Object) {
            return $false
        }

        $ObjectClasses = @(
            $Object.ObjectClass
        )

        if (
            $ObjectClasses -contains
            "organizationalUnit"
        ) {

            return $true
        }

        if (
            $Object.DN -match
            '^(?i)OU='
        ) {

            return $true
        }

        return $false
    }
    catch {

        return $false
    }
}

# ============================================================
# ABOUT WINDOW
# ============================================================

function Show-ReadmeWindow {
    $ReadmePath = Join-Path $ScriptDirectory 'README.md'
    try {
        $ReadmeText = [IO.File]::ReadAllText($ReadmePath)
    }
    catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Unable to open the user guide at:`r`n$ReadmePath`r`n`r`n$($_.Exception.Message)",
            'ARS User Guide', 'OK', 'Warning')
        return
    }

    $ReadmeForm = New-Object System.Windows.Forms.Form
    $ReadmeForm.Text = 'ARS Bulk Computer Prestage - User Guide (README.md)'
    $ReadmeForm.Size = New-Object System.Drawing.Size(900, 720)
    $ReadmeForm.MinimumSize = New-Object System.Drawing.Size(600, 400)
    $ReadmeForm.StartPosition = 'CenterParent'
    $ReadmeForm.MinimizeBox = $false

    $ReadmeBox = New-Object System.Windows.Forms.TextBox
    $ReadmeBox.Multiline = $true
    $ReadmeBox.ReadOnly = $true
    $ReadmeBox.ScrollBars = 'Vertical'
    $ReadmeBox.WordWrap = $true
    $ReadmeBox.Dock = 'Fill'
    $ReadmeBox.Font = New-Object System.Drawing.Font('Consolas', 10)
    $ReadmeBox.BackColor = [System.Drawing.SystemColors]::Window
    $ReadmeBox.Text = $ReadmeText -replace '\r?\n', "`r`n"
    $ReadmeBox.Select(0, 0)
    $ReadmeForm.Controls.Add($ReadmeBox)
    try { [void]$ReadmeForm.ShowDialog() }
    finally { $ReadmeForm.Dispose() }
}

function Show-AboutWindow {

    $AboutForm =
        New-Object System.Windows.Forms.Form

    $AboutForm.Text =
        "About ARS Bulk Computer Prestage"

    $AboutForm.Size =
        New-Object System.Drawing.Size(
            760,
            720
        )

    $AboutForm.StartPosition =
        "CenterParent"

    $AboutForm.MinimizeBox =
        $false

    $AboutForm.MaximizeBox =
        $false

    $AboutForm.FormBorderStyle =
        "FixedDialog"

    $TitleLabel =
        New-Object System.Windows.Forms.Label

    $TitleLabel.Text =
        "ARS Bulk Computer Prestage"

    $TitleLabel.Font =
        New-Object System.Drawing.Font(
            "Segoe UI",
            16,
            [System.Drawing.FontStyle]::Bold
        )

    $TitleLabel.Location =
        New-Object System.Drawing.Point(
            20,
            20
        )

    $TitleLabel.AutoSize =
        $true

    $AboutForm.Controls.Add(
        $TitleLabel
    )

    $VersionLabel =
        New-Object System.Windows.Forms.Label

    $VersionLabel.Text =
        "Version 2.1"

    $VersionLabel.ForeColor =
        [System.Drawing.Color]::DimGray

    $VersionLabel.Location =
        New-Object System.Drawing.Point(
            22,
            55
        )

    $VersionLabel.AutoSize =
        $true

    $AboutForm.Controls.Add(
        $VersionLabel
    )

    $DocumentationBox =
        New-Object System.Windows.Forms.TextBox

    $DocumentationBox.Location =
        New-Object System.Drawing.Point(
            20,
            90
        )

    $DocumentationBox.Size =
        New-Object System.Drawing.Size(
            700,
            530
        )

    $DocumentationBox.Multiline =
        $true

    $DocumentationBox.ReadOnly =
        $true

    $DocumentationBox.ScrollBars =
        "Vertical"

    $DocumentationBox.WordWrap =
        $true

    $DocumentationBox.Font =
        New-Object System.Drawing.Font(
            "Segoe UI",
            9
        )

    $DocumentationBox.BackColor =
        [System.Drawing.SystemColors]::Window

    $AttributeText =
        Get-AttributeDisplayText

    $DocumentationBox.Text = @"
AUTHOR
======================================================================

CJ Micklitsch

Created with assistance from:
Karen (OpenAI Codex)


PURPOSE
======================================================================

ARS Bulk Computer Prestage provides a graphical interface for bulk
prestaging computer accounts through One Identity Active Roles.

The utility imports computer names from a CSV file, allows an
administrator to select a configured destination Organizational Unit, and
creates or updates computer accounts through Active Roles.


ACTIVE ROLES CONNECTION
======================================================================

Connection Mode:

Proxy

The Windows credentials of the currently logged-in user are used
for Active Roles operations.

No Active Roles credentials are stored in the script or config.ini.


ACTIVE ROLES PREREQUISITES
======================================================================

Tools > Active Roles Prerequisites shows ADSI Provider, SDK and
PowerShell Module status, with individual Install buttons and Install All.
Install All installs missing components in that order using one
Windows administrator prompt. Installed components are skipped.
Installer paths are relative to the script's Prerequisites folder.
Hover over a component status for detection details or installation results.
The submenu's Refresh Module / Prerequisite Status refreshes all indicators.
Directory actions require all prerequisites and a successfully loaded module.
Windows restart requirements are displayed even when module loading succeeds.

DESTINATION ORGANIZATIONAL UNIT
======================================================================

Destination OU is a friendly-name dropdown loaded from [OU] in config.ini.
Use FriendlyName=OU distinguished name entries. Only the friendly name is
displayed; validation and import use its mapped distinguished name.
The selected destination is validated through Active Roles before import.


CONFIGURATION
======================================================================

Configuration is stored in:

$ConfigPath

The [Settings] section stores:

- Active Roles Server
- Domain Distinguished Name
- Computer CSV path
- Destination OU
- Report output folder


ATTRIBUTES
======================================================================

Attributes are dynamically loaded from the [Attributes] section of
config.ini.

Configured Attributes:

$AttributeText

Default:

edsvaCHSServer = FALSE
edsaJoinComputerToDomain = US\Domain Users

Additional attributes may be added as needed.


CSV FORMAT
======================================================================

The CSV must contain:

ComputerName

Example:

ComputerName
PC-001
PC-002
PC-003


COMPUTER ACCOUNT BEHAVIOR
======================================================================

NEW COMPUTER

The computer account is created in the selected destination OU.

All configured attributes are applied during creation.


EXISTING COMPUTER

Configured attributes are updated on the existing computer account.

Existing computer accounts are NOT automatically moved to the
selected destination OU.


REPORTING
======================================================================

A timestamped CSV report is generated after each import.

Report filenames use:

ComputerImportResults_yyyyMMdd_HHmmss.csv


CREDITS
======================================================================

This utility was created by CJ Micklitsch with the assistance of
Karen (OpenAI Codex).

USER GUIDE
======================================================================

Choose About > View User Guide (README.md) for the full instructions.
The guide is loaded from the README.md file beside this script.
"@

    $AboutForm.Controls.Add(
        $DocumentationBox
    )

    $CloseButton =
        New-Object System.Windows.Forms.Button

    $CloseButton.Text =
        "Close"

    $CloseButton.Location =
        New-Object System.Drawing.Point(
            620,
            635
        )

    $CloseButton.Size =
        New-Object System.Drawing.Size(
            100,
            32
        )

    $CloseButton.Add_Click({

        $AboutForm.Close()
    })

    $AboutForm.Controls.Add(
        $CloseButton
    )

    $AboutForm.AcceptButton =
        $CloseButton

    $AboutForm.CancelButton =
        $CloseButton

    [void]$AboutForm.ShowDialog()
}

# ============================================================
# ACTIVE ROLES OU BROWSER
# ============================================================

# ============================================================
# ASYNCHRONOUS OU RETRIEVAL WITH GUI PROGRESS (v1.6 behavior)
# ============================================================
function Get-ActiveRolesOrganizationalUnitsWithProgress {
    param(
        [Parameter(Mandatory)][string]$SearchBase,
        [Parameter(Mandatory)][string]$ARServer,
        [System.Windows.Forms.Form]$Owner
    )
    $ProgressForm = New-Object Windows.Forms.Form
    $ProgressForm.Text = 'Loading Organizational Units'
    $ProgressForm.ClientSize = New-Object Drawing.Size(470,170)
    $ProgressForm.StartPosition = 'CenterParent'
    $ProgressForm.FormBorderStyle = 'FixedDialog'
    $ProgressForm.ControlBox = $false
    $ProgressForm.ShowInTaskbar = $false
    $StatusLabel = New-Object Windows.Forms.Label
    $StatusLabel.SetBounds(20,20,430,45)
    $ProgressForm.Controls.Add($StatusLabel)
    $ProgressBar = New-Object Windows.Forms.ProgressBar
    $ProgressBar.SetBounds(20,75,430,24)
    $ProgressBar.Style = 'Marquee'
    $ProgressBar.MarqueeAnimationSpeed = 25
    $ProgressForm.Controls.Add($ProgressBar)
    $CancelButton = New-Object Windows.Forms.Button
    $CancelButton.Text = 'Cancel'
    $CancelButton.SetBounds(350,120,100,30)
    $ProgressForm.Controls.Add($CancelButton)
    $State = [hashtable]::Synchronized(@{ Cancelled=$false; Status='Connecting to Active Roles...'; StopResult=$null })
    $PowerShell = [PowerShell]::Create()
    $QueryScript = {
        param($Server,$Base,$State)
        $ErrorActionPreference = 'Stop'
        $ProgressPreference = 'SilentlyContinue'
        Import-Module ActiveRolesManagementShell -ErrorAction Stop
        $Connection = $null
        try {
            $Connection = Connect-QADService -Service $Server -Proxy -ErrorAction Stop
            $State.Status = 'Retrieving Organizational Units from Active Roles...'
            Get-QADObject -SearchRoot $Base -Type organizationalUnit -SizeLimit 0 -Connection $Connection -ErrorAction Stop |
                Select-Object Name,DN,DistinguishedName
        }
        finally {
            if ($Connection) { Disconnect-QADService -Connection $Connection -ErrorAction SilentlyContinue }
        }
    }
    [void]$PowerShell.AddScript($QueryScript.ToString()).AddArgument($ARServer).AddArgument($SearchBase).AddArgument($State)
    $CancelButton.Add_Click({
        $State.Cancelled = $true
        $State.Status = 'Cancelling the Active Roles query...'
        $CancelButton.Enabled = $false
        # BeginStop keeps the GUI responsive while a provider finishes its call.
        try { $State.StopResult = $PowerShell.BeginStop($null,$null) } catch { }
    })
    try {
        if ($Owner) { $Owner.Enabled=$false; $ProgressForm.Show($Owner) }
        else { $ProgressForm.Show() }
        $AsyncResult = $PowerShell.BeginInvoke()
        while (-not $AsyncResult.IsCompleted -or ($State.StopResult -and -not $State.StopResult.IsCompleted)) {
            $StatusLabel.Text = $State.Status
            [Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 75
        }
        if ($State.StopResult) { $PowerShell.EndStop($State.StopResult) }
        if ($State.Cancelled) { return $null }
        $StatusLabel.Text = 'Preparing Organizational Unit list...'
        [Windows.Forms.Application]::DoEvents()
        $Results = @($PowerShell.EndInvoke($AsyncResult))
        if ($PowerShell.HadErrors) { throw ($PowerShell.Streams.Error | Out-String) }
        # Preserve an empty or single-element collection for the tree/count UI.
        return ,$Results
    }
    catch {
        if (-not $State.Cancelled) {
            [void][Windows.Forms.MessageBox]::Show($ProgressForm, "Unable to retrieve Organizational Units through Active Roles.`r`n`r`n$($_.Exception.Message)", 'OU Browser Error', 'OK', 'Error')
        }
        return $null
    }
    finally {
        $ProgressForm.Close()
        $ProgressForm.Dispose()
        if ($Owner) { $Owner.Enabled=$true; $Owner.Activate() }
        $PowerShell.Dispose()
    }
}


function Select-ActiveRolesOrganizationalUnit {
    param (
        [Parameter(Mandatory)][string]$SearchBase,
        [Parameter(Mandatory)][string]$ARServer,
        [System.Windows.Forms.Form]$Owner
    )
    $OUs = Get-ActiveRolesOrganizationalUnitsWithProgress -SearchBase $SearchBase -ARServer $ARServer -Owner $Owner
    if ($null -eq $OUs) { return $null }
    if ($OUs.Count -eq 0) {
        [void][Windows.Forms.MessageBox]::Show('No Organizational Units were returned from Active Roles.', 'No Organizational Units Found', 'OK', 'Warning')
        return $null
    }
    $Form =
        New-Object System.Windows.Forms.Form

    $Form.Text =
        "Select Destination OU"

    $Form.Size =
        New-Object System.Drawing.Size(
            800,
            700
        )

    $Form.StartPosition =
        "CenterParent"

    $Form.MinimizeBox =
        $false

    $InstructionLabel =
        New-Object System.Windows.Forms.Label

    $InstructionLabel.Text =
        "Select the OU where the computer accounts will be created:"

    $InstructionLabel.Location =
        New-Object System.Drawing.Point(
            15,
            15
        )

    $InstructionLabel.AutoSize =
        $true

    $Form.Controls.Add(
        $InstructionLabel
    )

    $CountLabel = New-Object Windows.Forms.Label
    $CountLabel.SetBounds(540,15,225,20)
    $CountLabel.Text = "$($OUs.Count) Organizational Units loaded"
    $CountLabel.TextAlign = 'MiddleRight'
    $Form.Controls.Add($CountLabel)

    $Tree =
        New-Object System.Windows.Forms.TreeView

    $Tree.Location =
        New-Object System.Drawing.Point(
            15,
            45
        )

    $Tree.Size =
        New-Object System.Drawing.Size(
            750,
            520
        )

    $Tree.Anchor =
        "Top,Bottom,Left,Right"

    $Tree.HideSelection =
        $false

    $Form.Controls.Add(
        $Tree
    )

    $SelectedOUBox =
        New-Object System.Windows.Forms.TextBox

    $SelectedOUBox.Location =
        New-Object System.Drawing.Point(
            15,
            575
        )

    $SelectedOUBox.Size =
        New-Object System.Drawing.Size(
            750,
            23
        )

    $SelectedOUBox.ReadOnly =
        $true

    $Form.Controls.Add(
        $SelectedOUBox
    )

    $OKButton =
        New-Object System.Windows.Forms.Button

    $OKButton.Text =
        "Select OU"

    $OKButton.Location =
        New-Object System.Drawing.Point(
            560,
            610
        )

    $OKButton.Size =
        New-Object System.Drawing.Size(
            95,
            30
        )

    $OKButton.Enabled =
        $false

    $Form.Controls.Add(
        $OKButton
    )

    $CancelButton =
        New-Object System.Windows.Forms.Button

    $CancelButton.Text =
        "Cancel"

    $CancelButton.Location =
        New-Object System.Drawing.Point(
            670,
            610
        )

    $CancelButton.Size =
        New-Object System.Drawing.Size(
            95,
            30
        )

    $Form.Controls.Add(
        $CancelButton
    )

    $NodeLookup = @{}

    $RootNode =
        New-Object System.Windows.Forms.TreeNode

    $RootNode.Text =
        $SearchBase

    $RootNode.Tag =
        $SearchBase

    [void]$Tree.Nodes.Add(
        $RootNode
    )

    foreach ($OUItem in $OUs) {

        $DN = if ($OUItem.DN) {
            $OUItem.DN
        }
        elseif ($OUItem.DistinguishedName) {
            $OUItem.DistinguishedName
        }
        else {
            $null
        }

        if (
            [string]::IsNullOrWhiteSpace(
                $DN
            )
        ) {
            continue
        }

        $Node =
            New-Object System.Windows.Forms.TreeNode

        $Node.Text =
            $OUItem.Name

        $Node.Tag =
            $DN

        $NodeLookup[$DN] =
            $Node
    }

    foreach ($OUItem in $OUs) {

        $DN = if ($OUItem.DN) {
            $OUItem.DN
        }
        elseif ($OUItem.DistinguishedName) {
            $OUItem.DistinguishedName
        }
        else {
            $null
        }

        if (
            [string]::IsNullOrWhiteSpace($DN) -or
            -not $NodeLookup.ContainsKey($DN)
        ) {

            continue
        }

        $Node =
            $NodeLookup[$DN]

        $ParentDN =
            Get-ParentDistinguishedName `
                -DistinguishedName $DN

        if (
            $NodeLookup.ContainsKey(
                $ParentDN
            )
        ) {

            [void]$NodeLookup[$ParentDN].Nodes.Add(
                $Node
            )
        }
        else {

            [void]$RootNode.Nodes.Add(
                $Node
            )
        }
    }

    $RootNode.Expand()

    $Tree.Add_AfterSelect({

        if (
            $Tree.SelectedNode -and
            $Tree.SelectedNode.Tag -ne
            $SearchBase
        ) {

            $SelectedOUBox.Text =
                $Tree.SelectedNode.Tag

            $OKButton.Enabled =
                $true
        }
        else {

            $SelectedOUBox.Clear()

            $OKButton.Enabled =
                $false
        }
    })

    $Tree.Add_NodeMouseDoubleClick({

        if (
            $Tree.SelectedNode -and
            $Tree.SelectedNode.Tag -ne
            $SearchBase
        ) {

            $Form.Tag =
                $Tree.SelectedNode.Tag

            $Form.DialogResult =
                [System.Windows.Forms.DialogResult]::OK

            $Form.Close()
        }
    })

    $OKButton.Add_Click({

        if (
            $Tree.SelectedNode -and
            $Tree.SelectedNode.Tag -ne
            $SearchBase
        ) {

            $Form.Tag =
                $Tree.SelectedNode.Tag

            $Form.DialogResult =
                [System.Windows.Forms.DialogResult]::OK

            $Form.Close()
        }
    })

    $CancelButton.Add_Click({

        $Form.DialogResult =
            [System.Windows.Forms.DialogResult]::Cancel

        $Form.Close()
    })

    $Result =
        $(if ($Owner) { $Form.ShowDialog($Owner) } else { $Form.ShowDialog() })

    if (
        $Result -eq
        [System.Windows.Forms.DialogResult]::OK
    ) {

        return $Form.Tag
    }

    return $null
}

# ============================================================
# MAIN GUI
# ============================================================

function Show-ImportConfiguration {

    $Form =
        New-Object System.Windows.Forms.Form

    $Form.Text =
        "ARS Bulk Computer Prestage"

    $Form.Size =
        New-Object System.Drawing.Size(
            840,
            570
        )

    $Form.StartPosition =
        "CenterScreen"

    $Form.FormBorderStyle =
        "FixedDialog"

    $Form.MaximizeBox =
        $false

    $Form.MinimizeBox =
        $false

    # ========================================================
    # MENU
    # ========================================================

    $MenuStrip =
        New-Object System.Windows.Forms.MenuStrip

    $ConfigMenu =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $ConfigMenu.Text =
        "Config"

    $OpenConfigMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $OpenConfigMenuItem.Text =
        "Open Config File"

    $SaveConfigMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $SaveConfigMenuItem.Text =
        "Save Current Settings"

    $ResetConfigMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $ResetConfigMenuItem.Text =
        "Reset Saved Settings"

    $ExitMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $ExitMenuItem.Text =
        "Exit"

    [void]$ConfigMenu.DropDownItems.Add(
        $OpenConfigMenuItem
    )

    [void]$ConfigMenu.DropDownItems.Add(
        $SaveConfigMenuItem
    )

    [void]$ConfigMenu.DropDownItems.Add(
        $ResetConfigMenuItem
    )

    [void]$ConfigMenu.DropDownItems.Add(
        (
            New-Object System.Windows.Forms.ToolStripSeparator
        )
    )

    [void]$ConfigMenu.DropDownItems.Add(
        $ExitMenuItem
    )

    $ToolsMenu =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $ToolsMenu.Text =
        "Tools"

    $TestARConnectionMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $TestARConnectionMenuItem.Text =
        "Test Active Roles Connection"

    $ValidateOUMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $ValidateOUMenuItem.Text =
        "Validate Destination OU"

    $RefreshModuleStatusMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $RefreshModuleStatusMenuItem.Text =
        "Refresh Module Status"

    $OpenReportsMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $OpenReportsMenuItem.Text =
        "Open Reports Folder"

    $ShowAttributesMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $ShowAttributesMenuItem.Text =
        "Show Configured Attributes"

    [void]$ToolsMenu.DropDownItems.Add(
        $TestARConnectionMenuItem
    )

    [void]$ToolsMenu.DropDownItems.Add(
        $ValidateOUMenuItem
    )

    [void]$ToolsMenu.DropDownItems.Add(
        $RefreshModuleStatusMenuItem
    )

    [void]$ToolsMenu.DropDownItems.Add(
        $OpenReportsMenuItem
    )

    [void]$ToolsMenu.DropDownItems.Add(
        (
            New-Object System.Windows.Forms.ToolStripSeparator
        )
    )

    [void]$ToolsMenu.DropDownItems.Add(
        $ShowAttributesMenuItem
    )

    $AboutMenu =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $AboutMenu.Text =
        "About"

    $AboutProgramMenuItem =
        New-Object System.Windows.Forms.ToolStripMenuItem

    $AboutProgramMenuItem.Text =
        "About ARS Bulk Computer Prestage"

    [void]$AboutMenu.DropDownItems.Add(
        $AboutProgramMenuItem
    )

    $ReadmeMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $ReadmeMenuItem.Text = 'View User Guide (README.md)'
    $ReadmeMenuItem.Add_Click({ Show-ReadmeWindow })
    [void]$AboutMenu.DropDownItems.Add($ReadmeMenuItem)

    [void]$MenuStrip.Items.Add(
        $ConfigMenu
    )

    [void]$MenuStrip.Items.Add(
        $ToolsMenu
    )

    [void]$MenuStrip.Items.Add(
        $AboutMenu
    )

    $Form.MainMenuStrip =
        $MenuStrip

    $Form.Controls.Add(
        $MenuStrip
    )

    # ========================================================
    # HEADER
    # ========================================================

    $Header =
        New-Object System.Windows.Forms.Label

    $Header.Text =
        "Bulk Active Roles Computer Prestage"

    $Header.Font =
        New-Object System.Drawing.Font(
            "Segoe UI",
            14,
            [System.Drawing.FontStyle]::Bold
        )

    $Header.Location =
        New-Object System.Drawing.Point(
            20,
            45
        )

    $Header.AutoSize =
        $true

    $Form.Controls.Add(
        $Header
    )

    # ========================================================
    # MODULE STATUS INDICATOR
    # ========================================================

    # ========================================================
    # AR SERVER
    # ========================================================

    $ARServerLabel =
        New-Object System.Windows.Forms.Label

    $ARServerLabel.Text =
        "AR Server:"

    $ARServerLabel.Location =
        New-Object System.Drawing.Point(
            20,
            100
        )

    $ARServerLabel.AutoSize =
        $true

    $Form.Controls.Add(
        $ARServerLabel
    )

    $ARServerTextBox =
        New-Object System.Windows.Forms.TextBox

    $ARServerTextBox.Location =
        New-Object System.Drawing.Point(
            165,
            97
        )

    $ARServerTextBox.Size =
        New-Object System.Drawing.Size(
            620,
            23
        )

    $ARServerTextBox.Text =
        $DefaultARServer

    $Form.Controls.Add(
        $ARServerTextBox
    )

    # ========================================================
    # DOMAIN DN
    # ========================================================

    $DomainDNLabel =
        New-Object System.Windows.Forms.Label

    $DomainDNLabel.Text =
        "Domain DN:"

    $DomainDNLabel.Location =
        New-Object System.Drawing.Point(
            20,
            145
        )

    $DomainDNLabel.AutoSize =
        $true

    $Form.Controls.Add(
        $DomainDNLabel
    )

    $DomainDNTextBox =
        New-Object System.Windows.Forms.TextBox

    $DomainDNTextBox.Location =
        New-Object System.Drawing.Point(
            165,
            142
        )

    $DomainDNTextBox.Size =
        New-Object System.Drawing.Size(
            620,
            23
        )

    $DomainDNTextBox.Text =
        $DefaultDomainDN

    $Form.Controls.Add(
        $DomainDNTextBox
    )

    # ========================================================
    # CSV
    # ========================================================

    $CSVLabel =
        New-Object System.Windows.Forms.Label

    $CSVLabel.Text =
        "Computers CSV:"

    $CSVLabel.Location =
        New-Object System.Drawing.Point(
            20,
            195
        )

    $CSVLabel.AutoSize =
        $true

    $Form.Controls.Add(
        $CSVLabel
    )

    $CSVTextBox =
        New-Object System.Windows.Forms.TextBox

    $CSVTextBox.Location =
        New-Object System.Drawing.Point(
            165,
            192
        )

    $CSVTextBox.Size =
        New-Object System.Drawing.Size(
            520,
            23
        )

    $CSVTextBox.Text =
        $DefaultCSVPath

    $Form.Controls.Add(
        $CSVTextBox
    )

    $CSVBrowseButton =
        New-Object System.Windows.Forms.Button

    $CSVBrowseButton.Text =
        "Browse..."

    $CSVBrowseButton.Location =
        New-Object System.Drawing.Point(
            695,
            190
        )

    $CSVBrowseButton.Size =
        New-Object System.Drawing.Size(
            90,
            27
        )

    $Form.Controls.Add(
        $CSVBrowseButton
    )

    # ========================================================
    # OU
    # ========================================================

    $OULabel =
        New-Object System.Windows.Forms.Label

    $OULabel.Text =
        "Destination OU:"

    $OULabel.Location =
        New-Object System.Drawing.Point(
            20,
            245
        )

    $OULabel.AutoSize =
        $true

    $Form.Controls.Add(
        $OULabel
    )

    $OUComboBox = New-Object Windows.Forms.ComboBox
    $OUComboBox.SetBounds(165,242,620,23)
    $OUComboBox.DropDownStyle = 'DropDownList'
    $OUComboBox.DisplayMember = 'FriendlyName'
    $OUComboBox.ValueMember = 'DistinguishedName'
    $Form.Controls.Add($OUComboBox)
    $OUConfigHint = New-Object Windows.Forms.Label
    $OUConfigHint.SetBounds(165,268,620,18)
    $OUConfigHint.ForeColor = [Drawing.Color]::Firebrick
    $Form.Controls.Add($OUConfigHint)
    $OUComboBox.Add_DropDown({ [void](Refresh-ConfiguredOUs) })

    # ========================================================
    # REPORT FOLDER
    # ========================================================

    $ReportLabel =
        New-Object System.Windows.Forms.Label

    $ReportLabel.Text =
        "Report Folder:"

    $ReportLabel.Location =
        New-Object System.Drawing.Point(
            20,
            295
        )

    $ReportLabel.AutoSize =
        $true

    $Form.Controls.Add(
        $ReportLabel
    )

    $ReportTextBox =
        New-Object System.Windows.Forms.TextBox

    $ReportTextBox.Location =
        New-Object System.Drawing.Point(
            165,
            292
        )

    $ReportTextBox.Size =
        New-Object System.Drawing.Size(
            520,
            23
        )

    $ReportTextBox.Text =
        $DefaultReportPath

    $Form.Controls.Add(
        $ReportTextBox
    )

    $ReportBrowseButton =
        New-Object System.Windows.Forms.Button

    $ReportBrowseButton.Text =
        "Browse..."

    $ReportBrowseButton.Location =
        New-Object System.Drawing.Point(
            695,
            290
        )

    $ReportBrowseButton.Size =
        New-Object System.Drawing.Size(
            90,
            27
        )

    $Form.Controls.Add(
        $ReportBrowseButton
    )

    # ========================================================
    # ATTRIBUTE STATUS
    # ========================================================

    $AttributeLabel =
        New-Object System.Windows.Forms.Label

    $AttributeLabel.Text =
        "Configured Attributes: $($script:ComputerObjectAttributes.Count)"

    $AttributeLabel.Location =
        New-Object System.Drawing.Point(
            20,
            345
        )

    $AttributeLabel.AutoSize =
        $true

    $AttributeLabel.Font =
        New-Object System.Drawing.Font(
            "Segoe UI",
            9,
            [System.Drawing.FontStyle]::Bold
        )

    $Form.Controls.Add(
        $AttributeLabel
    )

    $ImportPreviewBox = New-Object Windows.Forms.TextBox
    $ImportPreviewBox.SetBounds(20,370,765,170)
    $ImportPreviewBox.Multiline = $true
    $ImportPreviewBox.ReadOnly = $true
    $ImportPreviewBox.WordWrap = $true
    $ImportPreviewBox.ScrollBars = 'Vertical'
    $ImportPreviewBox.BackColor = [Drawing.SystemColors]::Window
    $ImportPreviewBox.Font = New-Object Drawing.Font('Segoe UI',9)
    $ImportPreviewBox.AccessibleName = 'Destination OU and configured attributes to apply'
    $Form.Controls.Add($ImportPreviewBox)
    $OUComboBox.Add_SelectedIndexChanged({ Update-ImportPreview })

    # ========================================================
    # CONFIG STATUS
    # ========================================================

    $ConfigLabel =
        New-Object System.Windows.Forms.Label

    $ConfigLabel.Location =
        New-Object System.Drawing.Point(
            20,
            405
        )

    $ConfigLabel.Size =
        New-Object System.Drawing.Size(
            765,
            35
        )

    $ConfigLabel.ForeColor =
        [System.Drawing.Color]::DimGray

    if (
        Test-Path `
            $ConfigPath `
            -PathType Leaf
    ) {

        $ConfigLabel.Text =
            "Saved settings loaded from: $ConfigPath"
    }
    else {

        $ConfigLabel.Text =
            "No config.ini found. It will be created with the default attributes."
    }

    $Form.Controls.Add(
        $ConfigLabel
    )

    # ========================================================
    # BUTTONS
    # ========================================================

    $StartButton =
        New-Object System.Windows.Forms.Button

    $StartButton.Text =
        "Start Import"

    $StartButton.Location =
        New-Object System.Drawing.Point(
            555,
            470
        )

    $StartButton.Size =
        New-Object System.Drawing.Size(
            120,
            35
        )

    $Form.Controls.Add(
        $StartButton
    )

    $CancelButton =
        New-Object System.Windows.Forms.Button

    $CancelButton.Text =
        "Cancel"

    $CancelButton.Location =
        New-Object System.Drawing.Point(
            690,
            470
        )

    $CancelButton.Size =
        New-Object System.Drawing.Size(
            95,
            35
        )

    $Form.Controls.Add(
        $CancelButton
    )

    # ========================================================
    # CSV BROWSE
    # ========================================================

    $CSVBrowseButton.Add_Click({

        $Dialog =
            New-Object System.Windows.Forms.OpenFileDialog

        $Dialog.Title =
            "Select Computer Import CSV"

        $Dialog.Filter =
            "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"

        if (
            -not [string]::IsNullOrWhiteSpace(
                $CSVTextBox.Text
            ) -and
            (
                Test-Path `
                    $CSVTextBox.Text `
                    -PathType Leaf
            )
        ) {

            $Dialog.InitialDirectory =
                Split-Path `
                    $CSVTextBox.Text `
                    -Parent
        }

        if (
            $Dialog.ShowDialog() -eq
            [System.Windows.Forms.DialogResult]::OK
        ) {

            $CSVTextBox.Text =
                $Dialog.FileName

            if (
                [string]::IsNullOrWhiteSpace(
                    $ReportTextBox.Text
                )
            ) {

                $ReportTextBox.Text =
                    Split-Path `
                        $Dialog.FileName `
                        -Parent
            }
        }
    })

    # ========================================================
    # OU BROWSE
    # ========================================================

    # ========================================================
    # REPORT BROWSE
    # ========================================================

    $ReportBrowseButton.Add_Click({

        $Dialog =
            New-Object System.Windows.Forms.FolderBrowserDialog

        $Dialog.Description =
            "Select folder for Active Roles import reports"

        if (
            -not [string]::IsNullOrWhiteSpace(
                $ReportTextBox.Text
            ) -and
            (
                Test-Path `
                    $ReportTextBox.Text `
                    -PathType Container
            )
        ) {

            $Dialog.SelectedPath =
                $ReportTextBox.Text
        }

        if (
            $Dialog.ShowDialog() -eq
            [System.Windows.Forms.DialogResult]::OK
        ) {

            $ReportTextBox.Text =
                $Dialog.SelectedPath
        }
    })

    # ========================================================
    # CONFIG - OPEN
    # ========================================================

    $OpenConfigMenuItem.Add_Click({

        if (
            -not (
                Test-Path `
                    $ConfigPath `
                    -PathType Leaf
            )
        ) {

            try {

                Save-IniSettings `
                    -Path $ConfigPath `
                    -ARServer $ARServerTextBox.Text.Trim() `
                    -DomainDN $DomainDNTextBox.Text.Trim() `
                    -CSVPath $CSVTextBox.Text.Trim() `
                    -DestinationOU (Get-SelectedOuDesignation) `
                    -ReportPath $ReportTextBox.Text.Trim()
            }
            catch {

                [System.Windows.Forms.MessageBox]::Show(
                    "Unable to create config.ini.`r`n`r`n$($_.Exception.Message)",
                    "Configuration Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )

                return
            }
        }

        Start-Process `
            notepad.exe `
            -ArgumentList `
            "`"$ConfigPath`""
    })

    # ========================================================
    # CONFIG - SAVE
    # ========================================================

    $SaveConfigMenuItem.Add_Click({

        try {

            Save-IniSettings `
                -Path $ConfigPath `
                -ARServer $ARServerTextBox.Text.Trim() `
                -DomainDN $DomainDNTextBox.Text.Trim() `
                -CSVPath $CSVTextBox.Text.Trim() `
                -DestinationOU (Get-SelectedOuDesignation) `
                -ReportPath $ReportTextBox.Text.Trim()

            $ConfigLabel.Text =
                "Settings saved to: $ConfigPath"

            [System.Windows.Forms.MessageBox]::Show(
                "GUI settings and configured attributes have been saved.",
                "Configuration Saved",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "Unable to save config.ini.`r`n`r`n$($_.Exception.Message)",
                "Configuration Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    # ========================================================
    # CONFIG - RESET
    # ========================================================

    $ResetConfigMenuItem.Add_Click({

        $ConfirmReset =
            [System.Windows.Forms.MessageBox]::Show(
                "Clear the saved configuration?`r`n`r`nGUI settings will be cleared and attributes will return to their defaults.",
                "Reset Configuration",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

        if (
            $ConfirmReset -ne
            [System.Windows.Forms.DialogResult]::Yes
        ) {

            return
        }

        try {

            if (
                Test-Path `
                    $ConfigPath `
                    -PathType Leaf
            ) {

                Remove-Item `
                    -Path $ConfigPath `
                    -Force `
                    -ErrorAction Stop
            }

            $ARServerTextBox.Clear()
            $DomainDNTextBox.Clear()
            $CSVTextBox.Clear()
            $OUComboBox.SelectedIndex = -1
            $ReportTextBox.Clear()

            $script:ConfigAttributeStrings =
                [ordered]@{}

            foreach (
                $AttributeName in
                $DefaultAttributeStrings.Keys
            ) {

                $script:ConfigAttributeStrings[$AttributeName] =
                    $DefaultAttributeStrings[$AttributeName]
            }

            Update-ComputerObjectAttributes

            $AttributeLabel.Text =
                "Configured Attributes: $($script:ComputerObjectAttributes.Count)"

            $ConfigLabel.Text =
                "Configuration reset. Default attributes will be used."

            [System.Windows.Forms.MessageBox]::Show(
                "Configuration reset successfully.`r`n`r`nDefault attributes:`r`nedsvaCHSServer = FALSE`r`nedsaJoinComputerToDomain = US\Domain Users",
                "Configuration Reset",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "Unable to reset configuration.`r`n`r`n$($_.Exception.Message)",
                "Configuration Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    # ========================================================
    # EXIT
    # ========================================================

    $ExitMenuItem.Add_Click({

        $Form.DialogResult =
            [System.Windows.Forms.DialogResult]::Cancel

        $Form.Close()
    })

    # ========================================================
    # REFRESH MODULE STATUS
    # ========================================================

    $RefreshModuleStatusMenuItem.Add_Click({ Update-ActiveRolesPrerequisiteGui })

    # ========================================================
    # TEST ACTIVE ROLES CONNECTION
    # ========================================================

    $TestARConnectionMenuItem.Add_Click({
        Update-ActiveRolesPrerequisiteGui
        if (-not $StartButton.Enabled) { return }

        $CurrentARServer =
            $ARServerTextBox.Text.Trim()

        if (
            [string]::IsNullOrWhiteSpace(
                $CurrentARServer
            )
        ) {

            [System.Windows.Forms.MessageBox]::Show(
                "Enter an Active Roles server first.",
                "AR Server Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            return
        }

        $TestConnection = $null

        try {

            $TestConnection =
                Connect-QADService `
                    -Service $CurrentARServer `
                    -Proxy `
                    -ErrorAction Stop

            [System.Windows.Forms.MessageBox]::Show(
                "Successfully connected to Active Roles.`r`n`r`nServer: $CurrentARServer",
                "Connection Successful",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "Unable to connect to Active Roles.`r`n`r`n$($_.Exception.Message)",
                "Connection Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        finally {

            if ($TestConnection) {

                Disconnect-QADService `
                    -Connection $TestConnection `
                    -ErrorAction SilentlyContinue
            }
        }
    })

    # ========================================================
    # VALIDATE DESTINATION OU
    # ========================================================

    $ValidateOUMenuItem.Add_Click({
        Update-ActiveRolesPrerequisiteGui
        if (-not $StartButton.Enabled) { return }

        $CurrentARServer =
            $ARServerTextBox.Text.Trim()

        $CurrentOU =
            (Get-SelectedOuDesignation)

        if (
            [string]::IsNullOrWhiteSpace(
                $CurrentARServer
            )
        ) {

            [System.Windows.Forms.MessageBox]::Show(
                "Enter an Active Roles server first.",
                "AR Server Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            return
        }

        if (
            [string]::IsNullOrWhiteSpace(
                $CurrentOU
            )
        ) {

            [System.Windows.Forms.MessageBox]::Show(
                "Enter or select a destination OU first.",
                "OU Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            return
        }

        try {

            $ValidOU =
                Test-ActiveRolesOU `
                    -ARServer $CurrentARServer `
                    -DistinguishedName $CurrentOU

            if (-not $ValidOU) {

                throw "The specified Organizational Unit could not be found or is not an OU."
            }

            [System.Windows.Forms.MessageBox]::Show(
                "The destination OU is valid through Active Roles.`r`n`r`n$CurrentOU",
                "OU Validation Successful",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "OU validation failed through Active Roles.`r`n`r`n$($_.Exception.Message)",
                "OU Validation Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    # ========================================================
    # OPEN REPORTS
    # ========================================================

    $OpenReportsMenuItem.Add_Click({

        $CurrentReportPath =
            $ReportTextBox.Text.Trim()

        if (
            -not (
                Test-Path `
                    $CurrentReportPath `
                    -PathType Container
            )
        ) {

            [System.Windows.Forms.MessageBox]::Show(
                "The report folder does not exist.",
                "Invalid Report Folder",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            return
        }

        Start-Process `
            explorer.exe `
            -ArgumentList `
            "`"$CurrentReportPath`""
    })

    # ========================================================
    # SHOW ATTRIBUTES
    # ========================================================

    $ShowAttributesMenuItem.Add_Click({
        if (-not (Refresh-ConfiguredAttributes)) {
            [void][Windows.Forms.MessageBox]::Show($Form, "Unable to read config.ini attributes.`r`n`r`n$script:AttributeReadError", 'Configuration read error', 'OK', 'Warning')
            return
        }

        $AttributeText =
            Get-AttributeDisplayText

        [System.Windows.Forms.MessageBox]::Show(
            "The following Active Roles attributes will be applied to computer accounts:`r`n`r`n$AttributeText`r`n`r`nAdditional attributes may be added to the [Attributes] section of config.ini.",
            "Configured Attributes",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })

    # ========================================================
    # ABOUT
    # ========================================================

    $AboutProgramMenuItem.Add_Click({
        if (-not (Refresh-ConfiguredAttributes)) { return }

        Show-AboutWindow
    })

    # ========================================================
    # START IMPORT
    # ========================================================

    $StartButton.Add_Click({
        if (-not (Refresh-ConfiguredAttributes)) {
            [void][Windows.Forms.MessageBox]::Show($Form, "Unable to read config.ini attributes.`r`n`r`n$script:AttributeReadError", 'Configuration read error', 'OK', 'Warning')
            return
        }
        Update-ActiveRolesPrerequisiteGui
        if (-not $StartButton.Enabled) { return }

        $ARServer =
            $ARServerTextBox.Text.Trim()

        $DomainDN =
            $DomainDNTextBox.Text.Trim()

        $CSVPath =
            $CSVTextBox.Text.Trim()

        $OU =
            (Get-SelectedOuDesignation)

        $ReportPath =
            $ReportTextBox.Text.Trim()

        if (
            [string]::IsNullOrWhiteSpace($ARServer) -or
            [string]::IsNullOrWhiteSpace($DomainDN) -or
            [string]::IsNullOrWhiteSpace($CSVPath) -or
            [string]::IsNullOrWhiteSpace($OU) -or
            [string]::IsNullOrWhiteSpace($ReportPath)
        ) {

            [System.Windows.Forms.MessageBox]::Show(
                "All fields are required.",
                "Missing Information",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            return
        }

        if (
            -not (
                Test-Path `
                    $CSVPath `
                    -PathType Leaf
            )
        ) {

            [System.Windows.Forms.MessageBox]::Show(
                "The selected CSV file does not exist.",
                "Invalid CSV",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            return
        }

        if (
            -not (
                Test-Path `
                    $ReportPath `
                    -PathType Container
            )
        ) {

            [System.Windows.Forms.MessageBox]::Show(
                "The report folder does not exist.",
                "Invalid Report Folder",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            return
        }

        try {

            $DomainObject =
                Get-ActiveRolesObjectByDN `
                    -ARServer $ARServer `
                    -DistinguishedName $DomainDN

            if (-not $DomainObject) {

                throw "The specified domain could not be found."
            }
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "The Domain DN could not be validated through Active Roles.`r`n`r`n$($_.Exception.Message)",
                "Invalid Domain DN",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )

            return
        }

        try {

            $ValidOU =
                Test-ActiveRolesOU `
                    -ARServer $ARServer `
                    -DistinguishedName $OU

            if (-not $ValidOU) {

                throw "The specified Organizational Unit could not be found or is not an OU."
            }
        }
        catch {

            [System.Windows.Forms.MessageBox]::Show(
                "The destination OU could not be validated through Active Roles.`r`n`r`n$($_.Exception.Message)",
                "Invalid OU",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )

            return
        }

        try {

            Save-IniSettings `
                -Path $ConfigPath `
                -ARServer $ARServer `
                -DomainDN $DomainDN `
                -CSVPath $CSVPath `
                -DestinationOU $OU `
                -ReportPath $ReportPath

            $ConfigLabel.Text =
                "Settings saved to: $ConfigPath"
        }
        catch {

            $Continue =
                [System.Windows.Forms.MessageBox]::Show(
                    "Unable to save config.ini.`r`n`r`n$($_.Exception.Message)`r`n`r`nContinue anyway?",
                    "Configuration Warning",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )

            if (
                $Continue -ne
                [System.Windows.Forms.DialogResult]::Yes
            ) {

                return
            }
        }

        $Form.Tag =
            [PSCustomObject]@{
                ARServer   = $ARServer
                DomainDN   = $DomainDN
                CSVPath    = $CSVPath
                OU         = $OU
                ReportPath = $ReportPath
            }

        $Form.DialogResult =
            [System.Windows.Forms.DialogResult]::OK

        $Form.Close()
    })

    $CancelButton.Add_Click({

        $Form.DialogResult =
            [System.Windows.Forms.DialogResult]::Cancel

        $Form.Close()
    })

    # Reserve space for the destination/attribute preview and retain the footer.
    $Form.Height += 150
    $ConfigLabel.Top += 150
    $StartButton.Top += 150
    $CancelButton.Top += 150
    Update-ImportPreview

    # A hosted panel provides real Install buttons beside each status in Tools.
    $PrerequisitesMenu = New-Object Windows.Forms.ToolStripMenuItem
    $PrerequisitesMenu.Text = 'Active Roles Prerequisites'
    [void]$ToolsMenu.DropDownItems.Add($PrerequisitesMenu)
    $PrerequisitePanel = New-Object Windows.Forms.Panel
    $PrerequisitePanel.Size = New-Object Drawing.Size(700,185)
    $PrerequisiteToolTip = New-Object Windows.Forms.ToolTip
    $PrerequisiteToolTip.AutoPopDelay = 20000
    $PrerequisiteLabels = @{}
    $PrerequisiteInstallButtons = @{}
    $Row = 0
    foreach ($Component in $script:Prerequisites) {
        $Label = New-Object Windows.Forms.Label
        $Label.SetBounds(12,(14 + 35*$Row),555,25)
        $Label.AutoEllipsis = $true
        $PrerequisitePanel.Controls.Add($Label)
        $PrerequisiteLabels[$Component.Id] = $Label
        $InstallButton = New-Object Windows.Forms.Button
        $InstallButton.Text = 'Install'
        $InstallButton.AccessibleName = "Install $($Component.Name)"
        $InstallButton.Tag = $Component.Id
        $InstallButton.SetBounds(580,(10 + 35*$Row),105,28)
        $InstallButton.Add_Click({
            param($Sender,$EventArgs)
            $PrerequisitesMenu.HideDropDown()
            Install-ActiveRolesPrerequisites -Owner $Form -ComponentId ([string]$Sender.Tag)
        })
        $PrerequisitePanel.Controls.Add($InstallButton)
        $PrerequisiteInstallButtons[$Component.Id] = $InstallButton
        $Row++
    }
    $InstallPrerequisitesButton = New-Object Windows.Forms.Button
    $InstallPrerequisitesButton.Text = 'Install All'
    $InstallPrerequisitesButton.SetBounds(580,115,105,30)
    $PrerequisitePanel.Controls.Add($InstallPrerequisitesButton)
    $PrerequisiteHint = New-Object Windows.Forms.Label
    $PrerequisiteHint.SetBounds(12,118,555,55)
    $PrerequisiteHint.ForeColor = [Drawing.Color]::DimGray
    $PrerequisitePanel.Controls.Add($PrerequisiteHint)
    $PrerequisiteHost = New-Object Windows.Forms.ToolStripControlHost($PrerequisitePanel)
    $PrerequisiteHost.AutoSize = $false
    $PrerequisiteHost.Size = $PrerequisitePanel.Size
    [void]$PrerequisitesMenu.DropDownItems.Add($PrerequisiteHost)
    $ToolsMenu.DropDownItems.Remove($RefreshModuleStatusMenuItem)
    [void]$PrerequisitesMenu.DropDownItems.Add($RefreshModuleStatusMenuItem)
    $RefreshModuleStatusMenuItem.Text = 'Refresh Module / Prerequisite Status'
    $InstallPrerequisitesButton.Add_Click({
        $PrerequisitesMenu.HideDropDown()
        Install-ActiveRolesPrerequisites -Owner $Form
    })
    $PrerequisitesMenu.Add_DropDownOpening({ Update-ActiveRolesPrerequisiteGui })
    # WinForms timer executes on the GUI thread; it never modifies UI controls
    # from a filesystem watcher/background thread. Stop it when the form closes.
    $AttributeRefreshTimer = New-Object Windows.Forms.Timer
    $AttributeRefreshTimer.Interval = 1000
    $AttributeRefreshTimer.Add_Tick({ [void](Refresh-ConfiguredAttributes); if (-not $OUComboBox.DroppedDown) { [void](Refresh-ConfiguredOUs) } })
    $Form.Add_Shown({
        Update-ActiveRolesPrerequisiteGui
        [void](Refresh-ConfiguredAttributes)
        [void](Refresh-ConfiguredOUs)
        $AttributeRefreshTimer.Start()
    })
    $Form.Add_Activated({ [void](Refresh-ConfiguredAttributes); [void](Refresh-ConfiguredOUs) })
    $Form.Add_FormClosed({
        $AttributeRefreshTimer.Stop()
        $AttributeRefreshTimer.Dispose()
    })
    $Form.Add_FormClosed({ $PrerequisiteToolTip.Dispose() })

    $Result =
        $Form.ShowDialog()

    if (
        $Result -eq
        [System.Windows.Forms.DialogResult]::OK
    ) {

        return $Form.Tag
    }

    return $null
}

# ============================================================
# SHOW GUI
# ============================================================

$Configuration =
    Show-ImportConfiguration

if (-not $Configuration) {
    exit 0
}

$ARServer =
    $Configuration.ARServer

$DomainDN =
    $Configuration.DomainDN

$CSVPath =
    $Configuration.CSVPath

$OU =
    $Configuration.OU

$ReportPath =
    $Configuration.ReportPath

# ============================================================
# IMPORT CSV
# ============================================================

try {

    $Computers =
        Import-Csv `
            -Path $CSVPath `
            -ErrorAction Stop
}
catch {

    [System.Windows.Forms.MessageBox]::Show(
        "Unable to read CSV.`r`n`r`n$($_.Exception.Message)",
        "CSV Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )

    exit 1
}

if (
    -not $Computers -or
    -not (
        $Computers[0].PSObject.Properties.Name `
        -contains "ComputerName"
    )
) {

    [System.Windows.Forms.MessageBox]::Show(
        "CSV must contain a ComputerName column.",
        "Invalid CSV",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )

    exit 1
}

$ComputerNames = @(
    $Computers |
        ForEach-Object {

            if ($_.ComputerName) {

                $_.ComputerName.Trim()
            }
        } |
        Where-Object {

            -not [string]::IsNullOrWhiteSpace(
                $_
            )
        }
)

if ($ComputerNames.Count -eq 0) {

    [System.Windows.Forms.MessageBox]::Show(
        "No valid computer names were found in the CSV.",
        "Empty CSV",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    exit 1
}

# ============================================================
# CONFIRMATION
# ============================================================

$AttributeText =
    Get-AttributeDisplayText

$Preview =
    (
        $ComputerNames |
        Select-Object -First 10
    ) -join "`r`n"

if ($ComputerNames.Count -gt 10) {

    $Preview +=
        "`r`n...and $($ComputerNames.Count - 10) more"
}

$ConfirmText = @"
Active Roles Server:
$ARServer

Domain:
$DomainDN

Computers:
$($ComputerNames.Count)

Destination OU:
$OU

Configured Attributes:
$AttributeText

Computer Preview:
$Preview

Continue?
"@

$Confirm =
    [System.Windows.Forms.MessageBox]::Show(
        $ConfirmText,
        "Confirm Bulk Import",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

if (
    $Confirm -ne
    [System.Windows.Forms.DialogResult]::Yes
) {

    exit 0
}

# ============================================================
# CONNECT ACTIVE ROLES
# ============================================================

$QADConnection = $null

try {

    $QADConnection =
        Connect-QADService `
            -Service $ARServer `
            -Proxy `
            -ErrorAction Stop
}
catch {

    [System.Windows.Forms.MessageBox]::Show(
        "Unable to connect to Active Roles.`r`n`r`n$($_.Exception.Message)",
        "Active Roles Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )

    exit 1
}

# ============================================================
# PROCESS COMPUTERS
# ============================================================

$Results = @()

try {

    foreach ($Entry in $Computers) {

        if (-not $Entry.ComputerName) {
            continue
        }

        $ComputerName =
            $Entry.ComputerName.Trim()

        if (
            [string]::IsNullOrWhiteSpace(
                $ComputerName
            )
        ) {

            continue
        }

        try {

            $ExistingComputer =
                Get-QADComputer `
                    -Identity $ComputerName `
                    -Connection $QADConnection `
                    -ErrorAction SilentlyContinue

            if ($ExistingComputer) {

                Set-QADComputer `
                    -Identity $ExistingComputer `
                    -ObjectAttributes $script:ComputerObjectAttributes `
                    -Connection $QADConnection `
                    -ErrorAction Stop

                $Results +=
                    [PSCustomObject]@{
                        ComputerName      = $ComputerName
                        Status            = "Already Existed - Updated"
                        Location          = $ExistingComputer.DN
                        AttributesApplied = $script:ComputerObjectAttributes.Count
                        Error             = ""
                    }

                continue
            }

            $NewComputer =
                New-QADComputer `
                    -Name $ComputerName `
                    -SamAccountName ($ComputerName + '$') `
                    -ParentContainer $OU `
                    -ObjectAttributes $script:ComputerObjectAttributes `
                    -Connection $QADConnection `
                    -ErrorAction Stop

            $Results +=
                [PSCustomObject]@{
                    ComputerName      = $ComputerName
                    Status            = "Created"
                    Location          = $OU
                    AttributesApplied = $script:ComputerObjectAttributes.Count
                    Error             = ""
                }
        }
        catch {

            $Results +=
                [PSCustomObject]@{
                    ComputerName      = $ComputerName
                    Status            = "FAILED"
                    Location          = $OU
                    AttributesApplied = 0
                    Error             = $_.Exception.Message
                }
        }
    }
}
finally {

    if ($QADConnection) {

        Disconnect-QADService `
            -Connection $QADConnection `
            -ErrorAction SilentlyContinue
    }
}

# ============================================================
# REPORT
# ============================================================

$LogPath =
    Join-Path `
        $ReportPath `
        "ComputerImportResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

try {

    $Results |
        Export-Csv `
            -Path $LogPath `
            -NoTypeInformation `
            -ErrorAction Stop
}
catch {

    [System.Windows.Forms.MessageBox]::Show(
        "The import completed, but the results report could not be written.`r`n`r`n$($_.Exception.Message)",
        "Report Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )

    exit 1
}

# ============================================================
# SUMMARY
# ============================================================

$CreatedCount =
    @(
        $Results |
        Where-Object {
            $_.Status -eq "Created"
        }
    ).Count

$UpdatedCount =
    @(
        $Results |
        Where-Object {
            $_.Status -eq
            "Already Existed - Updated"
        }
    ).Count

$FailedCount =
    @(
        $Results |
        Where-Object {
            $_.Status -eq "FAILED"
        }
    ).Count

$Summary = @"
Bulk computer import complete.

Created:
$CreatedCount

Existing / Updated:
$UpdatedCount

Failed:
$FailedCount

Configured Attributes Applied:
$($script:ComputerObjectAttributes.Count)

$AttributeText

Report:
$LogPath
"@

[System.Windows.Forms.MessageBox]::Show(
    $Summary,
    "Active Roles Import Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

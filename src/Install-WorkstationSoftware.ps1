$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')

if ($isAdmin -eq $false) {
    Throw "Must be executed from an Administrator elevated session"
}

Function Set-PowerShell {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $execPolicy = Get-ExecutionPolicy
    if ($execPolicy -ne 'RemoteSigned') {
        Write-Host "Set PowerShell ExecutionPolicy to RemoteSigned"
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force
    }

    Write-Host "Verify latest NuGet version installed."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force

    $galleryInstallationPolicy = (Get-PSRepository -Name PSGallery).InstallationPolicy
    if ($galleryInstallationPolicy -ne 'Trusted') {
        Write-Host "Set PowerShell to Trust PowerShell Gallery as an install source"
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

    Write-Host "Checking PowerShell Gallery to verify what is the latest version of PowerShellGet."
    [version]$currentVersionPowerShellGet = (Find-Module -Name PowerShellGet).Version
    [version]$installedVersionPowerShellGet = (Get-Module -Name PowerShellGet -ListAvailable).Version[0]  #Gets latest version if there are multiple installed

    if ($currentVersionPowerShellGet -eq $installedVersionPowerShellGet) {
        Write-Host "The current version of PowerShellGet ($($currentVersionPowerShellGet)) is already installed."
    }
    else {
        Write-Host "An older version of PowerShellGet is installed, updating $($installedVersionPowerShellGet) to $($currentVersionPowerShellGet)"
        Install-Module -Name PowerShellGet -Force -AllowClobber -WarningAction SilentlyContinue
    }
}

Function Install-WinGet {
    Write-Host "Checking GitHub to verify what is the latest version of winget."

    # Get latest download url for WinGet
    $asset = Invoke-RestMethod -Method Get -Uri 'https://api.github.com/repos/microsoft/winget-cli/releases/latest' | ForEach-Object assets | Where-Object name -like "*.msixbundle"

    $InstallerWinGet = $env:TEMP + "\$($asset.name)"
    $currentVersionWinGet = $asset.browser_download_url | Select-String '(\bv?(?:\d+\.){2}\d+)' | ForEach-Object { $_.Matches[0].Groups[1].Value }

    Try {
        $installedVersionWinGet = Invoke-Expression 'winget --version'
    }
    Catch {
        $installedVersionWinGet = $null
    }

    if ($currentVersionWinGet -le $installedVersionWinGet) {
        Write-Host "The current version of winget ($($currentVersionWinGet)) is already installed"
    }
    elseif ($currentVersionWinGet -gt $installedVersionWinGet -and $null -ne $installedVersionWinGet) {
        Write-Host "An older version of winget is installed, updating $($installedVersionWinGet) to $($currentVersionWinGet)"
        $progresspreference = 'silentlyContinue'
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $InstallerWinGet
        Add-AppxPackage -Path $InstallerWinGet -Update
        $progressPreference = 'Continue'
    }
    else {
        Write-Host "Installing winget $($currentVersionWinGet)"

        $progresspreference = 'silentlyContinue'
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $InstallerWinGet
        Add-AppxPackage -Path $InstallerWinGet
        $progressPreference = 'Continue'
    }

    if (Test-Path -Path "$InstallerWinGet") {
        Remove-Item $InstallerWinGet -Force -ErrorAction SilentlyContinue
    }
}

Function Install-WinGetSoftware {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Id
    )

    if ((Invoke-Expression "winget list --exact --id $Id --accept-source-agreements") -eq 'No installed package found matching input criteria.') {
        Invoke-Expression "winget install --exact --id $Id"
    }
}

Function Join-PrependIdempotent {

    # the delimiter is expected to be just 1 unique character
    # otherwise there may be problems with trimming
    param (
        [string]$InputString,
        [string]$OriginalString,
        [string]$Delimiter = '',
        [bool]$CaseSensitive = $false
    )

    if ($CaseSensitive -and ("$OriginalString" -cnotlike "*${InputString}*")) {

        "$InputString".TrimEnd("$Delimiter") + "$Delimiter" + "$OriginalString".TrimStart("$Delimiter")

    }
    elseif (! $CaseSensitive -and ("$OriginalString" -inotlike "*${InputString}*")) {

        "$InputString".TrimEnd("$Delimiter") + "$Delimiter" + "$OriginalString".TrimStart("$Delimiter")

    }
    else {

        "$OriginalString"

    }

}

Function Join-AppendIdempotent {

    # the delimiter is expected to be just 1 unique character
    # otherwise there may be problems with trimming
    param (
        [string]$InputString,
        [string]$OriginalString,
        [string]$Delimiter = '',
        [bool]$CaseSensitive = $false
    )

    if ($CaseSensitive -and ("$OriginalString" -cnotlike "*${InputString}*")) {

        "$OriginalString".TrimEnd("$Delimiter") + "$Delimiter" + "$InputString".TrimStart("$Delimiter")

    }
    elseif (! $CaseSensitive -and ("$OriginalString" -inotlike "*${InputString}*")) {

        "$OriginalString".TrimEnd("$Delimiter") + "$Delimiter" + "$InputString".TrimStart("$Delimiter")

    }
    else {

        "$OriginalString"

    }

}

Function Add-Path {

    param (
        [string]$NewPath,
        [ValidateSet('Prepend', 'Append')]$Style = 'Prepend',
        [ValidateSet('User', 'System')]$Target = 'User'
    )

    try {

        # we need to do this to make sure not to expand the environment variables already inside the PATH
        if ($Target -eq 'User') {
            $Key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Environment', $true)
        }
        elseif ($Target -eq 'System') {
            $Key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SYSTEM\CurrentControlSet\Control\Session Manager\Environment', $true)
        }

        $Path = $Key.GetValue('Path', $null, 'DoNotExpandEnvironmentNames')

        # note that system path can only expand system environment variables and vice versa for user environment variables
        # in order to make sure this method is idempotent, we need to check if the new path already exists, this requires having a semicolon at the very end
        if ($Style -eq 'Prepend') {
            $key.SetValue('Path', (Join-PrependIdempotent ("$NewPath".TrimEnd(';') + ';') ("$Path".TrimEnd(';') + ';') ';' $false), 'ExpandString')
        }
        elseif ($Style -eq 'Append') {
            $key.SetValue('Path', (Join-AppendIdempotent ("$NewPath".TrimEnd(';') + ';') ("$Path".TrimEnd(';') + ';') ';' $false), 'ExpandString')
        }

        # update the path for the current process as well
        #$Env:Path = $key.GetValue('Path', $null)

    }
    finally {
        $key.Dispose()
    }
}

Function Invoke-SoftwareInstallProcess {
    param (
        [Parameter(Mandatory = $true)]
        $Checkboxes
    )

    # Count Software to Install
    $numberSoftware = 0

    ForEach ($checkbox in $Checkboxes) {
        if ($checkbox.Checked -eq $true) {
            $numberSoftware ++
        }
    }

    # Initialize Progress Bar
    [int]$percentIncrement = 100 / $numberSoftware
    [int]$percentCurrent = 0

    Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Initializing"

    # Process Selected Software and PowerShell Module Installs
    ForEach ($checkbox in $Checkboxes) {
        #----- Software Installs -----
        if ($checkbox.Text -eq 'Edge' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Edge"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'Microsoft.Edge'
        }

        if ($checkbox.Text -eq 'Firefox' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Firefox"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'Mozilla.Firefox'
        }

        if ($checkbox.Text -eq 'Chrome' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Chrome"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'Google.Chrome'
        }

        if ($checkbox.Text -eq 'PowerShell 7' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing PowerShell 7"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'Microsoft.PowerShell'
        }

        if ($checkbox.Text -eq 'Visual Studio Code' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Visual Studio Code"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'Microsoft.VisualStudioCode'
        }

        if ($checkbox.Text -eq 'Git Client' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Git Client"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'Git.Git'
        }

        if ($checkbox.Text -eq 'SQL Server Management Studio' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing SQL Server Management Studio"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'Microsoft.SQLServerManagementStudio'
        }

        if ($checkbox.Text -eq 'Active Directory Users and Computers' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Active Directory Users and Computers"
            $percentCurrent = $percentCurrent + $percentIncrement

            $capabilityName = (Get-WindowsCapability -Online | Where-Object { $_.Name -like 'Rsat.ActiveDirectory.DS-LDS.Tools*' }).Name

            if ((Get-WindowsCapability -Online -Name $capabilityName).State -ne 'Installed') {
                Add-WindowsCapability -Online -Name $capabilityName
            }
        }

        if ($checkbox.Text -eq 'Telnet Client' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Telnet Client"
            $percentCurrent = $percentCurrent + $percentIncrement

            if ((Get-WindowsOptionalFeature -Online -FeatureName TelnetClient).State -ne 'Enabled') {
                Enable-WindowsOptionalFeature -Online -FeatureName TelnetClient
            }
        }

        if ($checkbox.Text -eq 'Windows Terminal' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Windows Terminal"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'Microsoft.WindowsTerminal'
        }

        if ($checkbox.Text -eq 'Notepad++' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Notepad++"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'Notepad++.Notepad++'
        }

        if ($checkbox.Text -eq 'pgAdmin' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing pgAdmin"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'PostgreSQL.pgAdmin'
        }

        if ($checkbox.Text -eq 'PuTTY' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing PuTTY"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'PuTTY.PuTTY'
        }

        if ($checkbox.Text -eq 'VcXsrv Windows X Server' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing VcXsrv Windows X Server"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-WinGetSoftware -Id 'marha.VcXsrv'
        }




        #----- PowerShell Modules -----
        if ($checkbox.Text -eq 'Az' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Az Module"
            $percentCurrent = $percentCurrent + $percentIncrement

            Write-Host "The Az Module Install can take a very long time to start, as it gathers many dependencies.  Please be patient and do not cancel."
            Install-Module -Name Az -Force
        }

        if ($checkbox.Text -eq 'DbaTools' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing DbaTools Module"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-Module -Name DbaTools -Force
        }

        if ($checkbox.Text -eq 'SqlServer' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing SqlServer Module"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-Module -Name SqlServer -Force
        }

        if ($checkbox.Text -eq 'Pester' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing Pester Module"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-Module -Name Pester -Force -SkipPublisherCheck -WarningAction SilentlyContinue
        }

        if ($checkbox.Text -eq 'ImportExcel' -and $checkbox.Checked -eq $true) {
            Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installing ImportExcel Module"
            $percentCurrent = $percentCurrent + $percentIncrement

            Install-Module -Name ImportExcel -Force
        }
    }

    $percentCurrent = 100
    Write-Progress -Activity 'Install Software' -Status "$percentCurrent% Complete:" -PercentComplete $percentCurrent -CurrentOperation "Installs Complete"
}

Function Show-Form {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    $checkboxes = @()

    # Set the size of your form
    $Form = New-Object System.Windows.Forms.Form
    #$Form.AutoScaleMode = "Dpi"
    $Form.Width = 500
    $Form.Height = 500
    $Form.AutoSize = $true
    $Form.Text = "Workstation Software Install"
    $Form.StartPosition = "CenterScreen"
    $Form.TopMost = $false
    $Form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#b8b8b8")
    $Form.FormBorderStyle = 'FixedSingle'

    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Calibri Bold", 12)
    $Form.Font = $Font

    # Software groupbox
    $groupboxSoftware = New-Object System.Windows.Forms.groupbox
    $groupboxSoftware.Location = new-object System.Drawing.Point(30, 10)
    $groupboxSoftware.AutoSize = $true
    $groupboxSoftware.Text = "Software"
    $Form.Controls.Add($groupboxSoftware)

    # Edge checkbox
    $checkboxEdge = new-object System.Windows.Forms.checkbox
    $checkboxEdge.Location = new-object System.Drawing.Point(15, 30)
    $checkboxEdge.AutoSize = $true
    $checkboxEdge.Text = "Edge"
    $checkboxEdge.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxEdge)
    $checkboxes += $checkboxEdge

    # Firefox checkbox
    $checkboxFirefox = new-object System.Windows.Forms.checkbox
    $checkboxFirefox.Location = new-object System.Drawing.Point(15, 60)
    $checkboxFirefox.AutoSize = $true
    $checkboxFirefox.Text = "Firefox"
    $checkboxFirefox.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxFirefox)
    $checkboxes += $checkboxFirefox

    # Chrome checkbox
    $checkboxChrome = new-object System.Windows.Forms.checkbox
    $checkboxChrome.Location = new-object System.Drawing.Point(15, 90)
    $checkboxChrome.AutoSize = $true
    $checkboxChrome.Text = "Chrome"
    $checkboxChrome.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxChrome)
    $checkboxes += $checkboxChrome

    # PowerShell 7 checkbox
    $checkboxPowerShell = new-object System.Windows.Forms.checkbox
    $checkboxPowerShell.Location = new-object System.Drawing.Point(15, 120)
    $checkboxPowerShell.AutoSize = $true
    $checkboxPowerShell.Text = "PowerShell 7"
    $checkboxPowerShell.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxPowerShell)
    $checkboxes += $checkboxPowerShell

    # Visual Stuido Code checkbox
    $checkboxVsCode = new-object System.Windows.Forms.checkbox
    $checkboxVsCode.Location = new-object System.Drawing.Point(15, 150)
    $checkboxVsCode.AutoSize = $true
    $checkboxVsCode.Text = "Visual Studio Code"
    $checkboxVsCode.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxVsCode)
    $checkboxes += $checkboxVsCode

    # Git checkbox
    $checkboxGit = new-object System.Windows.Forms.checkbox
    $checkboxGit.Location = new-object System.Drawing.Point(15, 180)
    $checkboxGit.AutoSize = $true
    $checkboxGit.Text = "Git Client"
    $checkboxGit.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxGit)
    $checkboxes += $checkboxGit

    # SQL Server Management Studio checkbox
    $checkboxSSMS = new-object System.Windows.Forms.checkbox
    $checkboxSSMS.Location = new-object System.Drawing.Point(15, 210)
    $checkboxSSMS.AutoSize = $true
    $checkboxSSMS.Text = "SQL Server Management Studio"
    $checkboxSSMS.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxSSMS)
    $checkboxes += $checkboxSSMS

    # Active Directory Users and Computers checkbox
    $checkboxAD = new-object System.Windows.Forms.checkbox
    $checkboxAD.Location = new-object System.Drawing.Point(15, 240)
    $checkboxAD.AutoSize = $true
    $checkboxAD.Text = "Active Directory Users and Computers"
    $checkboxAD.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxAD)
    $checkboxes += $checkboxAD

    # Telnet Client checkbox
    $checkboxTelnet = new-object System.Windows.Forms.checkbox
    $checkboxTelnet.Location = new-object System.Drawing.Point(15, 270)
    $checkboxTelnet.AutoSize = $true
    $checkboxTelnet.Text = "Telnet Client"
    $checkboxTelnet.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxTelnet)
    $checkboxes += $checkboxTelnet

    # Windows Terminal checkbox
    $checkboxWindowsTerminal = new-object System.Windows.Forms.checkbox
    $checkboxWindowsTerminal.Location = new-object System.Drawing.Point(15, 300)
    $checkboxWindowsTerminal.AutoSize = $true
    $checkboxWindowsTerminal.Text = "Windows Terminal"
    $checkboxWindowsTerminal.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxWindowsTerminal)
    $checkboxes += $checkboxWindowsTerminal

    # Notepad++ checkbox
    $checkboxNotepadPlus = new-object System.Windows.Forms.checkbox
    $checkboxNotepadPlus.Location = new-object System.Drawing.Point(15, 330)
    $checkboxNotepadPlus.AutoSize = $true
    $checkboxNotepadPlus.Text = "Notepad++"
    $checkboxNotepadPlus.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxNotepadPlus)
    $checkboxes += $checkboxNotepadPlus

    # pgAdmin checkbox
    $checkboxpgAdmin = new-object System.Windows.Forms.checkbox
    $checkboxpgAdmin.Location = new-object System.Drawing.Point(15, 360)
    $checkboxpgAdmin.AutoSize = $true
    $checkboxpgAdmin.Text = "pgAdmin"
    $checkboxpgAdmin.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxpgAdmin)
    $checkboxes += $checkboxpgAdmin

    # Putty checkbox
    $checkboxPutty = new-object System.Windows.Forms.checkbox
    $checkboxPutty.Location = new-object System.Drawing.Size(15, 390)
    $checkboxPutty.AutoSize = $true
    $checkboxPutty.Text = "PuTTY"
    $checkboxPutty.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxPutty)
    $checkboxes += $checkboxPutty

    # VcXsrv checkbox
    $checkboxVcXsrv = new-object System.Windows.Forms.checkbox
    $checkboxVcXsrv.Location = new-object System.Drawing.Size(15, 420)
    $checkboxVcXsrv.AutoSize = $true
    $checkboxVcXsrv.Text = "VcXsrv Windows X Server"
    $checkboxVcXsrv.Checked = $false
    $groupboxSoftware.Controls.Add($checkboxVcXsrv)
    $checkboxes += $checkboxVcXsrv




    # PowerShell Module groupbox
    $groupboxPsModule = New-Object System.Windows.Forms.groupbox
    $groupboxPsModule.Location = new-object System.Drawing.Point(($groupboxSoftware.Right + 50), 10)
    $groupboxPsModule.AutoSize = $true
    $groupboxPsModule.Text = "PowerShell Modules"
    $Form.Controls.Add($groupboxPsModule)

    # Az checkbox
    $checkboxAz = new-object System.Windows.Forms.checkbox
    $checkboxAz.Location = new-object System.Drawing.Size(15, 30)
    $checkboxAz.AutoSize = $true
    $checkboxAz.Text = "Az"
    $checkboxAz.Checked = $false
    $groupboxPsModule.Controls.Add($checkboxAz)
    $checkboxes += $checkboxAz

    # DbaTools checkbox
    $checkboxDbaTools = new-object System.Windows.Forms.checkbox
    $checkboxDbaTools.Location = new-object System.Drawing.Size(15, 60)
    $checkboxDbaTools.AutoSize = $true
    $checkboxDbaTools.Text = "DbaTools"
    $checkboxDbaTools.Checked = $false
    $groupboxPsModule.Controls.Add($checkboxDbaTools)
    $checkboxes += $checkboxDbaTools

    # SqlServer checkbox
    $checkboxSqlServer = new-object System.Windows.Forms.checkbox
    $checkboxSqlServer.Location = new-object System.Drawing.Size(15, 90)
    $checkboxSqlServer.AutoSize = $true
    $checkboxSqlServer.Text = "SqlServer"
    $checkboxSqlServer.Checked = $false
    $groupboxPsModule.Controls.Add($checkboxSqlServer)
    $checkboxes += $checkboxSqlServer

    # Pester checkbox
    $checkboxPester = new-object System.Windows.Forms.checkbox
    $checkboxPester.Location = new-object System.Drawing.Size(15, 120)
    $checkboxPester.AutoSize = $true
    $checkboxPester.Text = "Pester"
    $checkboxPester.Checked = $false
    $groupboxPsModule.Controls.Add($checkboxPester)
    $checkboxes += $checkboxPester

    # ImportExcel checkbox
    $checkboxImportExcel = new-object System.Windows.Forms.checkbox
    $checkboxImportExcel.Location = new-object System.Drawing.Size(15, 150)
    $checkboxImportExcel.AutoSize = $true
    $checkboxImportExcel.Text = "ImportExcel"
    $checkboxImportExcel.Checked = $false
    $groupboxPsModule.Controls.Add($checkboxImportExcel)
    $checkboxes += $checkboxImportExcel


    # Adjust Form Size to accomodate all objects
    $Form.Width = $groupboxPsModule.Right + 30
    $Form.Height = $groupboxSoftware.Bottom + 75 + $OKButton.Height + 50

    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Size = new-object System.Drawing.Size(100, 40)
    $OKButton.Text = "OK"
    $OKButton.Add_Click( { $Form.Visible = $false; Invoke-SoftwareInstallProcess -Checkboxes $checkboxes; $Form.Close() })
    $form.Controls.Add($OKButton)

    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Size = new-object System.Drawing.Size(100, 40)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click( { $Form.Close() })
    $Form.Controls.Add($CancelButton)

    # Adjust Form Size to accomodate all objects
    $Form.Width = $groupboxPsModule.Right + 30
    $Form.Height = $groupboxSoftware.Bottom + 60 + $OKButton.Height + 30

    # Adjust Button positions
    $CancelButton.Location = new-object System.Drawing.Size(($Form.Right - 30 - $CancelButton.Width), ($Form.Bottom - 90))
    $OKButton.Location = new-object System.Drawing.Size(($CancelButton.Left - 15 - $OKButton.Width), $CancelButton.Location.Y)

    # Activate/Show the form
    $Form.Add_Shown( { $Form.Activate() })
    [void] $Form.ShowDialog()
}




#region Main
Write-Host "Checking Dependencies ..."
Set-PowerShell
Install-WinGet

Show-Form
#endregion Main

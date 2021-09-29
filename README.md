# Install-WorkstationSoftware

## Requirements

- Internet access without a proxy.
- You must be Administrator on the VM.
- You must execute the script from an Administrator PowerShell session.

## Instructions

- Open an Administrator PowerShell session.
- Execute the following command

```PowerShell
Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/TrisBits/Install-WorkstationSoftware/main/src/Install-WorkstationSoftware.ps1'))
```

## Alternate Instructions

- Download and unzip the script.
- In an Administrator PowerShell session browse to the location of Install-WorkstationSoftware.ps1
- Execute the script using the following command.  If you recieve the error "not digitally signed" you will first need to execute the command **Unblock-File -Path .\Install-WorkstationSoftware.ps1**

```PowerShell
.\Install-WorkstationSoftware.ps1
```

## Software Currently Supported

You will be presented with checkboxes to select from the following software.

- Edge
- Firefox
- Chrome
- PowerShell 7
- Visual Studio Code
- Git Client
- SQL Server Management Studio
- Active Directory Users and Computers
- Telnet Client
- Windows Terminal
- Notepad++
- pgAdmin
- PuTTY
- VcXsrv Windows X Server

## PowerShell Modules Currently Supported

You will be presented with checkboxes to select from the following modules.

- Az
- DbaTools
- SqlServer
- Pester
- ImportExcel

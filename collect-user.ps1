<#
    .SYNOPSIS
        Collect User Settings Information
    .DESCRIPTION
        A script that collects various bits of user information
    .EXAMPLE
        PS C:\> collect-user.ps1
    .INPUTS
        None
    .OUTPUTS
        To console, logs
    .NOTES
        By Nathan DeGruchy <nathan@degruchy.org>
    .VERSION
        1.2.0
#>

##
# Version: 1.2.0
# Changelog:
#  - Changed Outlook detection code to log to the same place
#
# Version: 1.0.0
# Changelog:
#  - Initial working script
##

##
# Global Scope Variables
##
$version                    = 1.2.0
$whoami                     = ( whoami )
$username                   = Split-Path $whoami -Leaf
$domain                     = Split-Path $whoami
$CurrentDate                = ( Get-Date -F yyyy-MM-dd )
# $Guid                       = New-Guid
$ScriptLog                  = "${CurrentDate}-CollectionLog-${ENV:COMPUTERNAME}-${username}.log"
$ConfigFolder               = "Config"
$WorkingDirectory           = "${ENV:OneDrive}\${ConfigFolder}\"
$InstalledPrograms          = ( Get-CimInstance -Query "SELECT * FROM Win32_Product" ) | Select-Object Name, Version
$Printers                   = ( Get-Printer )
$MappedDrives               = ( Get-SmbMapping )
$ProxySettings              = ( Get-ItemProperty -Path "Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings" )
$BogusPrinters              = @(
    "Adobe PDF",
    "Microsoft XPS Document Writer",
    "Microsoft Print to PDF",
    "Fax",
    "Send To OneNote 2016",
    "OneNote (Desktop)",
    "OneNote for Windows 10"
)
$FileList                   = @(
    "${ENV:LOCALAPPDATA}\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"
    # Add more here
)

##
# Functions
##

function Write-Log {
    <#
        .SYNOPSIS
            Generic logging function. Takes a type and message and logs it to a file.
        .PARAMETER Type
            The type of message your want to log. One of "Info", "Warn" or "Error"
        .PARAMETER Message
            The message you would like to log.
        .EXAMPLE
            Informational:
            Write-Log -Type "Info" -Message "Hello World!"
            Write-Log "Info" "Hello World"

            Warning:
            Write-Log -Type "Warn" -Message "Something looks screwy here."
            Write-Log "Warn" "Something looks screwy here."

            Error:
            Write-Log -Type "Error" -Message "I'm sorry, Dave. I'm afraid I can't do that."
            Write-Log "Error" "I'm sorry, Dave. I'm afraid I can't do that."
    #>
    param (
        [ValidateSet( "Info", "Warn", "Error" )]
        [String]
        $Type = "Info",

        [String]
        [Parameter( Mandatory )] $Message
    )
    if ( $LogLevel -eq 0 )
    {
        Return;
    }
    else
    {
        $TimeDate = Get-Date -Format "o"
        Add-Content -Path "${WorkingDirectory}\${ScriptLog}" -Value "${TimeDate}  ${Type}: ${Message}"
    }
}

function Get-LocalVersions {
    <#
        .SYNOPSIS
            Finds all the installed programs on a system and lists their name and version number.
    #>
    param ()
    foreach( $program in $InstalledPrograms )
    {
        if( ( [String]$program.Name -eq "" ) -or ( [String]$program.Version -eq "" ) )
        {
            continue
        }
        Write-Log "Info" "Program: $($program.Name); Version: $($program.Version)"
    }
}

function Get-OutlookPSTs
{
    <#
        .SYNOPSIS
            Opens Outlook via COM and enumerates the attached PSTs and OSTs
    #>
    param
    (
        [Parameter( Mandatory=$false,ValueFromPipeline=$true )]
        [ ValidateSet( "ALL","PST","OST" ) ]
        [string]
        $Extension = "ALL"
    )

    ##
    # Outlook Object
    ##
    $Outlook = New-Object -ComObject Outlook.Application

    try
    {
        if( $null -ne $Outlook )
        {
            foreach ( $store in $Outlook.Session.Stores )
            {
                if( $store.ExchangeStoreType -eq 3 )
                {
                    Write-Log "Info" "Outlook PST: `"$($store.DisplayName)`" found at $($store.FilePath)"
                }
            }
        }
    }
    catch
    {
        Write-Log "Error" "Outlook PST: Unable to get PST Information"
    }
    finally
    {
        [Runtime.InteropServices.Marshal]::ReleaseComObject( $Outlook ) | Out-Null
    }
}


function Get-UserPrinters
{
    <#
        .SYNOPSIS
            Pulls a list of user printers, to from what device their mapped from (blank if direct), the port (typically the IP)
            and the driver being used.
    #>
    param()
    $count = 0

    if ( $Printers.length -gt 0 )
    {
        foreach ( $printer in $Printers )
        {
            if ( $printer.Name -in $BogusPrinters )
            {
                continue
            }
            else
            {
                Write-Log "Info" "Printer: `"$($printer.Name)`", on $($printer.ComputerName) port $($printer.PortName) using $($printer.DriverName)"
                $count = $count + 1
            }
        }
        if ($count -eq 0)
        {
            Write-Log "Info" "Printer: No user printers found"
        }
    }
    else
    {
        Write-Log "Info" "Printer: No connected printers."
    }


}

function Get-UserMappedDrives
{
    <#
        .SYNOPSIS
            Pulls a list of the user's mapped drives and their letters.
    #>
    param()
    if ( $MappedDrives.length -gt 0 )
    {
        foreach ( $drive in $MappedDrives )
        {
            Write-Log "Info" "Drive: $($drive.LocalPath) is mapped to $($drive.RemotePath)"
        }
    }
    else
    {
        Write-Log "Info" "Drive: No mapped drives found"
    }
}


function Get-UserProxySettings {
    param ()
    $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections'
    $data = (Get-ItemProperty -Path $key -Name DefaultConnectionSettings).DefaultConnectionSettings


    if ( $ProxySettings.ProxyEnable -eq 1 )
    {
        Write-Log "Info" "Proxy is enabled with a custom setting: $($ProxySettings.ProxyServer)"
    }

    <#
        'Automatically detect settings' values
        --------------------------------------------
        | Proxy State    |  Checked  |  Unchecked  |
        -----------------+-----------+--------------
        | Proxy enabled  |    11     |      3      |
        -----------------+-----------+--------------
        | Proxy disabled |     9     |      1      |
        --------------------------------------------
    #>

    if ( $data[8] -ne 9 )
    # Looking for the value at array key '8'.
    {
        Write-Log "Warn" "Proxy: Autodetect proxy setting is un-set!"
    }
    else
    {
        Write-Log "Info" "Proxy: Proxy is set to autodetect"
    }
}

function Get-UserHostsSettings
{
    param ()
    $Pattern    = '^(?<IP>\d{1,3}(\.\d{1,3}){3})\s+(?<Host>.+)$'
    $File       = "${env:SystemDrive}\Windows\System32\Drivers\etc\hosts"
    $Contents   = Get-Content -Path $File
    $Count      = 0

    ForEach ($line in $Contents)
    {
        if ($line -match $Pattern) {
            Write-Log "Info" "Hosts file entry: $($Matches.IP), $($Matches.Host)"
            $Count++
        }
    }
    if( $Count -eq 0 )
    {
        Write-Log "Info" "Hosts file entry: No entries found."
    }
}

function Get-UserMiscFiles
{
    <#
        .SYNOPSIS
            Copies misc user files to a single directory, then compresses them.
    #>
    param ()

    Write-Log "Info" "Looking for files to back up"

    if(-NOT $FileList.length -eq 0)
    {
        # Todo copy files to temp folder and zip
        foreach ($file in $FileList)
        {
            if( Test-Path -PathType Leaf -Path $file )
            {
                Copy-Item -Path $file -Destination $WorkingDirectory
                Write-Log "Info" "Misc Files: Copied ${file} to ${WorkingDirectory}"
            } else {
                Write-Log "Info" "Misc Files: File ${file} not found, skipping"
            }
        }
    }
    else
    {
        Write-Log "Info" "No backup files found."
        return
    }
}

##
# Main Program
##

Write-Progress -Activity "Collecting User Custom Information" -Status "Collecting Basic Information" -PercentComplete 10

if ( -not ( Test-Path $WorkingDirectory ) )
{
    New-Item -Path $WorkingDirectory -ItemType Directory | Out-Null
    Write-Log "Info" "Created working directory ${WorkingDirectory}"
}

Write-Log "Info" "Collecting User Custom Information - Script Version ${version}"
Write-Log "Info" "Today's Date is ${CurrentDate}"
Write-Log "Info" "User is reported as ${username} from domain ${domain}"

Write-Progress -Activity "Collecting User Custom Information" -Status "Collecting Printer Information" -PercentComplete 20

Get-UserPrinters

Write-Progress -Activity "Collecting User Custom Information" -Status "Collecting Installed Program Information" -PercentComplete 30

Get-LocalVersions

Write-Progress -Activity "Collecting User Custom Information" -Status "Collecting Connected Mapped Drives" -PercentComplete 40

Get-UserMappedDrives

Write-Progress -Activity "Collecting User Custom Information" -Status "Collecting Information About Attached PST Files in Outlook" -PercentComplete 50

Get-OutlookPSTs

Write-Progress -Activity "Collecting User Custom Information" -Status "Collecting Information About Proxies and Custom Hosts" -PercentComplete 60

Get-UserProxySettings
Get-UserHostsSettings

Write-Progress -Activity "Collecting User Custom Information" -Status "Collecting Misc Files That May Have Been Missed" -PercentComplete 70

Get-UserMiscFiles

Write-Progress -Activity "Collecting User Custom Information" -Status "All set!" -PercentComplete 100 -Completed

Write-Log "Info" "Complete!"
Write-Host "User information collected to: ${WorkingDirectory}${ScriptLog}"

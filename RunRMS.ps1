<#PSScriptInfo

.VERSION 23.02.26

.GUID b33d6589-1851-4759-bdfb-241f0e44db33

.AUTHOR 

Chris Kelleher W3CTK w3ctk@delcoares.net

.COMPANYNAME 

ARES of Delaware County, Emergency Coordinator

.COPYRIGHT

Copyright 2023 Christopher Kelleher

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

.TAGS

.LICENSEURI     

http://www.apache.org/licenses/LICENSE-2.0

.PROJECTURI     

https://github.com/w3ctk/RunRMS 

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

AudioDeviceCmdlets

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

23.02.26 - Added function to load and install required modules when needed

.PRIVATEDATA

#>

<# 

.DESCRIPTION 

https://github.com/w3ctk/RunRMS

This is a script to check and run the various RMS sysop processes on Windows

To run, you can launch Windows PowerShell and run the command from 
the directory you downloaded the RunRMS.ps1 script into:

./runrms

You can download PowerShell for Windows here if it isn't already installed:

https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows

All of the allowed parameters are set in the script, and these are
displayed when the script is run.  To override any of these settings,
you can add them into the command line parameters, like

./runrms -RMSPacketDirectory "C:\RMS Packet\" -RestartRMSPacket $True -LogFileEmail $True

========= CHECKING PROGRAM PARAMETERS =====================
========= PARAMETERS FOR LOGGING ==========================
LogFile:              RunRMS.log
EventLogger:          True
EventSource:          RunRMS
LogFileWrite:         True
LogFileReset:         True
LogFileEmail:         False
IncludeVARALog:       True
========= PARAMETERS FOR INSTALLATION DETAILS =============
RMSPacketInstalled:   True
RMSPacketDirectory:   C:\RMS\RMS Packet\
RMSPacketProcess:     RMS Packet
RMSPacketSettings:    RMS Packet.ini
RMSTrimodeInstalled:  True
RMSTrimodeDirectory:  C:\RMS\RMS Trimode\
RMSTrimodeProcess:    RMS Trimode
RMSTrimodeSettings:   RMS Trimode.ini
RMSRelayInstalled:    True
RMSRelayDirectory:    C:\RMS\RMS Relay\
RMSRelayProcess:      RMS Relay
RMSRelaySettings:     RMS Relay.ini
VARAFMInstalled:      True
VARAFMDirectory:      C:\VARA FM\
VARAFMProcess:        VARAFM
VARAFMSettings:       VARAFM.ini
VARAFMLog:            VARAFM.log
VARAInstalled:        True
VARADirectory:        C:\VARA\
VARAProcess:          VARA
VARASettings:         VARA.ini
VARALog:              VARA.log
========= PARAMETERS FOR CHECKS TO PERFORM ================
CheckLogErrors:       True
CheckVARACPU:         True
CheckPort:            True
RestartOnLogError:    True
RestartOnError:       True
RestartOnLowCPU:      False
RestartOnPortError:   True
VARAFMCPURestart:     1.5
VARACPURestart:       0.5
EmailOnError:         True
========= PARAMETERS FOR LOG AND SETTINGS FILES ===========
LogRotate:            True
SettingsBackup:       True
ForceBackup:          False
SettingsRestore:      True
========= PARAMETERS FOR VOLUME ADJUSTMENTS ===============
AdjustSpeakerVolume:  True
SpeakerVolume:        25
========= PARAMETERS FOR ACTIONS TO PERFORM ===============
StatusRMSPacket:      True
StatusRMSTrimode:     True
StatusRMSPRelay:      True
StartRMSPacket:       False
StartRMSTrimode:      False
StartRMSRelay:        False
RestartRMSPacket:     False
RestartRMSTrimode:    False
RestartRMSRelay:      False
StopRMSPacket:        False
StopRMSTrimode:       False
StopRMSRelay:         False
========= GETTING SYSTEM SERIAL COM PORTS =================
System Serial Ports:  COM3 COM7 COM6 COM4 COM5
========= GETTING SYSTEM AUDIO DEVICES ====================
Audio Playback:       ICOM IC7300 Output Speakers (2- USB Audio CODEC ) Speakers (Realtek High Definition Audio) ICOM IC9700 Output Speakers (USB Audio CODEC )
Audio Recoding:       ICOM IC7300 Input Microphone (2- USB Audio CODEC ) ICOM IC9700 Input Microphone (USB Audio CODEC )
========= GETTING APPLICATION INI SETTINGS ================
RMS Packet Version:   2.1.43.0
RMS Packet INI PID:   13652 (RUNNING)
RMS Packet VARA FM:   True
VARA FM Sound In:     ICOM IC9700 Input Microphone (U (VALID)
VARA FM Sound Out:    ICOM IC9700 Output Speakers (US (VALID)
VARA FM PTT Port:     COM4 (VALID)
VARA FM CAT Port:     COM4 (VALID)
VARA FM Command Port: 8300 (IN USE)
RMS Trimode Version:  1.3.49.0
RMS Trimode INI PID:  15580 (RUNNING)
RMS Trimode Cmd Port: 8510 (IN USE)
RMS Trimode Rly Port: 8772 (IN USE)
VARA Sound In:        ICOM IC7300 Input Microphone (2 (VALID)
VARA Sound Out:       ICOM IC7300 Output Speakers (2- (VALID)
VARA PTT Port:        COM1 (NOT VALID)
VARA CAT Port:        COM1 (NOT VALID)
VARA Command Port:    8300 (IN USE)
RMS Relay Version:    3.2.3.0
RMS Relay INI PID:    26296 (RUNNING)
========= SETTINGS FILES BACKUP ===========================
Settings saved:       RMS Packet.ini
Settings saved:       VARAFM.ini
Settings saved:       RMS Trimode.ini
Settings saved:       VARA.ini
Settings saved:       RMS Relay.ini
========= SETTING THE DEFAULT DEVICE VOLUME ===============
Setting volume to:    25 on default audio device
========= CHECKING PROCESSES ==============================
Process RMS Packet:   Is Running (Windows Process: 13652)
Process VARAFM:       Is Running (Windows Process: 21644)
Process RMS Trimode:  Is Running (Windows Process: 15580)
Process VARA:         Is Running (Windows Process: 22608)
Process RMS Relay:    Is Running (Windows Process: 26296)
========= CHECKING PROCESS LOG ERRORS CONDITIONS ==========
Checking Log:         C:\VARA FM\VARAFM.log
Checking Log:         C:\VARA\VARA.log
========= CHECKING PROCESS CPU USE ========================
Checking VARA FM CPU: 7.65329238515776
Checking VARA CPU:    10.8843831219267
========= CHECKING PROCESS SOCKET PORT USE ================
RMS Packet Ports:     65239 65238 65239 65238
VARA FM Ports:        58787 58787 8301 8300
RMS Trimode Ports:    65245 65244 65245 65244 8510
VARA Ports:           58832 58832 8301 8300
RMS Relay Ports:      59360 59357 59360 59357 8772
========= SCRIPT COMPLETED ================================

#> 

## Script PARAMETERS
param (
    # Location of log file that contains all messages and errors 
    $LogFile="$PSScriptRoot\RunRMS.log",

    # Location of the lock file that is created whenever this script is running
    $LockFile="$PSScriptRoot\RunRMS.lck",

    # Log to Windows Application EventLog
    $EventLogger=$true,

    # Event Source Name to use in Windows EventLog
    $EventSource="RunRMS",

    # Set to true to write all messages and errors to a log file, as well as the screen
    $LogFileWrite=$true,

    # Set to true to clear out the log file each time the script is run
    $LogFileReset=$true,

    # Set to true to email the log file after the script is finished running each time
    $LogFileEmail=$false,

    # Set to true to also attach the VARA Logs to the email after the script is finished running
    $IncludeVARALog=$true,

    # Set to true if RMS Packet is installed
    $RMSPacketInstalled=$true,

    # Directory where RMS Packet is installed
    $RMSPacketDirectory="C:\RMS\RMS Packet\",

    # Name of RMS Packet Process.  LEAVE OFF THE .EXE AT THE END
    $RMSPacketProcess="RMS Packet",

    # Name of the RMS Packet Settings File.  NEEDS .INI AT END
    $RMSPacketSettings="RMS Packet.ini",

    # Set to true if RMS Trimode is installed
    $RMSTrimodeInstalled=$true,

    # Directory where RMS Trimode is installed
    $RMSTrimodeDirectory="C:\RMS\RMS Trimode\",

    # Name of RMS Trimode Process.  LEAVE OFF THE .EXE AT THE END
    $RMSTrimodeProcess="RMS Trimode",

    # Name of the RMS Trimode Settings File.  NEEDS .INI AT END
    $RMSTrimodeSettings="RMS Trimode.ini",

    # Set to true if RMS Relay is installed
    $RMSRelayInstalled=$true,

    # Directory where RMS Relay is installed
    $RMSRelayDirectory="C:\RMS\RMS Relay\",

    # Name of RMS Relay Process.  LEAVE OFF THE .EXE AT THE END
    $RMSRelayProcess="RMS Relay",

    # Name of the RMS Relay Settings File.  NEEDS .INI AT END
    $RMSRelaySettings="RMS Relay.ini",

    # Set to true if VARA FM is installed
    $VARAFMInstalled=$true,

    # Directory where VARA FM is installed
    $VARAFMDirectory="C:\VARA FM\",

    # Name of VARA FM Process.  LEAVE OFF THE .EXE AT THE END
    $VARAFMProcess="VARAFM",

    # Name of the VARA FM Settings File.  NEEDS .INI AT THE END
    $VARAFMSettings="VARAFM.ini",

    # Name of the VARA FM Log File.  NEEDS .LOG AT THE END
    $VARAFMLog="VARAFM.log",

    # Set to true if VARA is installed
    $VARAInstalled=$true,

    # Directory where VARA is installed
    $VARADirectory="C:\VARA\",

    # Name of VARA Process.  LEAVE OFF THE .EXE AT THE END
    $VARAProcess="VARA",

    # Name of the VARA Settings File.  NEEDS .INI AT THE END
    $VARASettings="VARA.ini",

    # Name of the VARA Log File.  NEEDS .LOG AT THE END
    $VARALog="VARA.log",

    # Check the VARA Log File for common error message strings
    $CheckLogErrors=$true,

    # Check the VARA CPU Usage
    $CheckVARACPU=$true,

    # Check the Application TCP Ports Are In Use
    $CheckPort=$true,

    # Restart the RMS process when VARA log file error messages are found
    $RestartOnLogError=$true,

    # Restart the RMS process when either the RMS Process or VARA Process not running
    $RestartOnError=$true,

    # Restart the RMS process whenever the CPU Use of VARA falls below this level
    $RestartOnLowCPU=$false,

    # Restart the RMS process whenever no Ports are detected in use
    $RestartOnPortError=$true,

    # Threshold of VARA FM CPU Use to be considered too low
    $VARAFMCPURestart=1.5,

    # Threshold of VARA CPU Use to be considered too low
    $VARACPURestart=0.5,

    # Set to true to email when either errors above (log file or process not running) occur
    $EmailOnError=$true,

    # Rename the VARA Log File with current date/time just before starting the RMS Process
    $LogRotate=$true,

    # If either of the .INI settings do not have a copy with .BAK, copy it for backup
    $SettingsBackup=$true,

    # Force the backup of the .INI settings to .BAK, even when backup already exists
    $ForceBackup=$false,

    # Restore the backup .INI.BAK files to .INI just before starting the RMS Process
    $SettingsRestore=$true,

    # Adjust our default system sound speaker device to SpeakerVolume
    $AdjustSpeakerVolume=$true,

    # Level to set our default system sound speaker device to.  
    # This value is half of the percentage level, ie 25 equals 50% volume
    $SpeakerVolume=25,

    # Checks to see if the RMS Process and VARA Process are both running
    $StatusRMSPacket=$true,
    $StatusRMSTrimode=$true,
    $StatusRMSRelay=$true,

    # Start the RMS process 
    $StartRMSPacket=$false,
    $StartRMSTrimode=$false,
    $StartRMSRelay=$false,

    # Restart (Stop/Start) the RMS process 
    $RestartRMSPacket=$false,
    $RestartRMSTrimode=$false,
    $RestartRMSRelay=$false,

    # Start the RMS process 
    $StopRMSPacket=$false,
    $StopRMSTrimode=$false,
    $StopRMSRelay=$false
    )

## Various Settings
# SMTP Server Settings
$smtpServer     = "email-smtp.us-east-1.amazonaws.com"
$smtpPort       = 587
$smtpUsername   = "USERNAME"
$smtpPassword   = "PASSWORD"
$smtpFrom       = "me@mydomain.com"
$smtpTo         = "me@mydomain.com"

# Number of seconds to wait after starting or stopping process before checking on them
$SleepAfterStart = 10
$SleepAfterStop  = 1

# Set errors to false
$OnError        = $False
$OnLogError     = $False
$OnLowCPUError  = $False
$OnPortError    = $False

# Set Restarted flags to false
$RestartedRMSPacket     = $False
$RestartedRMSTrimode    = $False
$RestartedRMSRelay      = $false

function Check-AudioDevice {
    param(
        [parameter(Mandatory = $true)] $AudioDevice
    )
    if((Get-AudioDevice -List | Select-String -SimpleMatch $AudioDevice) -eq $Null) {
        return $true 
    } else {
        return $true
    }
}

function Check-COMPort {
    param(
        [parameter(Mandatory = $true)] $COMPort
    )
    if(([System.IO.Ports.SerialPort]::getportnames() | Select-String -Pattern $COMPort) -eq $Null) {
        return $false
    } else {
        return $true
    }
}

function Check-Process {
    param(
        [parameter(Mandatory = $true)] $ProcessId
    )
    if((get-process -Id "$ProcessId" -ea SilentlyContinue) -eq $Null){ 
        return $false
    } else {
        return $true
    }
}

function Check-Port {
    param(
        [parameter(Mandatory = $true)] $PortId
    )
    if((Get-NetTCPConnection | where Localport -eq $PortId) -eq $Null){ 
        return $false
    } else {
        return $true
    }
}

function Get-AudioPlayback {
    $AudioPlayback = Get-AudioDevice -List | where Type -like "Playback" | select -Expand Name
    Write-Message -Message "Audio Playback:       $AudioPlayback" -fore Blue
}

function Get-AudioRecording {
    $AudioRecording = Get-AudioDevice -List | where Type -like "Recording" | select -Expand Name
    Write-Message -Message "Audio Recoding:       $AudioRecording" -fore Blue
}

function Get-COM-Ports {
    $COMPorts = [System.IO.Ports.SerialPort]::getportnames()
    Write-Message -Message "System Serial Ports:  $COMPorts" -fore Blue
}

function Get-IniFile 
{  
    param(  
        [parameter(Mandatory = $true)] [string] $filePath  
    )  
    
    $anonymous = "NoSection"
  
    $ini = @{}  
    switch -regex -file $filePath  
    {  
        "^\[(.+)\]$" # Section  
        {  
            $section = $matches[1]  
            $ini[$section] = @{}  
            $CommentCount = 0  
        }  

        "^(;.*)$" # Comment  
        {  
            if (!($section))  
            {  
                $section = $anonymous  
                $ini[$section] = @{}  
            }  
            $value = $matches[1]  
            $CommentCount = $CommentCount + 1  
            $name = "Comment" + $CommentCount  
            $ini[$section][$name] = $value  
        }   

        "(.+?)\s*=\s*(.*)" # Key  
        {  
            if (!($section))  
            {  
                $section = $anonymous  
                $ini[$section] = @{}  
            }  
            $name,$value = $matches[1..2]  
            $ini[$section][$name] = $value 
        }  
    }  

    return $ini  
}  

function Get-INI-Settings() {
    if ($RMSPacketInstalled -eq $true) {
        $RMSPacketINIValue  = Get-IniFile $RMSPacketDirectory$RMSPacketSettings
        if ($VARAFMInstalled -eq $true) {
            $VARAFMINIValue = Get-IniFile $VARAFMDirectory$VARAFMSettings
        }
        $RMSPacketVersion  = $RMSPacketINIValue.'MAIN'.'Program Version'
        $RMSPacketPID      = $RMSPacketINIValue.'MAIN'.'Process ID'
        $RMSPacketVARAFM   = $RMSPacketINIValue.'Vara FM TNC'.'Enable Vara FM'
        if ($VARAFMInstalled -eq $true) {
            $VARAFMSoundIn     = $VARAFMINIValue.'Soundcard'.'Input Device Name'
            $VARAFMSoundOut    = $VARAFMINIValue.'Soundcard'.'Output Device Name'
            $VARAFMPTTPort     = $VARAFMINIValue.'PTT'.'PTTPort'
            $VARAFMCATPort     = $VARAFMINIValue.'PTT'.'CATPort'
            $VARAFMCmdPort     = $VARAFMINIValue.'Setup'.'TCP Command Port'
        }
        Write-Message -Message "RMS Packet Version:   $RMSPacketVersion" -fore Green
        if ((Check-Process($RMSPacketPID)) -eq $true) {
            Write-Message -Message "RMS Packet INI PID:   $RMSPacketPID (RUNNING)" -fore Green
        } else {
            Write-Message -Message "RMS Packet INI PID:   $RMSPacketPID (NOT RUNNING)" -fore Red
        }
        Write-Message -Message "RMS Packet VARA FM:   $RMSPacketVARAFM" -fore Green
        if ($VARAFMInstalled -eq $true) {
            if ((Check-AudioDevice($VARAFMSoundIn)) -eq $true) {
                Write-Message -Message "VARA FM Sound In:     $VARAFMSoundIn (VALID)" -fore Green
            } else {
                Write-Message -Message "VARA FM Sound In:     $VARAFMSoundIn (NOT VALID)" -fore Red
            }
            if ((Check-AudioDevice($VARAFMSoundOut)) -eq $true) {   
                Write-Message -Message "VARA FM Sound Out:    $VARAFMSoundOut (VALID)" -fore Green
            } else {
                Write-Message -Message "VARA FM Sound Out:    $VARAFMSoundOut (NOT VALID)" -fore Red
            }
            if ((Check-COMPort($VARAFMPTTPort)) -eq $true) {
                Write-Message -Message "VARA FM PTT Port:     $VARAFMPTTPort (VALID)" -fore Green
            } else {
                Write-Message -Message "VARA FM PTT Port:     $VARAFMPTTPort (NOT VALID)" -fore Red
            }
            if ((Check-COMPort($VARAFMCATPort)) -eq $true) {
                Write-Message -Message "VARA FM CAT Port:     $VARAFMCATPort (VALID)" -fore Green
            } else {
                Write-Message -Message "VARA FM CAT Port:     $VARAFMCATPort (NOT VALID)" -fore Red
            }
            if ((Check-Port($VARAFMCmdPort)) -eq $true) {
                Write-Message -Message "VARA FM Command Port: $VARAFMCmdPort (IN USE)" -fore Green
            } else {
                Write-Message -Message "VARA FM Command Port: $VARAFMCmdPort (NOT IN USE)" -fore Red
            }
        }
    }
    if ($RMSTrimodeInstalled -eq $true) {
        $RMSTrimodeINIValue = Get-IniFile $RMSTrimodeDirectory$RMSTrimodeSettings
        if ($VARAInstalled -eq $true) {
            $VARAINIValue   = Get-IniFile $VARADirectory$VARASettings
        }
        $RMSTrimodeVersion = $RMSTrimodeINIValue.'MAIN'.'Product Version'
        $RMSTrimodePID     = $RMSTrimodeINIValue.'MAIN'.'Process ID'
        $RMSTrimodeCmdPort = $RMSTrimodeINIValue.'Site Properties'.'CmdPort'
        $RMSTrimodeRlyPort = $RMSTrimodeINIValue.'Site Properties'.'RMSRelayPort'
        if ($VARAInstalled -eq $true) {
            $VARASoundIn     = $VARAINIValue.'Soundcard'.'Input Device Name'
            $VARASoundOut    = $VARAINIValue.'Soundcard'.'Output Device Name'
            $VARAPTTPort     = $VARAINIValue.'PTT'.'PTTPort'
            $VARACATPort     = $VARAINIValue.'PTT'.'CATPort'
            $VARACmdPort     = $VARAINIValue.'Setup'.'TCP Command Port'
        }
        Write-Message -Message "RMS Trimode Version:  $RMSTrimodeVersion" -fore Green
        if ((Check-Process($RMSPacketPID)) -eq $true) {
            Write-Message -Message "RMS Trimode INI PID:  $RMSTrimodePID (RUNNING)" -fore Green
        } else {
            Write-Message -Message "RMS Trimode INI PID:  $RMSTrimodePID (NOT RUNNING)" -fore Red
        }
        if ((Check-Port($RMSTrimodeCmdPort)) -eq $true) {
            Write-Message -Message "RMS Trimode Cmd Port: $RMSTrimodeCmdPort (IN USE)" -fore Green
        } else {
            Write-Message -Message "RMS Trimode Cmd Port: $RMSTrimodeCmdPort (NOT IN USE)" -fore Red
        }
        if ((Check-Port($RMSTrimodeRlyPort)) -eq $true) {
            Write-Message -Message "RMS Trimode Rly Port: $RMSTrimodeRlyPort (IN USE)" -fore Green
        } else {
            Write-Message -Message "RMS Trimode Rly Port: $RMSTrimodeRlyPort (NOT IN USE)" -fore Red
        }
        if ($VARAInstalled -eq $true) {
            if ((Check-AudioDevice($VARASoundIn)) -eq $true) {
            Write-Message -Message "VARA Sound In:        $VARASoundIn (VALID)" -fore Green
            } else {
            Write-Message -Message "VARA Sound In:        $VARASoundIn (NOT VALID)" -fore red

            }
            if ((Check-AudioDevice($VARASoundOut)) -eq $true) {
            Write-Message -Message "VARA Sound Out:       $VARASoundOut (VALID)" -fore Green
            } else {
            Write-Message -Message "VARA Sound Out:       $VARASoundOut (NOT VALID)" -fore red
            }
            if ((Check-COMPort($VARAPTTPort)) -eq $true) {
                Write-Message -Message "VARA PTT Port:        $VARAPTTPort (VALID)" -fore Green
            } else {
                Write-Message -Message "VARA PTT Port:        $VARAPTTPort (NOT VALID)" -fore Red
            }
            if ((Check-COMPort($VARACATPort)) -eq $true) {
                Write-Message -Message "VARA CAT Port:        $VARACATPort (VALID)" -fore Green
            } else {
                Write-Message -Message "VARA CAT Port:        $VARACATPort (NOT VALID)" -fore Red
            }
            if ((Check-Port($VARACmdPort)) -eq $true) {
                Write-Message -Message "VARA Command Port:    $VARACmdPort (IN USE)" -fore Green
            } else {
                Write-Message -Message "VARA Command Port:    $VARACmdPort (NOT IN USE)" -fore Red
            }
        }
    }
    if ($RMSRelayInstalled -eq $true) {
        $RMSRelayINIValue   = Get-IniFile $RMSRelayDirectory$RMSRelaySettings
        $RMSRelayVersion = $RMSRelayINIValue.'MAIN'.'Program Version'
        $RMSRelayPID     = $RMSRelayINIValue.'MAIN'.'Process ID'
        Write-Message -Message "RMS Relay Version:    $RMSRelayVersion" -fore Green
        if ((Check-Process($RMSPacketPID)) -eq $true) {
            Write-Message -Message "RMS Relay INI PID:    $RMSRelayPID (RUNNING)" -fore Green
        } else {
            Write-Message -Message "RMS Relay INI PID:    $RMSRelayPID (NOT RUNNING)" -fore Red
        }
    }
}

function Load-Module ($m) {
    # If module is imported do nothing
    if (Get-Module | Where-Object {$_.Name -eq $m}) {
    }
    else {
        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) {
            Import-Module $m
        }
        else {
            # If module is not imported, not available on disk
            # but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) {
                Install-Module -Name $m -Force -Scope CurrentUser
                Import-Module $m
            }
            else {
                # If the module is not imported, not available and not in the online gallery
                write-host "WARNING: Module $m not imported, not available and not in an online gallery."
            }
        }
    }
}

# Function to import modules needed for script
function Import-Modules() {
	Load-Module ("AudioDeviceCmdlets")
}

# Function to create a new Windows Event Log entry
function Write-EventLog-Error() {
    param(
        $Message = "Error in RunRMS Script",
        $EventID = 1
    )
    if ($EventLogger -eq $True) {
        ## TODO - Write-EventLog removed in PS 7.2, use New-WinEvent instead
        ## Write-EventLog -LogName Application -Source $EventSource -EntryType Error -Message $Message -EventID $EventID
    }
}

# Function to send an email message, example here uses Amazon SES
function Send-Email-Log() {
    param(
        $subject = "Run RMS Script - Output Log File",
        $body = "Please review the Run RMS Script output log file for station attached.",
        $attachments = "$LogFile"
        )
    $smtp = new-object Net.Mail.SmtpClient($smtpServer, $smtpPort)
    $smtp.EnableSsl = $true 
    $smtp.Credentials = new-object Net.NetworkCredential($smtpUsername, $smtpPassword)
    $msg = new-object Net.Mail.MailMessage
    $msg.From = $smtpFrom
    $msg.To.Add($smtpTo)
    $msg.Subject = $subject
    $msg.Body = $body
    $msg.attachments.add($attachments)
    if ($VARAFMInstalled -eq $true -and $IncludeVARALog -eq $True) {
        $attachment = $VARAFMLogFileEmail
        if (Test-Path -Path $attachment -PathType Leaf) {
            $msg.attachments.add($attachment)
        }
    }
    if ($VARAInstalled -eq $true -and $IncludeVARALog -eq $True) {
        $attachment = $VARALogFileEmail
        if (Test-Path -Path $attachment -PathType Leaf) {
            $msg.attachments.add($attachment)
        }
    }
    $smtp.Send($msg)
}

# Function to send an email message, example here uses Amazon SES
function Send-Email-Error() {
    param(
        $subject = "Run RMS Script - ERROR",
        $body = "An error has occured."
        )
    $smtp = new-object Net.Mail.SmtpClient($smtpServer, $smtpPort)
    $smtp.EnableSsl = $true 
    $smtp.Credentials = new-object Net.NetworkCredential($smtpUsername, $smtpPassword)
    $msg = new-object Net.Mail.MailMessage
    $msg.From = $smtpFrom
    $msg.To.Add($smtpTo)
    $msg.Subject = $subject
    $msg.Body = $body
    $smtp.Send($msg)
}

# Function to write a message to the screen, and also a log file when turned on
function Write-Message() {
    param($Message, $Fore)
    if ($LogFileWrite -eq $True) {
        Write-Output $Message >> $LogFile
    }
    Write-Host $Message -ForegroundColor $Fore
}

# Function to rotate the VARA FM log file
function Log-RotateVARAFM() {
    if ($VARAFMInstalled -eq $true) {
        [string]$filePath = $VARAFMDirectory + $VARAFMLog;
        [string]$directory = [System.IO.Path]::GetDirectoryName($filePath);
        [string]$strippedFileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath);
        [string]$extension = [System.IO.Path]::GetExtension($filePath);
        [string]$newFileName = $strippedFileName + [DateTime]::Now.ToString("yyyyMMdd-HHmmss") + $extension;
        [string]$newFilePath = [System.IO.Path]::Combine($directory, $newFileName);
        Move-Item -LiteralPath $filePath -Destination $newFilePath;
        $VARAFMLogFileEmail = $newFilePath
        write-message -Message "Log File rotated:     $VARAFMDirectory$VARAFMLog to $newFilePath" -fore yellow
    }
}

# Function to rotate the VARA log file
function Log-RotateVARA() {
    if ($VARAInstalled -eq $true) {
        [string]$filePath = $VARADirectory + $VARALog;
        [string]$directory = [System.IO.Path]::GetDirectoryName($filePath);
        [string]$strippedFileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath);
        [string]$extension = [System.IO.Path]::GetExtension($filePath);
        [string]$newFileName = $strippedFileName + [DateTime]::Now.ToString("yyyyMMdd-HHmmss") + $extension;
        [string]$newFilePath = [System.IO.Path]::Combine($directory, $newFileName);
        Move-Item -LiteralPath $filePath -Destination $newFilePath;
        $VARALogFileEmail = $newFilePath
        write-message -Message "Log File rotated:     $VARADirectory$VARALog to $newFilePath" -fore yellow
    }
}

# Function to backup the RMS Packet and VARA FM .INI settings files
function Settings-Backup {
    if ($RMSPacketInstalled -eq $true) {
        if ($ForceBackup -eq $True -or -Not (Test-Path -Path $RMSPacketDirectory$RMSPacketsettings.bak -PathType Leaf)) {
            write-message "Setting file backup:  $RMSPacketSettings" -fore Yellow
            Copy-Item -Path $RMSPacketDirectory$RMSPacketSettings -Destination $RMSPacketDirectory$RMSPacketSettings.bak
        } else {
            write-message "Settings saved:       $RMSPacketSettings" -fore Green
        }
        if ($VARAFMInstalled -eq $true) {
            if ($ForceBackup -eq $True -or -Not (Test-Path -Path $VARAFMDirectory$VARAFMSettings.bak -PathType Leaf)) {
                write-message "Setting file backup:  $VARAFMSettings" -fore Yellow
                Copy-Item -Path $VARAFMDirectory$VARAFMSettings -Destination $VARAFMDirectory$VARAFMSettings.bak
            } else {
                write-message "Settings saved:       $VARAFMSettings" -fore Green
            }
        }
    }
    if ($RMSTrimodeInstalled -eq $true) {
        if ($ForceBackup -eq $True -or -Not (Test-Path -Path $RMSTrimodeDirectory$RMSTrimodeSettings.bak -PathType Leaf)) {
            write-message "Setting file backup:  $RMSTrimodeSettings" -fore Yellow
            Copy-Item -Path $RMSTrimodeDirectory$RMSTrimodeSettings -Destination $RMSTrimodeDirectory$RMSTrimodeSettings.bak
        } else {
            write-message "Settings saved:       $RMSTrimodeSettings" -fore Green
        }
        if ($VARAInstalled -eq $true) {
            if ($ForceBackup -eq $True -or -Not (Test-Path -Path $VARADirectory$VARASettings.bak -PathType Leaf)) {
                write-message "Setting file backup:  $VARASettings" -fore Yellow
                Copy-Item -Path $VARADirectory$VARASettings -Destination $VARADirectory$VARASettings.bak
            } else {
                write-message "Settings saved:       $VARASettings" -fore Green
            }
        }
    }
    if ($RMSRelayInstalled -eq $true) {
        if ($ForceBackup -eq $True -or -Not (Test-Path -Path $RMSRelayDirectory$RMSRelaySettings.bak -PathType Leaf)) {
            write-message "Setting file backup:  $RMSRelaySettings" -fore Yellow
            Copy-Item -Path $RMSRelayDirectory$RMSRelaySettings -Destination $RMSRelayDirectory$RMSRelaySettings.bak
        } else {
            write-message "Settings saved:       $RMSRelaySettings" -fore Green
        }
    }
}

# Function to "press" volume up/down keyboard keys to set level on default system audio output speaker
Function Set-Speaker($Volume) {
    $wshShell = new-object -com wscript.shell;1..50 | % {$wshShell.SendKeys([char]174)};1..$Volume | % {$wshShell.SendKeys([char]175)}
}

# Function to check the status of the RMS Packet and VARA FM processes
function Check-StatusRMSPacket() {
    if ($RMSPacketInstalled -eq $true) {
        if((get-process "$RMSPacketProcess" -ea SilentlyContinue) -eq $Null){ 
            write-message "Process ${RMSPacketProcess}:   Not Running" -fore Red
            return $OnError=$true
        } else { 
            $RMSPacketProcessPID = (get-process "$RMSPacketProcess" | Select -Expand ID)
            write-message "Process ${RMSPacketProcess}:   Is Running (Windows Process: $RMSPacketProcessPID)" -fore Green
        }
        if ($VARAFMInstalled -eq $true) {
            if((get-process "$VARAFMProcess" -ea SilentlyContinue) -eq $Null){ 
                write-message "Process ${VARAFMProcess}:       Not Running" -fore Red
                return $OnError=$true
            } else { 
                $VARAFMProcessPID = (get-process "$VARAFMProcess" | Select -Expand ID)
                write-message "Process ${VARAFMProcess}:       Is Running (Windows Process: $VARAFMProcessPID)" -fore Green
            }
        }
    }
}

# Function to check the status of the RMS Trimode and VARA processes
function Check-StatusRMSTrimode() {
    if ($RMSTrimodeInstalled -eq $true) {
        if((get-process "$RMSTrimodeProcess" -ea SilentlyContinue) -eq $Null){ 
            write-message "Process ${RMSTrimodeProcess}:  Not Running" -fore Red
            return $OnError=$true
        } else { 
            $RMSTrimodeProcessPID = (get-process "$RMSTrimodeProcess" | Select -Expand ID)
            write-message "Process ${RMSTrimodeProcess}:  Is Running (Windows Process: $RMSTrimodeProcessPID)" -fore Green
        }
        if ($VARAInstalled -eq $true) {
            if((get-process "$VARAProcess" -ea SilentlyContinue) -eq $Null){ 
                write-message "Process ${VARAProcess}:         Not Running" -fore Red
                return $OnError=$true
            } else { 
                $VARAProcessPID = (get-process "$VARAProcess" | Select -Expand ID)
                write-message "Process ${VARAProcess}:         Is Running (Windows Process: $VARAProcessPID)" -fore Green
            }
        }
    }
}

# Function to check the status of the RMS Trimode and VARA processes
function Check-StatusRMSRelay() {
    if ($RMSRelayInstalled -eq $true) {
        if((get-process "$RMSRelayProcess" -ea SilentlyContinue) -eq $Null){ 
            write-message "Process ${RMSRelayProcess}:    Not Running" -fore Red
            return $OnError=$true
        } else { 
            $RMSRelayProcessPID = (get-process "$RMSRelayProcess" | Select -Expand ID)
            write-message "Process ${RMSRelayProcess}:    Is Running (Windows Process: $RMSRelayProcessPID)" -fore Green
        }
    }
}

# Function to start the RMS Packet Process
function Start-ProcessRMSPacket() {
    if ($RMSPacketInstalled -eq $true) {
        if((get-process "$RMSPacketProcess" -ea SilentlyContinue) -eq $Null) { 
            if ($SettingsRestore -eq $true) {
                write-message "Setting file restore: $RMSPacketSettings" -fore Yellow
                Copy-Item -Force -Destination $RMSPacketDirectory$RMSPacketSettings -Path $RMSPacketDirectory$RMSPacketSettings.bak
                if ($VARAFMInstalled -eq $true) {
                    write-message "Setting file restore: $VARAFMSettings" -fore Yellow
                    Copy-Item -Force -Destination $VARAFMDirectory$VARAFMSettings -Path $VARAFMDirectory$VARAFMSettings.bak
                }
            }
            if ($LogRotate -eq $true) {
                Log-RotateVARAFM
            }
            write-message "Start Process:        $RMSPacketProcess" -fore Green
            Start-Process -FilePath $RMSPacketDirectory$RMSPacketProcess.exe        
            sleep $SleepAfterStart
            $OnError = Check-StatusRMSPacket
        } else {
            write-message "Start Process:        $RMSPacketProcess is already running" -fore Yellow
        }
    }
}

# Function to start the RMS Trimode Process
function Start-ProcessRMSTrimode() {
    if ($RMSTrimodeInstalled -eq $true) {
        if((get-process "$RMSTrimodeProcess" -ea SilentlyContinue) -eq $Null) { 
            if ($SettingsRestore -eq $true) {
                write-message "Setting file restore: $RMSTrimodeSettings" -fore Yellow
                Copy-Item -Force -Destination $RMSTrimodeDirectory$RMSTrimodeSettings -Path $RMSTrimodeDirectory$RMSTrimodeSettings.bak
                if ($VARAInstalled -eq $true) {
                    write-message "Setting file restore: $VARASettings" -fore Yellow
                    Copy-Item -Force -Destination $VARADirectory$VARASettings -Path $VARADirectory$VARASettings.bak
                }
            }
            if ($LogRotate -eq $true) {
                Log-RotateVARA
            }
            write-message "Start Process:        $RMSTrimodeProcess" -fore Green
            Start-Process -FilePath $RMSTrimodeDirectory$RMSTrimodeProcess.exe        
            sleep $SleepAfterStart
            $OnError = Check-StatusRMSTrimode
        } else {
            write-message "Start Process:        $RMSTrimodeProcess is already running" -fore Yellow
        }
    }
}

# Function to start the RMS Trimode Process
function Start-ProcessRMSRelay() {
    if ($RMSRelayInstalled -eq $true) {
        if((get-process "$RMSRelayProcess" -ea SilentlyContinue) -eq $Null) { 
            if ($SettingsRestore -eq $true) {
                write-message "Setting file restore: $RMSRelaySettings" -fore Yellow
                Copy-Item -Force -Destination $RMSRelayDirectory$RMSRelaySettings -Path $RMSRelayDirectory$RMSRelaySettings.bak
            }
            write-message "Start Process:        $RMSRelayProcess" -fore Green
            Start-Process -FilePath $RMSRelayDirectory$RMSRelayProcess.exe        
            sleep 10
            $OnError = Check-StatusRMSRelay
        } else {
            write-message "Start Process:        $RMSRelayProcess is already running" -fore Yellow
        }
    }
}

# Function to stop the RMS Packet Process
function Stop-ProcessRMSPacket() {
    if ($RMSPacketInstalled -eq $true) {
        if((get-process "$RMSPacketProcess" -ea SilentlyContinue) -eq $Null) { 
            write-message "Stop Process:         $RMSPacketProcess is not running" -fore Yellow
        } else {
            write-message "Stop Process:         $RMSPacketProcess" -fore Green
            Stop-Process -ProcessName $RMSPacketProcess -Force
            sleep $SleepAfterStop
            if ($VARAFMInstalled -eq $true) {
                if((get-process "$VARAFMProcess" -ea SilentlyContinue) -eq $Null) { 
                    write-message "Stop Process:         $VARAFMProcess is not running" -fore Yellow 
                } else {
                    write-message "Stop Process:         $VARAFMProcess" -fore Green
                    Stop-Process -ProcessName $VARAFMProcess -Force
                }
            }
        }
        sleep $SleepAfterStop
        $OnError = Check-StatusRMSPacket
    }
}

# Function to stop the RMS Trimode Process
function Stop-ProcessRMSTrimode() {
    if ($RMSTrimodeInstalled -eq $true) {
        if((get-process "$RMSTrimodeProcess" -ea SilentlyContinue) -eq $Null) { 
            write-message "Stop Process:         $RMSTrimodeProcess is not running" -fore Yellow
        } else {
            write-message "Stop Process:         $RMSTrimodeProcess" -fore Green
            Stop-Process -ProcessName $RMSTrimodeProcess -Force
            sleep $SleepAfterStop
            if ($VARAInstalled -eq $true) {
                if((get-process "$VARAProcess" -ea SilentlyContinue) -eq $Null) { 
                    write-message "Stop Process:         $VARAProcess is not running" -fore Yellow 
                } else {
                    write-message "Stop Process:         $VARAProcess" -fore Green
                    Stop-Process -ProcessName $VARAProcess -Force
                }
            }
        }
        sleep $SleepAfterStop
        $OnError = Check-StatusRMSTrimode
    }
}

# Function to stop the RMS Relay Process
function Stop-ProcessRMSRelay() {
    if ($RMSRelayInstalled -eq $true) {
        if((get-process "$RMSRelayProcess" -ea SilentlyContinue) -eq $Null) { 
            write-message "Stop Process:         $RMSRelayProcess is not running" -fore Yellow
        } else {
            write-message "Stop Process:         $RMSRelayProcess" -fore Green
            Stop-Process -ProcessName $RMSRelayProcess -Force
            sleep $SleepAfterStop
        }
        sleep $SleepAfterStop
        $OnError = Check-StatusRMSRelay
    }
}

# Function to restart (stop/start) the RMS Packet Process
function Restart-ProcessRMSPacket() {
    if ($RMSPacketInstalled -eq $true -and $RestartedRMSPacket -eq $False) {
        $RestartedRMSPacket = $True
        Stop-ProcessRMSPacket
        Start-ProcessRMSPacket
    }
}

# Function to restart (stop/start) the RMS Trimode Process
function Restart-ProcessRMSTrimode() {
    if ($RMSTrimodeInstalled -eq $true -and $RestartedRMSTrimode -eq $False) {
        $RestartedRMSTrimode = $True
        Stop-ProcessRMSTrimode
        Start-ProcessRMSTrimode
    }
}

# Function to restart (stop/start) the RMS Relay Process
function Restart-ProcessRMSRelay() {
    if ($RMSRelayInstalled -eq $true -and $RestartedRMSRelay -eq $False) {
        $RestartedRMSRelay = $True
        Stop-ProcessRMSRelay
        Start-ProcessRMSRelay
    }
}

# Function to check the VARA Log file for common error strings
function Check-LogErrors() {
    if ($VARAFMInstalled -eq $true) {
        write-message "Checking Log:         $VARAFMDirectory$VARAFMLog" -Fore Yellow
        if (Test-Path -Path $VARAFMDirectory$VARAFMLog -PathType Leaf) {
            if (Select-String -Path $VARAFMDirectory$VARAFMLog -Quiet -Pattern "Soundcard Device missing in action !") {
                write-message "Log Errors Found:     Soundcard Device missing in action !" -fore Red
                return $OnLogError=$true
            }
            if (Select-String -Path $VARAFMDirectory$VARAFMLog -Quiet -Pattern "WARNING: VARA is not connected to any App via TCP Port") {
                write-message "Log Errors Found:     WARNING: VARA is not connected to any App via TCP Port" -fore Red
                return $OnLogError=$true
            } 
        } else {
            New-Item -Path $VARAFMDirectory -Name $VARAFMLog -ItemType "file"
            write-message "Log File created:     $VARAFMDirectory$VARAFMLog" -fore yellow
        }
    }
    if ($VARAInstalled -eq $true) {
        write-message "Checking Log:         $VARADirectory$VARALog" -Fore Yellow
        if (Test-Path -Path $VARADirectory$VARALog -PathType Leaf) {
            if (Select-String -Path $VARADirectory$VARALog -Quiet -Pattern "Soundcard Device missing in action !") {
                write-message "Log Errors Found:     Soundcard Device missing in action !" -fore Red
                return $OnLogError=$true
            }
            if (Select-String -Path $VARADirectory$VARALog -Quiet -Pattern "WARNING: VARA is not connected to any App via TCP Port") {
                write-message "Log Errors Found:     WARNING: VARA is not connected to any App via TCP Port" -fore Red
                return $OnLogError=$true
            } 
        } else {
            New-Item -Path $VARADirectory -Name $VARALog -ItemType "file"
            write-message "Log File created:     $VARADirectory$VARALog" -fore yellow
        }
    }
}

# Function to check the VARA FM Process CPU time
function Check-VARAFMCPU() {
    if ($CheckVARACPU -eq $True) {
        $VARAFMCPU = (Get-Counter '\Process(*)\% Processor Time' -ErrorAction SilentlyContinue).CounterSamples  | Where-Object {$_.InstanceName -eq "varafm"} | Foreach-Object CookedValue
        write-message "Checking VARA FM CPU: $VARAFMCPU" -Fore Yellow
        if ($RestartOnLowCPU -eq $True -and $VARAFMCPU -lt $VARAFMCPURestart) {
            write-message "VARA FM CPU $VARAFMCPU Below Threshold Amount of $VARAFMCPURestart" -fore Red
            return $OnLowCPUError=$true
        }
    }
}

# Function to check the VARA Process CPU time
function Check-VARACPU() {
    if ($CheckVARACPU -eq $True) {
        $VARACPU = (Get-Counter '\Process(*)\% Processor Time' -ErrorAction SilentlyContinue).CounterSamples | Where-Object {$_.InstanceName -eq "vara"} | Foreach-Object CookedValue
        write-message "Checking VARA CPU:    $VARACPU" -Fore Yellow
        if ($RestartOnLowCPU -eq $True -and $VARACPU -lt $VARACPURestart) {
            write-message "VARA CPU $VARACPU Below Threshold Amount of $VARACPURestart" -fore Red
            return $OnLowCPUError=$true
        }
    }
}

# Function to check the RMS Packet Port
function Check-RMSPacketPort() {
    if ($RMSPacketInstalled -eq $true -and $CheckPort -eq $True) {
        $RMSPacketProcessPID = (get-process "$RMSPacketProcess" -ea SilentlyContinue | Select -Expand ID)
        $RMSPacketPorts = Get-NetTCPConnection | where OwningProcess -eq $RMSPacketProcessPID | Select -Expand Localport
        if ($RMSPacketPorts -eq $null) {
            write-message "RMS Packet Ports:     NONE ACTIVE" -Fore Red
            return $OnPortError=$true
        } else {
            write-message "RMS Packet Ports:     $RMSPacketPorts" -Fore Green
        }
    }
}

# Function to check the RMS Packet Port
function Check-VARAFMPort() {
    if ($VARAFMInstalled -eq $true -and $CheckPort -eq $True) {
        $VARAFMProcessPID = (get-process "$VARAFMProcess" -ea SilentlyContinue | Select -Expand ID)
        $VARAFMPorts = Get-NetTCPConnection | where OwningProcess -eq $VARAFMProcessPID | Select -Expand Localport
        if ($VARAFMPorts -eq $null) {
            write-message "VARA FM Ports:        NONE ACTIVE" -Fore Red
            return $OnPortError=$true
        } else {
            write-message "VARA FM Ports:        $VARAFMPorts" -Fore Green
        }
    }
}

# Function to check the RMS Trimode Port
function Check-RMSTrimodePort() {
    if ($RMSTrimodeInstalled -eq $true -and $CheckPort -eq $True) {
        $RMSTrimodeProcessPID = (get-process "$RMSTrimodeProcess" -ea SilentlyContinue | Select -Expand ID)
        $RMSTrimodePorts = Get-NetTCPConnection | where OwningProcess -eq $RMSTrimodeProcessPID | Select -Expand Localport
        if ($RMSTrimodePorts -eq $null) {
            write-message "RMS Trimode Ports:    NONE ACTIVE" -Fore Red
            return $OnPortError=$true
        } else {
            write-message "RMS Trimode Ports:    $RMSTrimodePorts" -Fore Green
        }
    }
}

# Function to check the VARA Port
function Check-VARAPort() {
    if ($VARAInstalled -eq $true -and $CheckPort -eq $True) {
        $VARAProcessPID = (get-process "$VARAProcess" -ea SilentlyContinue | Select -Expand ID)
        $VARAPorts = Get-NetTCPConnection | where OwningProcess -eq $VARAProcessPID | Select -Expand Localport
        if ($VARAPorts -eq $null) {
            write-message "VARA Ports:           NONE ACTIVE" -Fore Red
            return $OnPortError=$true
        } else { 
            write-message "VARA Ports:           $VARAPorts" -Fore Green
        }
    }
}

# Function to check the RMS Relay Port
function Check-RMSRelayPort() {
    if ($RMSRelayInstalled -eq $true -and $CheckPort -eq $True) {
        $RMSRelayProcessPID = (get-process "$RMSRelayProcess" -ea SilentlyContinue | Select -Expand ID)
        $RMSRelayPorts = Get-NetTCPConnection | where OwningProcess -eq $RMSRelayProcessPID | Select -Expand Localport
        if ($RMSRelayPorts -eq $null) {
            write-message "RMS Relay Ports:      NONE ACTIVE" -Fore Red
            return $OnPortError=$true
        } else {
            write-message "RMS Relay Ports:      $RMSRelayPorts" -Fore Green
        }
    }
}

# Function to check required command parameters to make sure they are correctly set
function Check-Parameters() {

    if ($RMSPacketInstalled -eq $true) {

        if ($RMSPacketDirectory -eq $null) {
            $RMSPacketDirectory = read-host -Prompt "Please enter RMS Packet Directory"
        }

        if (Test-Path -Path $RMSPacketDirectory) {
        } else {
            write-message "RMSPacketDirectory $RMSPacketDirectory is not valid" -fore Red
            exit 1101
        }

        if ($RMSPacketProcess -eq $null) {
            $RMSPacketProcess = read-host -Prompt "Please enter RMS Packet Process"
        }

        if (Test-Path -Path $RMSPacketDirectory$RMSPacketProcess.exe -PathType Leaf) {
        } else {
            write-message "RMSPacketProcess $RMSPacketProcess is not valid" -fore Red
            exit 1102
        }

        if ($RMSPacketSettings -eq $null) {
            $RMSPacketSettings = read-host -Prompt "Please enter RMS Packet Settings"
        }

        if (Test-Path -Path $RMSPacketDirectory$RMSPacketSettings -PathType Leaf) {
        } else {
            write-message "RMSPacketSettings $RMSPacketSettings is not valid" -fore Red
            exit 1103
        }

        if ($VARAFMInstalled -eq $true) {

            if ($VARAFMDirectory -eq $null) {
                $VARAFMDirectory = read-host -Prompt "Please enter VARA FM Directory"
            }

            if (Test-Path -Path $VARAFMDirectory) {
            } else {
                write-message "VARAFMDirectory $VARAFMDirectory is not valid" -fore Red
                exit 1104
            }

            if ($VARAFMProcess -eq $null) {
                $VARAFMProcess = read-host -Prompt "Please enter VARA FM Process"
            }

            if (Test-Path -Path $VARAFMDirectory$VARAFMProcess.exe -PathType Leaf) {
            } else {
                write-message "VARAFMProcess $VARAFMProcess is not valid" -fore Red
                exit 1105
            }

            if ($VARAFMSettings -eq $null) {
                $VARAFMSettings = read-host -Prompt "Please enter VARA FM Settings"
            }

            if (Test-Path -Path $VARAFMDirectory$VARAFMSettings -PathType Leaf) {
            } else {
                write-message "VARAFMSettings $VARAFMSettings is not valid" -fore Red
                exit 1106
            }

            if ($VARAFMLog -eq $null) {
                $VARAFMLog = read-host -Prompt "Please enter VARA FM Log"
            }

            if (Test-Path -Path $VARAFMDirectory$VARAFMLog -PathType Leaf) {
            } else {
                New-Item -Path $VARAFMDirectory -Name $VARAFMLog -ItemType File -Force
            }

        }

    }

    if ($RMSTrimodeInstalled -eq $true) {

        if ($RMSTrimodeDirectory -eq $null) {
            $RMSTrimodeDirectory = read-host -Prompt "Please enter RMS Trimode Directory"
        }

        if (Test-Path -Path $RMSTrimodeDirectory) {
        } else {
            write-message "RMSTrimodeDirectory $RMSTrimodeDirectory is not valid" -fore Red
            exit 1201
        }

        if ($RMSTrimodeProcess -eq $null) {
            $RMSTrimodeProcess = read-host -Prompt "Please enter RMS Trimode Process"
        }

        if (Test-Path -Path $RMSTrimodeDirectory$RMSTrimodeProcess.exe -PathType Leaf) {
        } else {
            write-message "RMSTrimodeProcess $RMSTrimodeProcess is not valid" -fore Red
            exit 1202
        }

        if ($RMSTrimodeSettings -eq $null) {
            $RMSTrimodeSettings = read-host -Prompt "Please enter RMS Trimode Settings"
        }

        if (Test-Path -Path $RMSTrimodeDirectory$RMSTrimodeSettings -PathType Leaf) {
        } else {
            write-message "RMSTrimodeSettings $RMSTrimodeSettings is not valid" -fore Red
            exit 1203
        }

        if ($VARAInstalled -eq $true) {

            if ($VARADirectory -eq $null) {
                $VARADirectory = read-host -Prompt "Please enter VARA Directory"
            }

            if (Test-Path -Path $VARADirectory) {
            } else {
                write-message "VARADirectory $VARADirectory is not valid" -fore Red
                exit 1204
            }

            if ($VARAProcess -eq $null) {
                $VARAProcess = read-host -Prompt "Please enter VARA Process"
            }

            if (Test-Path -Path $VARADirectory$VARAProcess.exe -PathType Leaf) {
            } else {
                write-message "VARAProcess $VARAProcess is not valid" -fore Red
                exit 1205
            }

            if ($VARASettings -eq $null) {
                $VARASettings = read-host -Prompt "Please enter VARA Settings"
            }

            if (Test-Path -Path $VARADirectory$VARASettings -PathType Leaf) {
            } else {
                write-message "VARASettings $VARASettings is not valid" -fore Red
                exit 1206
            }

            if ($VARALog -eq $null) {
                $VARALog = read-host -Prompt "Please enter VARA Log"
            }

            if (Test-Path -Path $VARADirectory$VARALog -PathType Leaf) {
            } else {
                New-Item -Path $VARADirectory -Name $VARALog -ItemType File -Force
            }

        }
    }

    if ($RMSRelayInstalled -eq $true) {

        if ($RMSRelayDirectory -eq $null) {
            $RMSRelayDirectory = read-host -Prompt "Please enter RMS Relay Directory"
        }

        if (Test-Path -Path $RMSRelayDirectory) {
        } else {
            write-message "RMSRelayDirectory $RMSRelayDirectory is not valid" -fore Red
            exit 1301
        }

        if ($RMSRelayProcess -eq $null) {
            $RMSRelayProcess = read-host -Prompt "Please enter RMS Relay Process"
        }

        if (Test-Path -Path $RMSRelayDirectory$RMSRelayProcess.exe -PathType Leaf) {
        } else {
            write-message "RMSRelayProcess $RMSRelayProcess is not valid" -fore Red
            exit 1302
        }

        if ($RMSRelaySettings -eq $null) {
            $RMSRelaySettings = read-host -Prompt "Please enter RMS Relay Settings"
        }

        if (Test-Path -Path $RMSRelayDirectory$RMSRelaySettings -PathType Leaf) {
        } else {
            write-message "RMSRelaySettings $RMSRelaySettings is not valid" -fore Red
            exit 1303
        }
    }

    write-message "========= PARAMETERS FOR LOGGING ==========================" -fore Magenta
    write-message "LogFile:              $LogFile"             -fore Blue
    write-message "EventLogger:          $EventLogger"         -fore Blue
    write-message "EventSource:          $EventSource"         -fore Blue
    write-message "LogFileWrite:         $LogFileWrite"        -fore Blue
    write-message "LogFileReset:         $LogFileReset"        -fore Blue
    write-message "LogFileEmail:         $LogFileEmail"        -fore Blue
    write-message "IncludeVARALog:       $IncludeVARALog"      -fore Blue
    write-message "========= PARAMETERS FOR INSTALLATION DETAILS =============" -fore Magenta
    write-message "RMSPacketInstalled:   $RMSPacketInstalled"  -fore Blue
    write-message "RMSPacketDirectory:   $RMSPacketDirectory"  -fore Blue
    write-message "RMSPacketProcess:     $RMSPacketProcess"    -fore Blue
    write-message "RMSPacketSettings:    $RMSPacketSettings"   -fore Blue
    write-message "RMSTrimodeInstalled:  $RMSTrimodeInstalled" -fore Blue
    write-message "RMSTrimodeDirectory:  $RMSTrimodeDirectory" -fore Blue
    write-message "RMSTrimodeProcess:    $RMSTrimodeProcess"   -fore Blue
    write-message "RMSTrimodeSettings:   $RMSTrimodeSettings"  -fore Blue
    write-message "RMSRelayInstalled:    $RMSRelayInstalled"   -fore Blue
    write-message "RMSRelayDirectory:    $RMSRelayDirectory"   -fore Blue
    write-message "RMSRelayProcess:      $RMSRelayProcess"     -fore Blue
    write-message "RMSRelaySettings:     $RMSRelaySettings"    -fore Blue
    write-message "VARAFMInstalled:      $VARAFMInstalled"     -fore Blue
    write-message "VARAFMDirectory:      $VARAFMDirectory"     -fore Blue
    write-message "VARAFMProcess:        $VARAFMProcess"       -fore Blue
    write-message "VARAFMSettings:       $VARAFMSettings"      -fore Blue
    write-message "VARAFMLog:            $VARAFMLog"           -fore Blue
    write-message "VARAInstalled:        $VARAInstalled"       -fore Blue
    write-message "VARADirectory:        $VARADirectory"       -fore Blue
    write-message "VARAProcess:          $VARAProcess"         -fore Blue
    write-message "VARASettings:         $VARASettings"        -fore Blue
    write-message "VARALog:              $VARALog"             -fore Blue
    write-message "========= PARAMETERS FOR CHECKS TO PERFORM ================" -fore Magenta
    write-message "CheckLogErrors:       $CheckLogErrors"      -fore Blue
    write-message "CheckVARACPU:         $CheckVARACPU"        -fore Blue
    write-message "CheckPort:            $CheckPort"           -fore Blue
    write-message "RestartOnLogError:    $RestartOnLogError"   -fore Blue
    write-message "RestartOnError:       $RestartOnError"      -fore Blue
    write-message "RestartOnLowCPU:      $RestartOnLowCPU"     -fore Blue
    write-message "RestartOnPortError:   $RestartOnPortError"  -fore Blue
    write-message "VARAFMCPURestart:     $VARAFMCPURestart"    -fore Blue
    write-message "VARACPURestart:       $VARACPURestart"      -fore Blue
    write-message "EmailOnError:         $EmailOnError"        -fore Blue
    write-message "========= PARAMETERS FOR LOG AND SETTINGS FILES ===========" -fore Magenta
    write-message "LogRotate:            $LogRotate"           -fore Blue
    write-message "SettingsBackup:       $SettingsRestore"     -fore Blue
    write-message "ForceBackup:          $ForceBackup"         -fore Blue
    write-message "SettingsRestore:      $SettingsRestore"     -fore Blue
    write-message "========= PARAMETERS FOR VOLUME ADJUSTMENTS ===============" -fore Magenta
    write-message "AdjustSpeakerVolume:  $AdjustSpeakerVolume" -fore Blue 
    write-message "SpeakerVolume:        $SpeakerVolume"       -fore Blue
    write-message "========= PARAMETERS FOR ACTIONS TO PERFORM ===============" -fore Magenta
    write-message "StatusRMSPacket:      $StatusRMSPacket"     -fore Blue
    write-message "StatusRMSTrimode:     $StatusRMSTrimode"    -fore Blue
    write-message "StatusRMSPRelay:      $StatusRMSRelay"      -fore Blue
    write-message "StartRMSPacket:       $StartRMSPacket"      -fore Blue
    write-message "StartRMSTrimode:      $StartRMSTrimode"     -fore Blue
    write-message "StartRMSRelay:        $StartRMSRelay"       -fore Blue
    write-message "RestartRMSPacket:     $RestartRMSPacket"    -fore Blue
    write-message "RestartRMSTrimode:    $RestartRMSTrimode"   -fore Blue
    write-message "RestartRMSRelay:      $RestartRMSRelay"     -fore Blue
    write-message "StopRMSPacket:        $StopRMSPacket"       -fore Blue
    write-message "StopRMSTrimode:       $StopRMSTrimode"      -fore Blue
    write-message "StopRMSRelay:         $StopRMSRelay"        -fore Blue
}

##
## MAIN LOGIC
##

## Modules needed to import
Import-Modules

# Create Event Log Source if not already done
if ($EventLogger -eq $True) {
    if ([System.Diagnostics.EventLog]::SourceExists($EventSource) -eq $false) {
        [System.Diagnostics.EventLog]::CreateEventSource($EventSource, "Application")
    }
}

# Add date/time stamp to log file when LogFileWrite set to true
if ($LogFileWrite -eq $true) {
    # Clear out log file when parameter LogFileReset set to true
    if ($LogFileReset -eq $True) {
        Clear-Content $LogFile
        Write-Output "" > $LogFile
    }
    [string]$timeString = [DateTime]::Now.ToString("yyyyMMdd-HHmmss") 
    Write-Output "Script Start Time:    " $timeString >> $LogFile
}

# Check and validate that our basic directory/file/process/settings parameters all exist
write-message "========= CHECKING PROGRAM PARAMETERS =====================" -fore Magenta

Check-Parameters

# Global Variables, track these for email
$VARALogFileEmail = "$VARADirectory$VARALog"
$VARAFMLogFileEmail = "$VARAFMDirectory$VARAFMLog"

# Read the various INI file settings from programs installed
write-message "========= GETTING SYSTEM SERIAL COM PORTS =================" -fore Magenta

Get-COM-Ports

# Read the various INI file settings from programs installed
write-message "========= GETTING SYSTEM AUDIO DEVICES ====================" -fore Magenta

Get-AudioPlayback
Get-AudioRecording

# Read the various INI file settings from programs installed
write-message "========= GETTING APPLICATION INI SETTINGS ================" -fore Magenta

Get-INI-Settings

# Backup our current Winlink RMS & VARA FM settings if SettingsBackup set to true
if ($SettingsBackup -eq $true) {
    write-message "========= SETTINGS FILES BACKUP ===========================" -fore Magenta
    Settings-Backup
}

# Adjust our default system sound speaker device to SpeakerVolume if AdjustSpeakerVolume set to true
if ($AdjustSpeakerVolume -eq $true) {
    write-message "========= SETTING THE DEFAULT DEVICE VOLUME ===============" -fore Magenta
    write-message "Setting volume to:    $SpeakerVolume on default audio device" -fore Yellow
    Set-Speaker -Volume $SpeakerVolume
} else {
    write-message "========= NOT SETTING THE DEFAULT DEVICE VOLUME ===========" -fore Magenta
}

if ($StatusRMSPacket -eq $true -or $StatusRMSTrimode -eq $true -or $StatusRMSRelay -eq $true) {
    write-message "========= CHECKING PROCESSES ==============================" -fore Magenta

    # Check the status of the RMS/VARA processes if Status set to true
    if ($StatusRMSPacket -eq $true) {
        if ($RMSPacketInstalled -eq $true) {
            $OnError = Check-StatusRMSPacket
            # If RestartOnError set to true and Check-StatusRMS returns true in OnError, Restart the processes
            if ($RestartOnError -eq $true -and $OnError -eq $true) {
                write-message "========= RESTARTING ON ERROR IN RMS PACKET ===============" -fore red
                Restart-ProcessRMSPacket
            }
        }
    }

    if ($StatusRMSTrimode -eq $true) {
        if ($RMSTrimodeInstalled -eq $true) {
            $OnError = Check-StatusRMSTrimode
            # If RestartOnError set to true and Check-StatusRMS returns true in OnError, Restart the processes
            if ($RestartOnError -eq $true -and $OnError -eq $true) {
                write-message "========= RESTARTING ON ERROR IN RMS TRIMODE ==============" -fore red
                Restart-ProcessRMSTrimode
            }
        }
    }

    if ($StatusRMSRelay -eq $true) {
        if ($RMSRelayInstalled -eq $true) {
            $OnError = Check-StatusRMSRelay
            # If RestartOnError set to true and Check-StatusRMS returns true in OnError, Restart the processes
            if ($RestartOnError -eq $true -and $OnError -eq $true) {
                write-message "========= RESTARTING ON ERROR IN RMS RELAY ================" -fore red
                Restart-ProcessRMSRelay
            }
        }
    }
}

if ($CheckLogErrors -eq $true) {
    write-message "========= CHECKING PROCESS LOG ERRORS CONDITIONS ==========" -fore Magenta

    # Check the VARA Log file for any typical error strings if CheckLogErrors set to true
    if ($CheckLogErrors -eq $true) {
        $OnLogError = Check-LogErrors
        # If RestartOnLogError set to true and Check-LogErrors returns true in OnLogError, Restart processes
        if ($RestartOnLogError -eq $true -and $OnLogError -eq $true) {
            write-message "========= RESTARTING ON VARA LOG ERROR ====================" -fore red
            if ($RMSPacketInstalled -eq $true) {
                Restart-ProcessRMSPacket
            }
            if ($RMSTrimodeInstalled -eq $true) {
                Restart-ProcessRMSTrimode
            }
        }
    }
}

if ($CheckVARACPU -eq $true) {
    write-message "========= CHECKING PROCESS CPU USE ========================" -fore Magenta

    # Check the VARA FM CPU usage and restart if under the threshold parameter
    if ($VARAFMInstalled -eq $true) {
        $OnLowCPUError = Check-VARAFMCPU
        # If RestartOnLogError set to true and Check-LogErrors returns true in OnLogError, Restart processes
        if ($RestartOnLogError -eq $true -and $OnLowCPUError -eq $true) {
            write-message "========= RESTARTING ON VARA FM LOW CPU USE ===============" -fore red
            Restart-ProcessRMSPacket
        }
    }

    # Check the VARA CPU usage and restart if under the threshold parameter
    if ($VARAInstalled -eq $true) {
        $OnLowCPUError = Check-VARACPU
        # If RestartOnLogError set to true and Check-LogErrors returns true in OnLogError, Restart processes
        if ($RestartOnLogError -eq $true -and $OnLowCPUError -eq $true) {
            write-message "========= RESTARTING ON VARA LOW CPU USE ==================" -fore red
            Restart-ProcessRMSTrimode
        }
    }
}

if ($CheckPort -eq $true) {
    write-message "========= CHECKING PROCESS SOCKET PORT USE ================" -fore Magenta

    # Check the RMS Packet Port Use and restart if issue
    if ($RMSPacketInstalled -eq $true) {
        $OnPortError = Check-RMSPacketPort
        # If RestartOnLogError set to true and Check-LogErrors returns true in OnLogError, Restart processes
        if ($RestartOnPortError -eq $true -and $OnPortError -eq $true) {
            write-message "========= RESTARTING ON RMS PACKET ON PORT USE ============" -fore red
            Restart-ProcessRMSPacket
        } else {
            if ($VARAFMInstalled -eq $true) {
                $OnPortError = Check-VARAFMPort
                # If RestartOnLogError set to true and Check-LogErrors returns true in OnLogError, Restart processes
                if ($RestartOnPortError -eq $true -and $OnPortError -eq $true) {
                    write-message "========= RESTARTING ON VARA FM ON PORT USE ===============" -fore red
                    Restart-ProcessRMSPacket
                }
            }
        }
    }

    # Check the RMS Trimode Port Use and restart if issue
    if ($RMSTrimodeInstalled -eq $true) {
        $OnPortError = Check-RMSTrimodePort
        # If RestartOnLogError set to true and Check-LogErrors returns true in OnLogError, Restart processes
        if ($RestartOnPortError -eq $true -and $OnPortError -eq $true) {
            write-message "========= RESTARTING ON RMS TRIMODE ON PORT USE ===========" -fore red
            Restart-ProcessRMSTrimode
        } else {
            if ($VARAInstalled -eq $true) {
                $OnPortError = Check-VARAPort
                # If RestartOnLogError set to true and Check-LogErrors returns true in OnLogError, Restart processes
                if ($RestartOnPortError -eq $true -and $OnPortError -eq $true) {
                    write-message "========= RESTARTING ON VARA ON PORT USE ==================" -fore red
                    Restart-ProcessRMSTrimode
                }
            }
        }
    }

    # Check the RMS Relay Port Use and restart if issue
    if ($RMSRelayInstalled -eq $true) {
        $OnPortError = Check-RMSRelayPort
        # If RestartOnLogError set to true and Check-LogErrors returns true in OnLogError, Restart processes
        if ($RestartOnPortError -eq $true -and $OnPortError -eq $true) {
            write-message "========= RESTARTING ON RMS RELAY ON PORT USE =============" -fore red
            Restart-ProcessRMSRelay
        }
    }
}

if ($StopRMSPacket -eq $true -or $StopRMSTrimode -eq $true -or $StopRMSRelay -eq $true) {
    write-message "========= STOPPING PROCESSES ==============================" -fore Magenta

    # Stop the processes if Stop set to true
    if ($StopRMSPacket -eq $true) {
        if ($RMSPacketInstalled -eq $true) {
            Stop-ProcessRMSPacket
        }
    }

    # Stop the processes if Stop set to true
    if ($StopRMSTrimode -eq $true) {
        if ($RMSTrimodeInstalled -eq $true) {
            Stop-ProcessRMSTrimode
        }
    }

    # Stop the processes if Stop set to true
    if ($StopRMSRelay -eq $true) {
        if ($RMSRelayInstalled -eq $true) {
            Stop-ProcessRMSRelay
        }
    }

}


if ($RestartRMSPacket -eq $true -or $RestartRMSTrimode -eq $true -or $RestartRMSRelay -eq $true) {
    write-message "========= RESTARTING PROCESSES ============================" -fore Magenta

    # Stop and then Start the processes if Restart set to true
    if ($RestartRMSPacket -eq $true) {
        if ($RMSPacketInstalled -eq $true) {
            Restart-ProcessRMSPacket
        }
    }

    # Stop and then Start the processes if Restart set to true
    if ($RestartRMSTrimode -eq $true) {
        if ($RMSPacketInstalled -eq $true) {
            Restart-ProcessRMSTrimode
        }
    }

    # Stop and then Start the processes if Restart set to true
    if ($RestartRMSRelay -eq $true) {
        if ($RMSRelayInstalled -eq $true) {
            Restart-ProcessRMSRelay
        }
    }

}


if ($StartRMSPacket -eq $true -or $StartRMSTrimode -eq $true -or $StartRMSRelay -eq $true) {
    write-message "========= STARTING PROCESSES ==============================" -fore Magenta

    # Start the processes if Start set to true
    if ($StartRMSPacket -eq $true) {
        if ($RMSPacketInstalled -eq $true) {
            Start-ProcessRMSPacket
        }
    }

    # Start the processes if Start set to true
    if ($StartRMSTrimode -eq $true) {
        if ($RMSTrimodeInstalled -eq $true) {
            Start-ProcessRMSTrimode
        }
    }

    # Start the processes if Start set to true
    if ($StartRMSRelay -eq $true) {
        if ($RMSRelayInstalled -eq $true) {
            Start-ProcessRMSRelay
        }
    }

}

# Send Email and Create Windows Event Log for Process Errors
if ($OnError -eq $True) {
    if ($EmailOnError -eq $True) {
        Send-Email-Error -subject "Run RMS Script - ERROR - Process Not Running" -body "The Winlink RMS Packet or VARA FM process was not running.  Please check log files for more details."
    }
    if ($EventLogger -eq $True) {
        Write-EventLog-Error -Message "Run RMS Script - ERROR - Process Not Running" -EventID 2001
    }
}

# Send Email and Create Windows Event Log for VARA FM Log File Errors
if ($OnLogError -eq $True) {
    if ($EmailOnError -eq $True) {
        Send-Email-Error -subject "Run RMS Script - ERROR - VARA FM Log Errors" -body "The VARA FM log file has some error detected in it.  Please check log files for more details."
    } 
    if ($EventLogger -eq $True) {
        Write-EventLog-Error -Message "Run RMS Script - ERROR - VARA FM Log Errors" -EventID 2002
    }
}

# Send Email and Create Windows Event Log for VARA FM Low CPU Errors
if ($OnLowCPUError -eq $True) {
    if ($EmailOnError -eq $True) {
        Send-Email-Error -subject "Run RMS Script - ERROR - VARA FM Low CPU Error" -body "The VARA FM process is using a lower than expected amount of CPU processing time.  Please check log files for more details."
    } 
    if ($EventLogger -eq $True) {
        Write-EventLog-Error -Message "Run RMS Script - ERROR - VARA FM Low CPU Error" -EventID 2003
    }
}

# Script completed
write-message "========= SCRIPT COMPLETED ================================" -fore Magenta

# Send Email when LogFileEmail set to true, or on errors above
if ($LogFileEmail -eq $True -or 
        ($EmailOnError -eq $true -and
            ($OnError -eq $true -or
             $OnLogError -eq $true -or
             $OnLowCPUError -eq $true -or
             $OnPortError -eq $true
            )
        )
   ) {
    Send-Email-Log
}

# Exit with error code 2001 if process error is set
if ($OnError -eq $True) {
    exit 2001
}

# Exit with error code 2002 if log file error is set
if ($OnLogError -eq $True) {
    exit 2002
}

# Exit with error code 2003 if low CPU error is set
if ($OnLowCPUError -eq $True) {
    exit 2003
}

# Exit with error code 2004 if Port in use error
if ($OnPortError -eq $True) {
    exit 2004
}

# If we get all the way down here, exit successfully with 0 status
exit 0

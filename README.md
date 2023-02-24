# RunRMS
PowerShell script to run Winlink RMS Packet, RMS Trimode and RMS Relay programs along with VARA HF and VARA FM

https://github.com/w3ctk/RunRMS

This is a script to check and run the various RMS sysop processes on Windows

## Running

To run, you can launch Windows PowerShell and run the command from 
the directory you downloaded the RunRMS.ps1 script into:

<code>./runrms</code>

## PowerShell

You can download PowerShell for Windows here if it isn't already installed:

https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows

## Script Setup and Parameters

All of the allowed parameters are set in the script, and these are
displayed when the script is run.  To override any of these settings,
you can add them into the command line parameters, like

### Example Command Line Parameters
<code>./runrms -RMSPacketDirectory "C:\RMS Packet\" -RestartRMSPacket $True -LogFileEmail $True</code>

### Example Command Output
<code>
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
</code>

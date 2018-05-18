#######################################################################################################################
## 
## name:        
##      robocopy-working-files.ps1
##
##      powershell script to run robocopy and 7z to backup the files folder to E and H
##
##      C:\files\apps\it-powershell\2018-05-16-robocopy-working-files\robocopy-working-files.ps1
##
## syntax:
##      .\robocopy-working-files.ps1
##
## dependencies:
##      windows task to run this every day 
##      7zip installed on source PC
##      script needs elevated permissions to run robocopy
##
## updated:
##      -- Wednesday, May 16, 2018 3:06 PM converted from .bat to .ps1
##
## todo:
##

## Functions ##########################################################################################################

##
## LogWrite - write messages to log file 
##

Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring 
}


## Main Code ##########################################################################################################

try {

##                      
## set local code path and initialize settings file 
##
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
[xml]$ConfigFile = Get-Content "$myDir\Settings.xml"

## setup the logfile
$LogDir = $myDir + "\logs"
if(-not ([IO.Directory]::Exists($LogDir))) {New-Item -ItemType directory -Path $LogDir}
$Logfile = ($LogDir + "\robocopy_working_files-" + $(get-date -f yyyy-MM-dd-HHmmss) + ".log")
echo "results are logged to:  "$Logfile 
LogWrite ("Started at:  " + $(get-date -f yyyy-MM-dd-HHmmss))
$date1 = Get-Date

##
## Get variables from the settings.xml file 
##
$roboSourceDir = $ConfigFile.robocopy_working_files.roboSourceDir  
$roboDestDir   = $ConfigFile.robocopy_working_files.roboDestDir    
$roboLogDir    = $ConfigFile.robocopy_working_files.roboLogDir     
$roboLogFile   = $ConfigFile.robocopy_working_files.roboLogFile    
$exePath7z     = $ConfigFile.robocopy_working_files.exePath7z      
$ziplist       = $ConfigFile.robocopy_working_files.ziplist        
$zipDestDir    = $ConfigFile.robocopy_working_files.zipDestDir     
$netDestDir    = $ConfigFile.robocopy_working_files.netDestDir     

LogWrite ("roboSourceDir :  " + $roboSourceDir)
LogWrite ("roboDestDir   :  " + $roboDestDir  )
LogWrite ("roboLogDir    :  " + $roboLogDir   )
LogWrite ("roboLogFile   :  " + $roboLogFile  )
LogWrite ("exePath7z     :  " + $exePath7z    )
LogWrite ("ziplist       :  " + $ziplist      )
LogWrite ("zipDestDir    :  " + $zipDestDir   )
LogWrite ("netDestDir    :  " + $netDestDir   )

$roboLogFile = $roboLogFile + $(get-date -f yyyy-MM-dd-HHmmss) + ".txt"
LogWrite ("roboLogFile   :  " + $roboLogFile   + " <-- with datetime")

##
## use a "here string" aka "splat operator", insert the parameters into the robocopy command string
##
## robocopy c:/files/ e:/files/ /mir /zb /copyall /np /r:1 /w:1 /xd Recycler "System Volume Information" /log+:c:\files\apps\it-powershell\2018-05-16-robocopy-working-files\logs\wfiles-log-20180516_163500.txt  
##          ^^^^^^^^^ ^^^^^^^^^                                                                                ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^ ^^^^^^^^^^^^^^^ ^^^  
##          $rsd      $rdd                                                                                     $rld                                                               $rlf 
##
$command = @"
robocopy {0} {1} /mir /zb /copyall /np /r:1 /w:1 /xd Recycler "System Volume Information" /log+:{2}{3}
"@ -f $roboSourceDir, $roboDestDir, $roboLogDir, $roboLogFile  

echo "--------------------------------------"
$command
echo "--------------------------------------"
LogWrite ("command               :  " + $command)

Invoke-Expression -Command:$command -OutVariable out | Tee-Object -Variable out
LogWrite ("output                :  " + $out)

##
## use a "here string" aka "splat operator", insert the parameters into the 7zip command string
##
## 7zip.exe  a -tzip -y destfile  sourcedir
## ^^^^^^^^^            ^^^^^^^^^ ^^^^^^^^^
## $ep7                 $zdffp    $zf        
##

$zipArray = New-Object string[] 10
$zipArray = $ziplist -split ","

foreach($zipableFolder in $zipArray) { 
    ## build the zip file name by taking the name of the branch folder
    ##   e:\files\textpad               -->  working-files-textpad.zip
    ##            ^^^^^^^

    $zipDestFile = $zipableFolder -replace '.*\\'          
    $zipDestFileFullPath = $zipDestDir + "working-files-" + $zipDestFile + ".zip"

    ## remove prior zip
    if (Test-Path $zipDestFileFullPath) { Remove-Item $zipDestFileFullPath }

    $quoteChar = [char]34

$command = @"
&{0}{1}{0} a -tzip -y -pSFX {2} {3}
"@ -f $quoteChar, $exePath7z, $zipDestFileFullPath, $zipableFolder

    echo "--------------------------------------"
    $command
    echo "--------------------------------------"
    LogWrite ("command               :  " + $command)

    Invoke-Expression -Command:$command -OutVariable out | Tee-Object -Variable out
    LogWrite ("output                :  " + $out)   
}

##
## copy the zip files in the zip list to the network drive
##
foreach($zipableFolder in $zipArray) { 
    ## build the zip file name by taking the name of the branch folder
    ## same as above:
    ##   e:\files\textpad               -->  working-files-textpad.zip
    ##            ^^^^^^^

    $zipDestFile = $zipableFolder -replace '.*\\'          
    $zipDestFileFullPath = $zipDestDir + "working-files-" + $zipDestFile + ".zip"

    $netDestFileFullPath = $netDestDir + "working-files-" + $zipDestFile + ".zip"

    echo "--------------------------------------"
    echo "move $zipDestFileFullPath to $netDestFileFullPath"
    echo "--------------------------------------"
    LogWrite ("move $zipDestFileFullPath to $netDestFileFullPath")

    ## copy zip file to network drive
    if (Test-Path $zipDestFileFullPath) { Move-Item $zipDestFileFullPath $netDestFileFullPath -force}
}

##
## need to cleanup log files after 1 month
##
$logRetention = (Get-Date).AddDays(-30)
LogWrite ("purging log files older than   :  " + $($logRetention) )
Get-ChildItem -Path $roboLogDir -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $logRetention } | Remove-Item -Force


throw ("Halted.  This is the end.  Who knew.")


}
Catch {
    ##
    ## log any error
    ##    
    LogWrite $Error[0]
}
Finally {

    ##
    ## go back to the software directory where we started
    ##
    set-location $myDir

    LogWrite ("finished at:  " + $(get-date -f yyyy-MM-dd-HHmmss))
}
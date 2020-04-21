<#
.SYNOPSIS
    Script to repair the MECM\SCCM client and set windows updates.
.DESCRIPTION
    Prompts the user for a selction to run against a single computer or a list of computers.
    Script will then uninstall the MECM\SCCM client if it is installed, stop all windows update services,
    rename the SoftwareDistribution folder and the Catroot2 folders, restart all windows update servics,
    and reinstall/install the MECM\SCCM client.
.NOTES
    Author: Nick Ortiz
    Created: 2/25/2020
#>
# MECM\SCCM Client isntall file location on network share
$SCCMClientLoc = '\\sccm\cm_dsl$\ConfigMgrClientHealth\client'

#Add .NET Assebmlies for Open file dialog
    Add-Type -AssemblyName System.Windows.Forms

#File Dialog
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('MyComputer') }
    <#
    .NOTES Commented out because no longer used in simplified script NO 4/20/2020
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    #>

<#
.NOTES
Display description and Prompt for Input from user.
#>
Write-Host "DESCRIPTION: `n Use this script to repair the MECM\SCCM client and reset Software Distribution folder.`r 
This should fix clients that are in unkown status with a Client Check Passed/Active message for monitoring software update deployments.`n
INSTRUCTIONS:`n
Select 1 to run client repair on a single computer or comma seperated list.`r
Select 2 to run client repair on a list of computers.
" -ForegroundColor DarkYellow
Write-Host "Enter selection 1 or 2 and press Enter:" -ForegroundColor Yellow
$Sel = Read-Host
switch ($Sel) {
    1 {
        $Comps = (Read-Host -Prompt "Enter computer(s) seperated by commas")
        $Comps = $Comps.Split(',').Trim(' ,')
    }
    2 {
        <#
        .NOTES
        Specify path of Computer .TXT list on the local computer
        #>
        Write-Host "Use dialog box to browse to location of computer list" -ForegroundColor Yellow
        $null = $FileBrowser.ShowDialog()
        $CompList = $FileBrowser.FileName
        $comps = Get-Content $CompList
    }
    Default {
        Write-Host "Not a valid selection! Exiting script..."
        Exit
    }
}
<#
.NOTES
Copy SCCM source files locally from SCCM share
#>
<#
.Notes Commented out to simplify process NO 4-21-2020
    Write-Host "Select local path for SCCM Client source files" -ForegroundColor Yellow
    $null = $FolderBrowser.ShowDialog()
    $localPath = $FolderBrowser.SelectedPath
#>
$localPath = "C:\SCCMClientRepair"
$localSource = $localPath + '\*'
$ClientEXE = $localPath + '\client\ccmsetup.exe'
if (Test-Path $ClientEXE) {
    Write-Host "Found SCCM Client Files in $localPath" -ForegroundColor Yellow
}
elseif (Test-Path -Path $localPath) {
    Write-Host "SCCM Client Files Not Found Copying from $SCCMClientLocl..." -ForegroundColor Yellow
    Copy-Item -Path $SCCMClientLoc -Destination $localPath -Force -Recurse
}
else {
    New-Item -Path $localPath -ItemType Directory
    Write-Host "SCCM Client Files Not Found Copying from $SCCMClientLocl..." -ForegroundColor Yellow
    Copy-Item -Path $SCCMClientLoc -Destination $localPath -Force -Recurse
}


<#
.NOTES
Enter location for .TXT output file on local computer
#>

<#
.NOTES Commented out to simplify process NO 4/21/2020
    Write-Host "Use dialog to specify the output file" -ForegroundColor Yellow
    $null = $FileBrowser.ShowDialog()
    $outfile = $FileBrowser.FileName
#>
$outfoldername = (get-date -Format MM-dd-yyyyTHH.MM.ss).ToString()
$outfolder = $localPath + "\$outfoldername"
$outfile = $outfolder + "\ClientRepairLog.csv"
if (Test-Path $outfile) {}
else {
    New-Item -Path $outfolder -ItemType Directory | Out-Null
    New-Item -Path $outfile -ItemType File | Out-Null
}
Add-Content $outfile "DeviceName,Date,Status,SoftwareDistributionStatus,CatrootStatus"
<#
.NOTES
Prompt for WA credentials from user
#>
Write-Host "Enter WA credentials in dialog box..." -ForegroundColor Yellow
$cred = Get-Credential

<#
.NOTES
Start the process for sccm client uninstall and Windows update repair.
#>
    foreach($comp in $comps){
        <#
        .NOTES
        Start Foreach Comp loop against txt list of computer device names
        #>
        
        try {
            <#
            .NOTES
            Begin error capture using Try/Catch
            #>     
       
        
        if(Test-Connection $comp -Count 1) {
            <#
            .NOTES
            Start check for C:\ClientHealth folder on destinotion computer
            #>
            Invoke-Command -ComputerName $comp -Credential $cred -ScriptBlock{
                $Dest = 'C:\ClientHealth'
                
            if (Test-Path $Dest) {
                <#
                .NOTES
                If Destination exists do nothing and continue to next step
                #>
            }
            else {
                <#
                .NOTES
                Else if destination doesn't exist create C:\ClientHealth folder
                #>
                New-Item -Path $Dest -ItemType Directory -Force
            }
            <#
            .NOTES
            End Check for C:\ClientHealth
            #>
        }
        <#
        .NOTES
        Declare remote Destination folder \\CompName\C$\ClientHealth
        #>
        $Dest = '\\' + $comp + '\c$\ClientHealth'
            <#
            .NOTES
            Create PSDrive J: mapped to \\CompName\C$\ClientHealth
            and copy Client source files from Local PC to Destination PC
            #>
            Write-Host "Creating PSDrive to $Dest" -ForegroundColor Yellow
            New-PSDrive -Name "J" -PSProvider FileSystem -Root $Dest -Credential $cred -Persist -ErrorAction Stop
            Write-Host "Copying Client source files to $Dest" -ForegroundColor Yellow
            copy-item -path $localSource -Destination "J:\" -Force -Recurse            
            Write-Host "Copy completed. Removing PSDrive to $Dest" -ForegroundColor Yellow
            Remove-PSDrive -Name J
            <#
            .NOTES
            End Copy of client source files from Local PC to Destination PC and
            Remove PSDrive J:
            #>
    
    invoke-command -Credential $cred -cn $comp -ScriptBlock {
        <#
        .NOTES
        Start invoke Command Script Block with Captured Credentials on DeviceName
        First Step is to check for SCCM clinet and Uninstall if exists
        #>
        if ((Get-Service -Name "ccmexec" -ErrorAction SilentlyContinue) -ne $null) {
            Write-Host "Starting SCCM Client Uninstall on $using:comp" -ForegroundColor Yellow
            start-process "C:\ClientHealth\Client\ccmsetup.exe" -ArgumentList "/uninstall" -Wait
            while ((Get-Service -Name "ccmexec" -ErrorAction SilentlyContinue) -ne $null) {
                Start-sleep -Seconds 10            
            }            
        }
        else {
            Write-Host "SCCM Client not installed on $using:comp" -ForegroundColor Yellow
        }
        
        <#
        .NOTES
        Stop Wuauserv and kill process if wuauserv gets stuck in StopPending State.
        #>
        $Wuauserv = Get-Service -name "wuauserv"
        Stop-Service -name "wuauserv" -Force | Out-Null
        $Services = Get-WmiObject -Class win32_service -Filter "state = 'stop pending'"
        if ($Services) {
            foreach ($service in $Services) {
                try {
                    Stop-Process -Id $service.processid -Force -PassThru -ErrorAction Stop
                }
                catch {
                    Write-Warning -Message "Unexpected Error. Error details: $_.Exception.Message"
                }
            }
        }           
        $Wuauserv.WaitForStatus("Stopped")
        Write-Host "Stopped wuauserv on $using:comp" -ForegroundColor Green
        <#
        .NOTES
        Stop CryptSvc and kill process if CryptSvc gets stuck in StopPending State.
        #>
        $CryptSvc = Get-Service -Name "CryptSvc"
        Stop-Service -Name "CryptSvc" -Force | Out-Null
        $Services = Get-WmiObject -Class win32_service -Filter "state = 'stop pending'"
        if ($Services) {
            foreach ($service in $Services) {
                try {
                    Stop-Process -Id $service.processid -Force -PassThru -ErrorAction Stop
                }
                catch {
                    Write-Warning -Message "Unexpected Error. Error details: $_.Exception.Message"
                }
            }
        }
        $CryptSvc.WaitForStatus("Stopped")
        Write-Host "Stopped CryptSvc $using:comp" -ForegroundColor Green
        <#
        .NOTES
        Stop BITS Service and kill process if BITS gets stuck in StopPending State.
        #>
        $BitsSVC = Get-Service -Name "bits"
        Stop-Service -Name "Bits" -Force | Out-Null
        $Services = Get-WmiObject -Class win32_service -Filter "state = 'stop pending'"
        if ($Services) {
            foreach ($service in $Services) {
                try {
                    Stop-Process -Id $service.processid -Force -PassThru -ErrorAction Stop
                }
                catch {
                    Write-Warning -Message "Unexpected Error. Error details: $_.Exception.Message"
                }
            }
        }
        $BitsSVC.WaitForStatus("Stopped")
        Write-Host "Stopped BITS $using:comp" -ForegroundColor Green
        <#
        .NOTES
        Stop msiserver Service and kill process if msiserver gets stuck in StopPending State.
        #>
        $Msiserver = Get-Service -Name "msiserver"
        Stop-Service -Name "msiserver" -Force | Out-Null
        $Services = Get-WmiObject -Class win32_service -Filter "state = 'stop pending'"
        if ($Services) {
            foreach ($service in $Services) {
                try {
                    Stop-Process -Id $service.processid -Force -PassThru -ErrorAction Stop
                }
                catch {
                    Write-Warning -Message "Unexpected Error. Error details: $_.Exception.Message"
                }
            }
        }
        $Msiserver.WaitForStatus("Stopped")
        Write-Host "Stopped msiserver $using:comp" -ForegroundColor Green
        <#
        .NOTES
        Rename Software Distribution folder. Checks for renamed folder and removes it if it exists.
        #>

        if (Test-Path 'C:\Windows\SoftwareDistributionBackup') {            
            Remove-Item -Path "C:\Windows\SoftwareDistributionBackup" -Force -Recurse            
        }

        Rename-Item -Path "C:\Windows\SoftwareDistribution" -NewName "C:\Windows\SoftwareDistributionBackup" -force

        if (Test-Path 'C:\Windows\SoftwareDistributionBackup') {
            Write-Host "Renamed SoftwareDistribution folder on $using:comp" -ForegroundColor Green
            $SDStatus = "SoftwareDistribution Rename Success"            
        }
        else {
            Write-Host "Rename SoftwareDistribution folder failed on $using:comp" -ForegroundColor Red
            $SDStatus = "SoftwareDistribution Rename Failed"
        }
        <#
        .NOTES
        Rename Catroot2 folder. Checks for renamed folder and removes it if it exists.
        #>
        if (Test-Path "C:\Windows\System32\catroot2Backup") {
            Remove-Item -Path "C:\Windows\System32\catroot2Backup" -Force -Recurse
        }

        Rename-Item -Path "C:\Windows\System32\catroot2" -NewName "C:\Windows\System32\catroot2Backup" -Force

        if (Test-Path "C:\Windows\System32\catroot2Backup") {
            Write-Host "Renamed Catroot2 folder on $using:comp" -ForegroundColor Green
            $CatStatus = "CatRoot2 Rename Success" 
        }
        else {
            Write-Host "Rename CatRoot2 folder failed on $using:comp" -ForegroundColor Red
            $CatStatus = "CatRoot2 Rename Failed"
        }
        <#
        .NOTES
        Restart wuauserv service.
        #>
        Start-Service -name "wuauserv"
        $Wuauserv.WaitForStatus("Running")
        Write-Host "wuauserv restarted on $using:comp" -ForegroundColor Green
        <#
        .NOTES
        Restart CryptSvc Service
        #>
        Start-Service -name "CryptSvc"
        $CryptSvc.WaitForStatus("Running")
        Write-Host "CryptSvc restarted on $using:comp" -ForegroundColor Green
        <#
        .NOTES
        Restart Bits Service
        #>
        Start-Service -Name "Bits"
        $BitsSVC.WaitForStatus("Running")
        Write-Host "Bits restarted on $using:comp" -ForegroundColor Green
        <#
        .NOTES
        Restart msiserver
        #>
        Start-Service -Name "msiserver"
        $Msiserver.WaitForStatus("Running")
        Write-Host "msiserver restarted on $using:comp" -ForegroundColor Green
        <#
        .NOTES
        Start SCCM Client reinstall
        #>
        Write-Host "Starting SCCM Client Install on $using:comp" -ForegroundColor Yellow
        start-process "C:\ClientHealth\Client\ccmsetup.exe" -ArgumentList "/mp:sccm.svrmc.local SMSSITECODE=CSV /forceinstall" -Wait
        while ((Get-Service -Name "ccmexec" -ErrorAction SilentlyContinue) -eq $null) {
            Start-sleep -Seconds 10
        }
        Write-Host "SCCM Client Reinstall Complete on $using:comp" -ForegroundColor Green
        $Time = Get-Date -Format d
        Remove-item 'C:\ClientHealth\Client' -Force -Recurse
        #Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule “{00000000-0000-0000-0000-000000000021}”
        Write-Output "$Env:COMPUTERNAME,$Time,Successfully completed SCCM Client Reset,$SDStatus,$CatStatus"
        
        
} | Out-File -Append $outfile
<#
.NOTES
End Invoke command script block
#>
}
else {
    <#
    .NOTES
    If test conection fails this step is run
    #> 
    $Time = Get-Date -Format d
    Write-Host "$comp offline" -ForegroundColor Red
    Add-Content $outfile "$comp,$time,offline,,"
    
} 
# End Try block
Write-Host "Process Completed See $outfile for results" -ForegroundColor Green

}
Catch {
    $Time = Get-Date -Format d
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Write-Host "Error occured on $comp`r $ErrorMessage`r$FailedItem"
    Add-Content $outfile "$comp,$time,$ErrorMessage,$FailedItem"
    Write-Host "Process Failed See $outfile for results" -ForegroundColor Red
    
}
<#
.NOTES
Commented out for better error capture NO 2/25/2020
catch [System.IO.IOException]{
    #Start Catch
    $Time = Get-Date -Format d
    Write-Host "Network path not found $comp" -ForegroundColor Red
    Add-Content $outfile "$comp,Network Path Not Found,,$Time"
    #End Catch
}
catch {
    #Start Catch
    $Time = Get-Date -Format d
    Write-Host "Unknown Error $comp" -ForegroundColor Red
    Add-Content $outfile "$comp,Unknown Error,,$Time"
    #End Catch
}
#>
#End Foreach Comp
}
Read-Host "Press any key to exit..."
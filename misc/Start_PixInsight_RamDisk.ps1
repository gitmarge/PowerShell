#Requires -RunAsAdministrator

# Quick and dirty PowerShell script that creates RAM disks with ImDisk and then launches PixInsight
# Waits for PixInsight to exit before removing the ram disks
# Created for personal use, to prevent PixInsight to overwrite the specified RAM disk paths should they not exist upon startup

# VARIABLES
$PixInsight_Executable = "C:\Program Files\PixInsight\bin\PixInsight.exe"
$PixInsight_WorkingDirectory = "C:\Program Files\PixInsight\bin\"

$DriveSize = "5G"
$DriveFileSystem = "ntfs"
$DriveLabel = "RamDisk"
$DriveDirectoryName = "Temp"
$DriveLetters = "W", "X", "Y", "Z"

# FUNCTIONS
Function New-RamDisks{
    ForEach($DriveLetter in $DriveLetters){
        imdisk -a -s $DriveSize -m "$($DriveLetter):" -p "/fs:$($DriveFileSystem) /v:$($DriveLabel) /q /y" | Out-Null
        If($LASTEXITCODE -eq 0){
            Write-Host "Successfully created drive '$($DriveLetter):' with parameters: DriveSize $($DriveSize) | DriveFileSystem $($DriveFileSystem) | DriveLabel $($DriveLabel)" -ForegroundColor Green
        }Else{
            Write-Host "Error creating drive '$($DriveLetter):' - LastExitCode '$($LASTEXITCODE)'!" -ForegroundColor Red
            [switch]$initError = $true
        }

        Try{
            New-Item -Path "$($DriveLetter):" -Name $DriveDirectoryName -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }Catch{
            Write-Host "Error creating directory '$($DriveDirectoryName)' in '$($DriveLetter):' - $($Error[0].Exception.Message)" -ForegroundColor Red
            [switch]$initError = $true
        }
    }
}

Function Remove-RamDisks{
    ForEach($DriveLetter in $DriveLetters){
        imdisk -d -m "$($DriveLetter):" | Out-Null
        If($LASTEXITCODE -eq 0){
            Write-Host "Successfully removed drive '$($DriveLetter):'" -ForegroundColor Green
        }Else{
            Write-Host "Error removing drive '$($DriveLetter):' - LastExitCode '$($LASTEXITCODE)'!" -ForegroundColor Red
        }
    }
}

Function Get-Choice{
    $choice = Read-Host "Enter 'x' to start removing drives"
    Switch($choice){
        x {}
        Default {Get-Choice}
    }
}

Function Show-ExitPrompt{
    Write-Host -NoNewLine "`nPress any key to exit ..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

# START
New-RamDisks

If(!$initError){
    Start-Sleep -Seconds 5
    Try{
        Start-Process -FilePath $PixInsight_Executable -WorkingDirectory $PixInsight_WorkingDirectory -ErrorAction Stop
        Write-Host "`nSuccessfully started PixInsight" -ForegroundColor Green
    }Catch{
        Write-Host "`nError starting PixInsight - $($Error[0].Exception.Message)" -ForegroundColor Red
    }

    $PI_ID = (Get-Process PixInsight).Id
    If($null -ne $PI_ID){
        Write-Host "Got PixInsight PID - $($PI_ID)" -ForegroundColor Green
        Write-Host "`n>>> Waiting for PixInsight to exit, please do not close shell ..." -ForegroundColor Yellow
        Wait-Process $PI_ID
        Write-Host "`nPixInsight exited`nSleeping for 10 seconds before removing drives ...`n" -ForegroundColor Yellow
        Start-Sleep -Seconds 10
    }Else{
        Write-Host "Couldn't get PixInsight PID" -ForegroundColor Red
        Write-Host "`nTo start removing drives, enter 'x'. Make sure PixInsight is not running, otherwise close after work." -ForegroundColor Red
        Get-Choice
    }    
    Remove-RamDisks
    Show-ExitPrompt
}Else{
    Write-Host "`nError(s) encountered when creating drives! Proceeding to remove drives ...`n" -ForegroundColor Red
    Remove-RamDisks
    Show-ExitPrompt
}

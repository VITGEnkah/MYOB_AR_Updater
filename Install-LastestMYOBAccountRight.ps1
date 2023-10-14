## 
## 
## 

$Baseurl = "https://download.myob.com/arl/msi/"

$scriptdir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$ProgressPreference = 'SilentlyContinue'

# Log events to MYOB-AR-Downloader.log
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -path ($scriptdir + "\Install-MYOB-AR-Client.log")

# Get the current version
$InstalledApps = Get-ChildItem -Path "HKLM:\SOFTWARE\WOW6432Node\MYOB\AccountRight Client" | Get-ItemProperty |select PSChildName
$current = ($installedApps | Measure-Object -Property PSChildName -Maximum).maximum
Write-host "Latest Installed Version: " $current

# Split the current version into year and month
$year, $month = $current -split '\.'
# Increment the month
$newMonth = [int]$month + 1
# Check if the month exceeds 12 (reset to 1 in that case) and adjust the year accordingly
if ($newMonth -gt 12) {
    $newMonth = 1
    $year = [int]$year + 1
}

# Construct the new version string
$newVersion = "$year.$newMonth"
$clienturl = $Baseurl + "MYOB_AccountRight_Client_" + $newVersion + ".msi"
$apiurl = $Baseurl + "MYOB_AccountRight_API_AddOnConnector_Installer_" + $newVersion + ".msi"

# Check if new version is up for download
try{
    # Use wget to download the file
    $response = wget $clienturl -ErrorAction Stop

    # Get the status code
    $statusCode = $response.StatusCode

    if($statusCode -eq 200) {
        Write-host "New Version Available:" $newVersion ", Downloading Installers.." -ForegroundColor Green
        wget $clienturl -OutFile ($scriptdir + "\MYOB_AccountRight_Client_" + $newVersion + ".msi")
        Unblock-File -Path ($scriptdir + "\MYOB_AccountRight_Client_" + $newVersion + ".msi")
        wget $apiurl -OutFile ($scriptdir + "\MYOB_AccountRight_API_" + $newVersion + ".msi")
        Unblock-File -Path ($scriptdir + "\MYOB_AccountRight_API_" + $newVersion + ".msi")
    }

}catch{
    Write-host "Already Installed, nothing to do.." -ForegroundColor Green
    Return
}

# Start Installing..
$logFile = '{0}-{1}.log' -f $file.fullname,(get-date -Format yyyyMMddTHHmmss)

$MYOB_AccountRight_Setup_msi = (gci -Path $scriptdir -Filter "MYOB_AccountRight_Client_*.msi").name
$MYOB_AccountRight_API_Setup_msi = (gci -Path $scriptdir -Filter "MYOB_AccountRight_API_*.msi").name

Write-host "Installing" $MYOB_AccountRight_Setup_msi  -ForegroundColor Green
$FullPkgPath = $scriptdir + "\" + $MYOB_AccountRight_Setup_msi
$Args = "/i " + $FullPkgPath + " ALLUSERS=1 ACCEPTEULA=1 /qb /norestart /log " + $FullPkgPath + ".log"
write-host "MSIEXEC" $Args
$InstallCMD = Start-Process -NoNewWindow -FilePath MSIEXEC -ArgumentList $Args -wait -passthru

# Check if install worked
If (@(0,3010) -contains $InstallCMD.exitcode) { 
    write-host "Install Successful" 
    
    # Clean-up and remove the MSI
    Remove-Item -Path $FullPkgPath -Force
}else{
    write-error "Something went wrong with the Client MSI Installer. Check the log file. Try uninstalling previous version first"; 
    Return
}

Write-host "Cleaning up Start Menu"
rm -Path ("C:\programdata\Microsoft\Windows\Start Menu\Programs\MYOB\MYOB AccountRight " + $DownloadVersion + "\Tools\") -Recurse -Force
rm -Path ("C:\programdata\Microsoft\Windows\Start Menu\Programs\MYOB\MYOB AccountRight " + $DownloadVersion + "\AccountRight User Guide (AU).lnk") -Force
rm -Path ("C:\programdata\Microsoft\Windows\Start Menu\Programs\MYOB\MYOB AccountRight " + $DownloadVersion + "\AccountRight User Guide (NZ).lnk") -Force

Write-host "Installing" $MYOB_AccountRight_API_Setup_msi  -ForegroundColor Green
$FullPkgPath = $scriptdir + "\" + $MYOB_AccountRight_API_Setup_msi
$Args = "/i " + $FullPkgPath + " ACCEPTEULA=1 /qb /norestart /log " + $FullPkgPath + ".log"
write-host "MSIEXEC" $Args
$InstallCMD = Start-Process -NoNewWindow -FilePath MSIEXEC -ArgumentList $Args -wait -passthru
Start-Sleep 5
If (@(0,3010) -contains $InstallCMD.exitcode){ 
    write-host "Install Successful"
    
    # Clean-up and remove the MSI
    Remove-Item -Path $FullPkgPath -Force 
}else{
    write-error "Something went wrong with the API MSI Installer. Check the log file. Try uninstalling previous version first"
}

Write-host "Generating Registry Keys and Intergrating MYOB AE and AR"  -ForegroundColor Green
# Populate the HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\MYOB\AccountRightSharedServicesConsumers KEY

  $InstalledApps = Get-ChildItem -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Select-Object -Property DisplayName, DisplayVersion, InstallLocation, UninstallString  | Where {($_.DisplayName -like "MYOB AccountRight 20*") -and ($_.UninstallString -like "MsiExec*")}
  $registryPath = "HKLM:\SOFTWARE\WOW6432Node\MYOB\AccountRightSharedServicesConsumers"

  foreach ($app in $InstalledApps){
	$appName = $app.DisplayName.substring(18)
	$appGUID = $app.UninstallString.substring(14) 
	write-host "Linking: " $appName

	IF ($appName -NotLike "*SE") {
		$appName = $appName + " SE"
	}

	IF(!(Test-Path $registryPath))  {
	    New-Item -Path $registryPath -Force | Out-Null
	    New-ItemProperty -Path $registryPath -Name $appName -Value $appGUID -PropertyType String -Force | Out-Null}
	 ELSE {
	    New-ItemProperty -Path $registryPath -Name $appName -Value $appGUID -PropertyType String -Force | Out-Null}}


# Populate the HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Huxley.Application.exe KEY

  $InstalledApps = Get-ChildItem -Path "HKLM:\SOFTWARE\WOW6432Node\MYOB\AccountRight Client" | Get-ItemProperty |select PSChildName,"(Default)"
  $Latest = ($installedApps | Measure-Object -Property PSChildName -Maximum).maximum
  $LatestInstallPath = ($installedApps | Where-Object -Property PSChildName -eq $latest |select -ExpandProperty "(Default)")
  $registryPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Huxley.Application.exe"
  
	IF(!(Test-Path $registryPath))  {
	    New-Item -Path $registryPath -Force | Out-Null
	    New-ItemProperty -Path $registryPath -Name "Path" -Value $LatestInstallPath -PropertyType String -Force | Out-Null
	    New-ItemProperty -Path $registryPath -Name "(Default)"-Value ($LatestInstallPath + "Huxley.Application.exe") -PropertyType String -Force | Out-Null}
	 ELSE {
	    New-ItemProperty -Path $registryPath -Name "Path" -Value $LatestInstallPath -PropertyType String -Force | Out-Null
	    New-ItemProperty -Path $registryPath -Name "(Default)" -Value ($LatestInstallPath + "Huxley.Application.exe") -PropertyType String -Force | Out-Null}

Write-host "Finished..." -ForegroundColor Green
Stop-Transcript
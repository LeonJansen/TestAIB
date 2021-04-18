#######################
# CloudAssist LanguagePack Installation
# Run this script with Admin Permissions

#######################

$systemlocale = "nl-NL"
$languages = @("nl-nl", "de-de")
$timezone = "W. Europe Standard Time"

$LIPContent = "C:\SetupCloudAssist\LanguagePacks" #Root folder where the copied sourcefiles are
$LanguagePackLocation = @{
    "LanguageISO" = "https://software-download.microsoft.com/download/pr/19041.1.191206-1406.vb_release_CLIENTLANGPACKDVD_OEM_MULTI.iso"
    "FOD"         = "https://software-download.microsoft.com/download/pr/19041.1.191206-1406.vb_release_amd64fre_FOD-PACKAGES_OEM_PT1_amd64fre_MULTI.iso"
    "InboxApps"   = "https://software-download.microsoft.com/download/pr/19041.508.200905-1327.vb_release_svc_prod1_amd64fre_InboxApps.iso"
} 

####################
# Download and extract the LanguagePacks
####################
If (!(test-path $LIPContent)) {
    New-Item -ItemType Directory -Force -Path $LIPContent | Out-Null
}
 
try {  
    foreach ($key in $LanguagePackLocation.Keys) {
        $value = $LanguagePackLocation.$key
        Write-Output "Downloading: $key"
    
        $filename = $value.Substring($value.LastIndexOf("/") + 1)

        #First Download all the LanguagePacks
        Invoke-WebRequest -Uri $value -OutFile "$LIPContent\$filename"
    }
}
catch {
    throw "Something went wrong downloading the language packs."
}
    
##Disable Language Pack Cleanup##
Disable-ScheduledTask -TaskPath "\Microsoft\Windows\AppxDeploymentClient\" -TaskName "Pre-staged app cleanup"
    
# Mount ISO FILES loop through all
Get-ChildItem –Path $LIPContent -Recurse -Filter *.iso |
Foreach-Object {
    # Mount ISO File
    $vol = Mount-DiskImage -ImagePath $_.FullName -PassThru | Get-DiskImage | Get-Volume
    $driveletter = $vol.DriveLetter + ":"

    # Loop through all languages that we need to install
    foreach ($lang in $languages) {
        Write-Host "Installing Language: $lang"
        # Check what to do based on filename:
        if ($_.Name -like "*CLIENTLANGPACK*" ) {
            Write-Host "Install LanguageExperiencepack"
            Add-AppProvisionedPackage -Online -PackagePath $driveletter\LocalExperiencePack\$lang\LanguageExperiencePack.$lang.Neutral.appx -LicensePath $driveletter\LocalExperiencePack\$lang\License.xml
            Add-WindowsPackage -Online -PackagePath "$driveletter\x64\langpacks\Microsoft-Windows-Client-Language-Pack_x64_$lang.cab"
        }
        if ($_.Name -like "*FOD-PACKAGES*" ) {
            Write-Host "Install FOD Packages"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-LanguageFeatures-Basic-$lang-Package~31bf3856ad364e35~amd64~~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-LanguageFeatures-Handwriting-$lang-Package~31bf3856ad364e35~amd64~~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-LanguageFeatures-OCR-$lang-Package~31bf3856ad364e35~amd64~~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-LanguageFeatures-Speech-$lang-Package~31bf3856ad364e35~amd64~~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-LanguageFeatures-TextToSpeech-$lang-Package~31bf3856ad364e35~amd64~~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-NetFx3-OnDemand-Package~31bf3856ad364e35~amd64~$lang~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-InternetExplorer-Optional-Package~31bf3856ad364e35~amd64~$lang~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-MSPaint-FoD-Package~31bf3856ad364e35~amd64~$lang~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-Notepad-FoD-Package~31bf3856ad364e35~amd64~$lang~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-PowerShell-ISE-FOD-Package~31bf3856ad364e35~amd64~$lang~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-Printing-WFS-FoD-Package~31bf3856ad364e35~amd64~$lang~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-StepsRecorder-Package~31bf3856ad364e35~amd64~$lang~.cab"
            Add-WindowsPackage -Online -PackagePath "$driveletter\Microsoft-Windows-WordPad-FoD-Package~31bf3856ad364e35~amd64~$lang~.cab"
        }
        if ($_.Name -like "*InboxApps*" ) {
            Write-Host "Install InboxApps"
            $inboxapps = $driveletter + "\amd64fre"
            $AllAppx = Get-Item $inboxapps\*.appx | Select-Object name
            $AllAppxBundles = Get-Item $inboxapps\*.appxbundle | Select-Object name
            $allAppxXML = Get-Item $inboxapps\*.xml | Select-Object name
            foreach ($Appx in $AllAppx) {
                $appname = $appx.name.substring(0, $Appx.name.length - 5)
                $appnamexml = $appname + ".xml"
                $pathappx = $InboxApps + "\" + $appx.Name
                $pathxml = $InboxApps + "\" + $appnamexml
                if ($allAppxXML.name.Contains($appnamexml)) {
                    Add-AppxProvisionedPackage -Online -PackagePath $pathappx -LicensePath $pathxml
                }
                else {
                    Add-AppxProvisionedPackage -Online -PackagePath $pathappx -skiplicense
                }
            }
            foreach ($Appx in $AllAppxBundles) {
                $appname = $appx.name.substring(0, $Appx.name.length - 11)
                $appnamexml = $appname + ".xml"
                $pathappx = $InboxApps + "\" + $appx.Name
                $pathxml = $InboxApps + "\" + $appnamexml
                if ($allAppxXML.name.Contains($appnamexml)) {
                    Add-AppxProvisionedPackage -Online -PackagePath $pathappx -LicensePath $pathxml
                }
                else {
                    Add-AppxProvisionedPackage -Online -PackagePath $pathappx -skiplicense
                }
            }
        }
    }

    # Unmount the volume
    Start-Sleep 10
    $vol | get-diskimage | Dismount-DiskImage

}


# Update Languagelist according installed    
$LanguageList = Get-WinUserLanguageList
foreach ($lang in $languages) {
    $LanguageList.Add($lang)
}
Set-WinUserLanguageList $LanguageList -force
    
#setting system local
Write-Host "$systemlocale - Setting the system locale" -ForegroundColor Green
Set-WinSystemLocale -SystemLocale $systemlocale
Set-TimeZone -Name $timezone
Set-Culture $systemlocale

#remove folder when done
Remove-Item -LiteralPath $LIPContent -Force -Recurse
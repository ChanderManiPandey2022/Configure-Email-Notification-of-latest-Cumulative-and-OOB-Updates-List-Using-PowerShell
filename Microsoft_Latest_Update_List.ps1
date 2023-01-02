<#
.SYNOPSIS
  <Intune Report Automation -Configure Email Notification of latest Cumulative and OOB Updates List Using PowerShell>
.DESCRIPTION
  <Intune Report Automation -Configure Email Notification of latest Cumulative and OOB Updates List Using PowerShell>
.Demo
<YouTube video link-->
.INPUTS
  <Provide all required inforamtion in User Input Section>
.OUTPUTS
  <You will get mail subscrition the attachement in CSV>
.NOTES
  Version:        1.0
  Author:         ChanderMani Pandey
  Creation Date:  2 Jan 2023
  Find Author on 
  Youtube:-        https://www.youtube.com/@chandermanipandey8763
  Twitter:-        https://twitter.com/Mani_CMPandey
  Facebook:-       https://www.facebook.com/profile.php?id=100087275409143&mibextid=ZbWKwL
  LinkedIn:-       https://www.linkedin.com/in/chandermanipandey
  Reddit:-         https://www.reddit.com/u/ChanderManiPandey 

 #>
Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' 

$error.clear() ## this is the clear error history 
cls
$ErrorActionPreference = 'SilentlyContinue';

#######################  User Input Section ##################################################################################
$WorkingFolder = "C:\TEMP\MicroSoftPatchList\LatestCumulativeUpdateList" 
$From          = "abcd@xyz.com"
$To            = "efgh@xyz.com"
$CC            = "igkl@xyz.com"
$Subject       = "Intune Patching- Latest Cumulative Updates List for Windows 10/11 Patch as on $(Get-Date -Format "yyyy-MM-dd")"
$SmtpServer    = "smtp.gmail.com"
$Port          = '587'
$Priority      = "Normal"
    
#######################################################################################################################
$Date = Get-Date -Format "MMMMMMMM dd, yyyy";
$OutFileMP = "$WorkingFolder\MicrosoftPatchList.csv";
$OutFileLP = "$WorkingFolder\MicrosoftLatestPatchList.csv";
$PatchingMonth = "";
$PatchReleaseDays = 0;
# Create an empty array of PSObject objects
$buildInfoArray = @()
#====================================================================================
#Creating working Folder
New-Item -ItemType Directory -Path $WorkingFolder -Force

# Add each Build and Operating System to the array
"22623,Windows 11 22H2","22621,Windows 11 22H2 B1","22471,Windows 11 21H2","22468,Windows 11 21H2 B6","22463,Windows 11 21H2 B5",
"22458,Windows 11 21H2 B4","22454,Windows 11 21H2 B3","22449,Windows 11 21H2 B2","22000,Windows 11 21H2 B1","21996,Windows 11 Dev",
"19045,Windows 10 22H2","19044,Windows 10 21H2","19043,Windows 10 21H1","19042,Windows 10 20H2","19041,Windows 10 2004","19008,Windows 10 20H1",
"18363,Windows 10 1909","18362,Windows 10 1903","17763,Windows 10 1809","17134,Windows 10 1803","16299,Windows 10 1709 FC","15254,Windows 10 1709",
"15063,Windows 10 1703","14393,Windows 10 1607","10586,Windows 10 1511","10240,Windows 10 1507","9600,Windows 8.1",
"7601,Windows 7" | ForEach-Object {
    # Create a new PSObject object
    $buildInfo = New-Object -TypeName PSObject

    # Add the Build and Operating System properties to the object
    $buildInfo | Add-Member -MemberType NoteProperty -Name "Build" -Value ($_ -split ",")[0]
    $buildInfo | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value ($_ -split ",")[1]

    # Add the object to the array
    $buildInfoArray += $buildInfo
}

# Print the array of PSObject objects to the screen
#$buildInfoArray

#===============================================================================================================
$CollectedData = $BuildDetails = $PatchDetails = $MajorBuilds = $LatestPatches = @();

$BuildDetails = $buildInfoArray


#Download Windows Master Patch List
Write-Host "Downoading Patch List from Microsoft"-ForegroundColor yellow
$URI = "https://aka.ms/Windows11UpdateHistory";
$CollectedData += (Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction Continue).Links;
$URI = "https://aka.ms/WindowsUpdateHistory";
$CollectedData += (Invoke-WebRequest -Uri $URI -UseBasicParsing -ErrorAction Continue).Links;

#Filter Windows Master Patch List
Write-Host "Filtering Patch List"-ForegroundColor yellow
IF ($CollectedData) {
	$CollectedData = ($CollectedData | Where-Object {$_.class -eq "supLeftNavLink" -and $_.outerHTML -match "KB" -and $_.outerHTML -notmatch "out-of-band" -and $_.outerHTML -notmatch "preview" -and $_.outerHTML -notmatch "mobile"}).outerHTML
}

#Consolidate the Master Patch and Format the output
Write-Host "Consolidating Patch List"-ForegroundColor yellow
Foreach ($Line in $CollectedData) {
	$ReleaseDate = $PatchID = ""; $Builds = @();	
    $ReleaseDate = (($Line.Split(">")[1]).Split("&”)[0]).trim();
        IF ($ReleaseDate -match "build") {$ReleaseDate = ($ReleaseDate.split("-")[0]).trim();}
	$PatchID = ($Line.Split(" ;-") | Where-Object {$_ -match "KB"}).trim();
    $Builds = ($Line.Split(",) ") | Where-Object {$_ -like "*.*"}).trim();
	Foreach ($BLD in $Builds) {
		$MjBld = $MnBld = ""; $MjBld = $BLD.Split(".")[0]; $MnBld = $BLD.Split(".")[1];
            Foreach ($Line1 in $BuildDetails) {
                $BldNo = $OS = ""; $BldNo = $Line1.Build; $OS = $Line1.OperatingSystem; $MajorBuilds += $BldNo;
                IF ($MjBld -eq $BldNo) {Break;}
                ELSE {$OS = "Unknown";}
            }
            $PatchDetails += [PSCustomObject] @{OperatingSystem = $OS; Build = $BLD; MajorBuild = $MjBld; MinorBuild = $MnBld; PatchID = $PatchID; ReleaseDate = $ReleaseDate;}
       }
}
$MajorBuilds = $MajorBuilds | Select-Object -Unique | Sort-Object -Descending;
$PatchDetails = $PatchDetails | Select-Object OperatingSystem, Build, MajorBuild, MinorBuild, PatchID, ReleaseDate -Unique | Sort-Object MajorBuild,PatchID -Descending;
$PatchDetails | Export-Csv -Path $OutFileMP -NoTypeInformation;

Write-Host "Finalizing Patch List";-ForegroundColor yellow
IF ($PatchingMonth) {
	Foreach ($Bld in $MajorBuilds) {$LatestPatches += $PatchDetails | Where-Object {$_.MajorBuild -eq $Bld -and 
        $_.ReleaseDate -match $PatchingMonth.Year -and $_.ReleaseDate -match $PatchingMonth.Month} | Sort-Object PatchID -Descending | Select-Object -First 1;}
}
ELSE {
    $Today = Get-Date; $LatestDate = ($PatchDetails | Select-Object -First 1).ReleaseDate; $DiffDays = ([datetime]$Today - [datetime]$LatestDate).Days;
    IF ([int]$DiffDays -gt [int]$PatchReleaseDays) {
        Foreach ($Bld in $MajorBuilds) {$LatestPatches += $PatchDetails | Where-Object {$_.MajorBuild -eq $Bld} | Sort-Object PatchID -Descending | Select-Object -First 1;}
    }
    ELSE {
        $Month = ""; $Month = ((Get-Date).AddMonths(-1)).ToString("MMMMMMMM dd, yyyy").Split(" ,") | Select-Object -First 1;
        $Year = ""; $Year = ((Get-Date).AddMonths(-1)).ToString("MMMMMMMM dd, yyyy").Split(" ,") | Select-Object -Last 1;
        $PatchingMonth = [PSCustomObject]@{Month = $Month; Year = $Year;};
        Foreach ($Bld in $MajorBuilds) {$LatestPatches += $PatchDetails | Where-Object {$_.MajorBuild -eq $Bld -and 
            $_.ReleaseDate -match $PatchingMonth.Year -and $_.ReleaseDate -match $PatchingMonth.Month} | Sort-Object PatchID -Descending | Select-Object -First 1;}
        
	    #Adding Latest Patches for Other Builds Missing above
        $M = ""; $M = ((Get-Date).ToString("MMMMMMMM dd, yyyy")).split(" ,") | Select-Object -First 1;
        $Y = ""; $Y = ((Get-Date).ToString("MMMMMMMM dd, yyyy")).split(" ,") | Select-Object -Last 1;
        Foreach ($Bld1 in $MajorBuilds) {
            $Found = 0; Foreach ($Line in $LatestPatches) {$Bld2 = ""; $Bld2 = $Line.MajorBuild; IF ($Bld1 -eq $Bld2) {$Found = 1; Break;}}
            IF ($Found -eq 0) {$LatestPatches += $PatchDetails | Where-Object {$_.MajorBuild -eq $Bld1 -and 
                               $_.ReleaseDate -notlike "$M*$Y"} | Sort-Object PatchID -Descending | Select-Object -First 1;}
        }
    }
}
$LatestPatches | Export-Csv -Path $OutFileLP -NoTypeInformation;




### PARAMS #################################################################################################################################
     
# Function: Array > Html  
Function Create-HTMLTable{ 
    param([array]$Array)      
    $arrHTML = $Array | ConvertTo-Html
    $arrHTML[-1] = $arrHTML[-1].ToString().Replace('</body></html>'," ")    
    Return $arrHTML[5..2000]
}
  
# Define empty array
$Report = @()
$R =  $LatestPatches | Select-Object OperatingSystem,Build,MajorBuild,MinorBuild,PatchID,ReleaseDate 
 
##########################################Report ############################################################################################
Try{
    $Report =     $R                 
                    # Define object with properties
                    $Object = @{} | select "OperatingSystem", 
                                           "Build", 
                                           "MajorBuild", 
                                           "MinorBuild", 
                                           "PatchID", 
                                           "ReleaseDate"
                 
                    
      
                    $Object.OperatingSystem       = $R.OperatingSystem
                    $Object.Build           = $R.Build
                    $Object.MajorBuild          =$R.MajorBuild
                    $Object.MinorBuild     = $R.MinorBuild
                    $Object.PatchID       = $R.PatchID
                    $Object.ReleaseDate  = $R.ReleaseDate 
 
                    # Return object  
                    #$Object
                  
      
     Select-Object DeviceName
}
Catch [System.Exception]{
    Write-host "Error" -backgroundcolor red -foregroundcolor yellow
    $_.Exception.Message
}
      
#Display results
$Report | Format-Table -AutoSize -Wrap
 
 
### HTML REPORT ############################################################################################################################
# Creating head style
$Style = @"
      
    <style>
      body {
        font-family: "Arial";
        font-size: 8pt;
        color: #4C607B;
        }
      th, td { 
        border: 1px solid #005fe5;
        border-collapse: collapse;
        padding: 5px;
        }
      th {
        font-size: 1.2em;
        text-align: left;
        background-color: #003366;
        color: #ffffff;
        }
      td {
        color: #000000;
        }
      .even { background-color: #ffffff; }
      .odd { background-color: #bfbfbf; }
    </style>
      
"@
# Creating head style and header title
$output = $null
$output = @()
$output += '<html><head></head><body>'
$output +=
 
#Import hmtl style file
$Style
 
$output += "<h3 style='color: #0B2161'>Intune Patching</h3>"
$output += '<strong><font color="Green">Informaton: </font></strong>'
$output += "Latest Cumulative Update for Windows 10/11</br>"
$output += '</br>'
$output += '<hr>'
$output += '<p>'
$output += "<h4>Hello All</h4>"
$output += "<h4>Please Find the  Latest Cumulative Update (Patch Tuseday) for Windows 10/11 update list released by Microsoft as on $(Get-Date -Format "yyyy-MM-dd").</h4>"
$output += '<p>'
$output += Create-HTMLTable $Report
$output += '</p>'
$output += '</body></html>'
$output =  $output | Out-String
$output += "<h4>Regards</h4>"
$output += "<h4>Patch Management Team</h4>"
  
### SENDING EMAIL ##########################################################################################################################

 Write-Host "Sending Mail to $to,$cc" -ForegroundColor yellow
#Email Params
$Parameters = @{
    From        = $from
    To          = $To
    Subject     = $Subject
    Body        = $Output
    BodyAsHTML  = $True
    CC          = $CC
    Port        = $Port
    Priority    = $Priority 
    SmtpServer  = $SmtpServer
    Attachments = $OutFileLP 
}

#Sending email
Send-MailMessage @Parameters
#$credential
#-Attachments $data
Write-Host "Mail successfully sent to $to,$cc" -ForegroundColor Green



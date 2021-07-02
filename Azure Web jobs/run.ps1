# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format.
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' property is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Output "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Output "PowerShell timer trigger function ran! TIME: $currentUTCtime"

#$Script={
#    param ()

#$csvFileDir = "C:\CsvData\";
$jsonFileDir = "C:\home\JsonData\";

$OUs = "OU=Central,OU=LM Users,DC=lm,DC=lmig,DC=com", `
"OU=Mid-Atlantic,OU=LM Users,DC=lm,DC=lmig,DC=com", `
"OU=Mid-West,OU=LM Users,DC=lm,DC=lmig,DC=com", `
"OU=New-England,OU=LM Users,DC=lm,DC=lmig,DC=com", `
"OU=Pacific,OU=LM Users,DC=lm,DC=lmig,DC=com", `
"OU=Southern,OU=LM Users,DC=lm,DC=lmig,DC=com", `
"OU=Southwest,OU=LM Users,DC=lm,DC=lmig,DC=com", `
"OU=BlancIR011E,OU=Europe,OU=LM Users,DC=lm,DC=lmig,DC=com", `
"OU=BelfaUK022D,OU=Europe,OU=LM Users,DC=lm,DC=lmig,DC=com";


$logFile = "getSpolBulkProfileLoadDataProd";


$timeRun = get-date -Format hhmmss;
$log = $fileDir + $logFile + $timeRun + ".log";

#todo - get some error handling in this
function errorLog($service, $message)
{
	
    "error " + $service + " " + $message >> $log;
	
}

#clear all working files
#Remove-Item $jsonFileDir* -Recurse
try{
    #clear all working files
    Remove-Item $jsonFileDir* -Recurse
}
catch{
    #write-Error -MessageData 'Could not delete all files from file folder.
}

#$User = "Suman.kumar01@libertymutual.com"
#$PlainPassword = 'Sum@1234Kum@1234'

$User = "sagrspgrs-hrsync-admin@LibertyMutual.onmicrosoft.com"
$PlainPassword = 'T#xm4#rK'

$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $SecurePassword

 $dataJson = $jsonFileDir + $logFile + $timeRun + ".json"
 #$dataJson = $logFile + $timeRun + ".json"



 Import-Module .\Modules\AzureAD
 Import-Module .\Modules\PnP.PowerShell
     Connect-AzureAD -Credential $cred
      Write-Output "Azure AD Connection successful in Web Jobs"
       #$users = Get-AzureADUser -ObjectId Suman.kumar01@libertymutual.com |
       $users = Get-AzureADUser -All $true 
       <#|

        Select-Object @{N='IdName';E={if($_.UserPrincipalName){$_.UserPrincipalName}else{""}}},
        @{N='LMOfficeNumber';E={ if ($_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_extensionAttribute1"]) { $_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_extensionAttribute1"] } else { "" }}},
        @{N='DeptID';E={if($_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_extensionAttribute2"]){$_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_extensionAttribute2"]}else {""}}},
         @{N='WorkPhone';E={if($_.TelephoneNumber){$_.TelephoneNumber}else{""}}},
         @{N='CellPhone';E={if($_.Mobile){$_.Mobile}else{""}}},
         @{N='Mailstop';E={if($_.info){$_.info}else{""}}},
         @{N='Company';E={if($_.CompanyName){$_.CompanyName}else{""}}},
        @{N='Division';E={if($_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_division"]){$_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_division"]}else{""}}},
        @{N='LM-MarketName';E={if($_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LM_MarketName"]){$_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LM_MarketName"]}else{""}}},
        @{N='LM-MarketCode';E={if($_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LM_MarketCode"]){$_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LM_MarketCode"]}else{""}}},
        @{N='LM-SBUCode';E={if($_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LM_SBUCode"]){$_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LM_SBUCode"]}else{""}}},
        @{N='LMJobCode';E={if($_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LmJobCode"]){$_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LmJobCode"]}else{""}}},
        @{N='LMJobFamily';E={if($_.LMJobFamily){$_.LMJobFamily}else{""}}},
        @{N='LMJobFunction';E={if($_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LmJobFunction"]){$_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LmJobFunction"]}else{""}}},
        @{N='LMSegmentCode';E={if($_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LmSegmentCode"]){$_.ExtensionProperty["extension_128b6233d06d4df391d7de26c982b64e_LmSegmentCode"]}else{""}}},
        @{N='enterprisePIN';E={if($_.ExtensionProperty["employeeId"]){$_.ExtensionProperty["employeeId"]}else{""}}} 

        #$users | Export-Csv -Path $dataCsv -NoTypeInformation
        $json = $users 
        $json = ConvertTo-Json @($json)  -Compress
        $json ='{ "value":'+ $json + '}'
        #@($json) | Out-File $dataJson

        $Stream = [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($json))

        Connect-PnPOnline -Url https://libertymutual.sharepoint.com/sites/HRSync -Credential $cred  
	    Write-Output "SharePoint Connection successful in Web Jobs"
        $splib = "UserProfileData"  
        $data=Add-PnPFile  -Folder $splib -Stream $Stream -FileName $dataJson
 #>       

    #}
 
#&$env:64bitPowerShellPath -WindowStyle Hidden -NonInteractive -Command $Script 

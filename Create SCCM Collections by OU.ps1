<#	
	.NOTES
	===========================================================================
	 Created on:   	24th September 2017
	 Created by:   	Matthew Swanston
     Contact        http://thebeardysccmguy.blogspot.com
     Filename:      CollectionCreator	  	
	===========================================================================
	.DESCRIPTION
	Creates SCCM Collections for any enviroment based on Active Directory OUs with computers.
#>
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-location $SiteCode":"

$ErrorActionPreference= 'SilentlyContinue'
$Error1 = 0

#Refresh Schedule for collections
$Schedule = New-CMSchedule –RecurInterval Days –RecurCount 7
#$LimitingCollection = "All Systems"
#Create Collections
try{
   $array = @()
   $array = Import-Csv C:\temp\Collections.csv
   $CollectionsArray = @()

   foreach ($item in $array) {
                $CollectionsArray += New-Object -TypeName psobject -Property @{ 
                "Office" = $($item.office); 
                "Query" = "select * from SMS_R_System where SMS_R_System.SystemOUName like '$($item.ou)'";
                "OU" = $($item.ou);
                "Region" = $($item.region)
                }}			
    foreach ($item1 in $CollectionsArray) 
                {New-CMDeviceCollection -Name $item1.Office -Comment $item1.office -LimitingCollectionName $item1.region -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
	            Add-CMDeviceCollectionQueryMembershipRule -CollectionName $item1.Office -QueryExpression "select * from SMS_R_System where SMS_R_System.SystemOUName like '$($item1.ou)'" -RuleName $item1.Office
	            Write-host Collection $($item1.Office) created
				}
    }
catch{
$Error1 = 1
}
Finally{
    If ($Error1 -eq 1){
        Write-host -ForegroundColor Red "At least one collection already exists in SCCM."
        Pause
    }
    Else{
        Write-Host -ForegroundColor Green "Script execution completed sucessfully"
        Pause
        }
        }
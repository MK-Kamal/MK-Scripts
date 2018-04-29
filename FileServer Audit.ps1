$events =  Get-WinEvent  -LogName Security | where{$_.ID -eq "5145" -or $_.ID -eq "4663" -or $_.ID -eq "4656"}




$log = @()

foreach($event in $events){


$message = $event.Message -split "`n"


foreach($line in $message){



if($line -like "*File*" ){ $action = "1"}


if($line -like "*Account Name*" -and $line -notlike "*$*"){$username = $line }
if($line -like "*Object Name*" -and $line -notlike "*C:\Windows*" -and $line -notlike "*\Device*"){ $filename = $line }

if($line -like "*Accesses*"){$actiononfile  = $line }



}




$Properties = @{


"TimeGenerated" = $event.TimeCreated;

"MachineName" = $event.MachineName; 

"EventID" = $event.ID;

"file" = $filename;

"user" = $username;

"Action" = $actiononfile


}

if($action -eq "1"){


$Obj = New-Object -TypeName PSObject -Property $Properties



$log += $Obj

}

}



$log |Export-Csv c:\fillserveraudit.csv

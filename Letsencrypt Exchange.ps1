


if(!(Test-Path -Path "c:\cert")){mkdir -Path "c:\cert"}


#  Site Version  In Bottom Of Page https://www.powershellgallery.com/packages/ACMESharp/0.8.5.313


Import-Module ACMESharp

if($? -eq $false){


Install-Module -Name "ACMESharp"  -RequiredVersion 0.8.1  -Force -Scope AllUsers

#Install-Module -Name ACMESharp

Import-Module ACMESharp

}



## List the discovered Challenge Handler Providers Should Contain IIS


$handler = Get-ACMEChallengeHandlerProfile -ListChallengeHandlers

if ($handler -notcontains "iis"){

Install-Module ACMESharp.Providers.IIS
Enable-ACMEExtensionModule ACMESharp.Providers.IIS

}



if(!((Get-ACMERegistration).Contacts)){

New-ACMERegistration -Contacts mailto:mk_kamal@hotmail.com -AcceptTos


}

if(!((get-ACMEVault).id.Guid)){


Initialize-ACMEVault

}




# PowershellGet Package https://download.microsoft.com/download/C/4/1/C41378D4-7F41-4BBE-9D0D-0E4F98585C61/PackageManagement_x64.msi



######################################################
### getting Autodiscover and Outlookanywhere URL #####
######################################################

# Load Exchange Snapin
Add-PSSnapin *exchange*

# Get Outlook anywhere External URL And AutoDiscover Url

$outlookanywhere = (Get-OutlookAnywhere).InternalHostname.ToString()
$autodiscover = $outlookanywhere -replace  "^\w+","autodiscover"



# mail0417, auto0417, cert0417 and multiNameCert0417 should be automatically generated from the following combination

# mail + current month + current year, for example: mail0617
# auto + current month + current year, for example: auto0617
# cert + current month + current year, for example: cert0617
# multiNameCert + current month + current year, for example: multiNameCert0617


# define variable For Data And Alias Name 

$date = get-date -uformat "%Y%m%d%I%M%p"
$mailalias = "mail" + $date
$autoalias = "auto" + $date
$multialias = "multinamecert" + $date
$certalias = "cert" + $date

New-ACMEIdentifier -Dns $outlookanywhere -Alias $mailalias
New-ACMEIdentifier -Dns $autodiscover -Alias $autoalias
New-ACMEIdentifier -Dns $dns2 -Alias  $dns2alias

# The generated variables on the first step now should be automatically inserted in the rest of the script

Complete-ACMEChallenge $mailalias -ChallengeType http-01 -Handler iis -HandlerParameters @{ WebSiteRef = 'Default Web Site' }
Complete-ACMEChallenge $autoalias -ChallengeType http-01 -Handler iis -HandlerParameters @{ WebSiteRef = 'Default Web Site' }
Complete-ACMEChallenge $dns2alias -ChallengeType http-01 -Handler iis -HandlerParameters @{ WebSiteRef = 'Default Web Site' }

Submit-ACMEChallenge $mailalias -ChallengeType http-01 
Submit-ACMEChallenge $autoalias -ChallengeType http-01 
Submit-ACMEChallenge $dns2alias -ChallengeType http-01 

# Now the script should wait and keep check if the status marked as Valid.
do{
$checkmail = Update-ACMEIdentifier $mailalias
sleep 5


}until($checkmail.status -eq "valid" )

do{
$checkalias = Update-ACMEIdentifier $autoalias

}until($checkalias.status -eq "valid")



# Once the status is vaild and ONLY IF the status become vaild, should the script continue to the rest.

New-ACMECertificate $mailalias -Generate -AlternativeIdentifierRefs $autoalias  -Alias $multialias
Submit-ACMECertificate $multialias
Update-ACMECertificate $multialias

# Here the certification (cert0417.pfx) should be exported with the name from the variable as on the first step, for example (cert0617.pfx):

mkdir  -Path "C:\Temp\le"


Get-ACMECertificate $multialias -ExportPkcs12 "C:\Temp\le\$certalias.pfx" -CertificatePassword 'EX123!'


$password = ConvertTo-SecureString -String "EX123!" -Force –AsPlainText


#Get-ChildItem -Path "C:\Temp\le\$certalias.pfx" | Import-PfxCertificate  -CertStoreLocation cert:\localMachine\my –Exportable -Password $password

$importcert = Import-PfxCertificate -FilePath "C:\Temp\le\$certalias.pfx" -CertStoreLocation cert:\localMachine\my –Exportable -Password $password


# after the previous command, it should be Thumbprint output generated.
# In this last step the script should wait for this Thumbprint output, once we get the Thumbprint, the script should insert the generated Thumbprint automatically to finish the job.

sleep 30

#### need to check the result from this command


Enable-ExchangeCertificate -Thumbprint $importcert.Thumbprint -Services POP,IMAP,SMTP,IIS -DoNotRequireSsl -Confirm:$false -force






#########################
### Set FriendlyName  ###
#########################

$mypath = "Cert:\LocalMachine\my"
$thumb = $importcert.Thumbprint
Get-ChildItem "$mypath\$thumb" | foreach {$_.FriendlyName = $outlookanywhere}






#################################
##### Remove Expire certificate #
#################################

$mypath = "Cert:\LocalMachine\my"

$date = Get-Date

$mypathcert = Get-ChildItem $mypath

 foreach ($certificate in $mypathcert){

 $certsub = $certificate.Subject
 $expiredate = $certificate.NotAfter
 $thumb = $certificate.Thumbprint

 
 if($certificate.NotAfter -le $date){


Write-Host "Expire $certsub At $expiredate"

Remove-Item $mypath\$thumb

 }
 }


 ################################


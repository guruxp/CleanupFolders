param([string]$EmailAddress,[string]$Username,[switch]$Impersonate,[string]$EwsUrl,[string]$EWSManagedApiDLLFilePath,[string]$CSVFile,[switch]$NoCertValidation)

if ($NoCertValidation)
{
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}

if (!$EmailAddress)
{
    throw "EmailAddress is missing";
}
if (!$EmailAddress.Contains("@"))
{
    throw "EmailAddress not valid";
}

if (!$EWSManagedApiDLLFilePath)
{
    $EWSManagedApiDLLFilePath = "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"
}
if (!(Get-Item -Path $EWSManagedApiDLLFilePath -ErrorAction SilentlyContinue))
{
    throw "$($EWSManagedApiDLLFilePath) not found. You can download it from https://github.com/OfficeDev/ews-managed-api/blob/master/README.md";
}
else {
    Import-Module $EWSManagedApiDLLFilePath
}

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)

if ($Username)
{
    $Password = Read-Host "Password for $($Username): ";
    $service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password);   
    
} else {
    $service.UseDefaultCredentials = $true;
}

if ($EwsUrl)
{
    $service.URL = New-Object Uri($EwsUrl);
} else {
    Write-Host "EWSUrl is missing"
exit
}

$service.URL = New-Object Uri($EwsUrl);

if ($Impersonate)
{
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
}
$Folderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$Folderview.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.Webservices.Data.BasePropertySet]::FirstClassProperties)
$Folderview.Traversal = [Microsoft.Exchange.Webservices.Data.FolderTraversal]::Deep

$objExchange = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::msgFolderRoot)  ###Inbox

$fv = [Microsoft.Exchange.WebServices.Data.FolderView]100
$fv.Traversal = 'Deep'
$objExchange.Load()

$folders = $objExchange.FindFolders($fv)|select DisplayName,ID
$folders | Export-Csv $CSVFile -NoTypeInformation -Force

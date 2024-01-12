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
    throw "EWSApi not found at $($EWSManagedApiDLLFilePath). You can download it from https://github.com/OfficeDev/ews-managed-api/blob/master/README.md";
}

[void][Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll");
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

if ($Impersonate)
{
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress);
}

if ($CSVFile)
{
    $FoldersToPurge = Import-Csv $CSVFile
    ForEach ($folder in $FoldersToPurge) {
        $FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($Folder.Id);
        $DupFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $FolderId);
        $DupFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete);
    }

} else {
    Write-Host "CSV file not specified"
exit
}


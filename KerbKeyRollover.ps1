function New-RandomPassword {
    param(
        [Parameter()]
        [int]$MinimumPasswordLength = 10,
        [Parameter()]
        [int]$MaximumPasswordLength = 20,
        [Parameter()]
        [int]$NumberOfAlphaNumericCharacters = 5,
        [Parameter()]
        [switch]$ConvertToSecureString
    )
    
    Add-Type -AssemblyName 'System.Web'
    $length = Get-Random -Minimum $MinimumPasswordLength -Maximum $MaximumPasswordLength
    $password = [System.Web.Security.Membership]::GeneratePassword($length,$NumberOfAlphaNumericCharacters)
    if ($ConvertToSecureString.IsPresent) {
        ConvertTo-SecureString -String $password -AsPlainText -Force
    } else {
        $password
    }
}

Connect-AzAccount -Identity

$GAU = 'Global_admin@domain.com'
$DAU = 'Domain\Domain_Admin'
$GA = ConvertTo-SecureString( Get-AzKeyVaultSecret -VaultName 'Azure Vault Name' -Name Global_admin -AsPlainText) -AsPlainText -Force
$DA = ConvertTo-SecureString( Get-AzKeyVaultSecret -VaultName 'Azure Vault Name' -Name Domain_admin -AsPlainText) -AsPlainText -Force
$GAO = New-Object System.Management.Automation.PSCredential ($GAU, $DA)
$DAO = New-Object System.Management.Automation.PSCredential ($DAU, $DA)

Import-Module 'C:\Program Files\Microsoft Azure Active Directory Connect\AzureADSSO.psd1'

New-AzureADSSOAuthenticationContext -CloudCredentials $GAO

Update-AzureADSSOForest -OnPremCredentials $DAO

$password_generated = New-RandomPassword -MinimumPasswordLength 16 -ConvertToSecureString 
Set-ADAccountPassword -Identity Global_admin -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $password_generated -Force)
Set-AzKeyVaultSecret -VaultName 'Azure Vault Name' -Name Global_admin -SecretValue $password_generated
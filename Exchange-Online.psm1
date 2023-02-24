function Connect-Online {
    [CmdletBinding()]
    param (
        # The CertificateThumbprint parameter specifies the certificate that's used for CBA.
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string]
        $CertificateThumbprint,

        # The AppId parameter specifies the application ID of the service principal that's used in certificate based authentication (CBA).
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string]
        $ClientID,

        # The Organization parameter specifies the organization that's used in CBA.
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string]
        $Organization,

        # Parameter help description
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string]
        $TenantId
    )
    Connect-MgGraph -ClientID $ClientID -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
    Connect-ExchangeOnline -AppId $ClientID -CertificateThumbPrint $CertificateThumbprint -Organization $Organization

}

function Disconnect-Online {
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
}

function Get-GraphUser {
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Mandatory=$true)]
        [string]
        $UserPrincipalName
    )

    $Filter = 'UserPrincipalName eq ''{0}'' and mail eq ''{0}''' -f $UserPrincipalName
    $MgUser = Get-MgUser -Filter $Filter -ConsistencyLevel eventual -CountVariable UserCount

    return $UserCount
}

function Get-MigrationBatches {
    $MigrationBatches = Get-MigrationBatch

    return $MigrationBatches
}
function Move-Mailbox {
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Mandatory=$true)]
        [string]
        $EmailAddress,

        # Specifies an unique name for the migration batch.
        [Parameter(Mandatory=$true)]
        [string]
        $FullName,

        # The NotificationEmails parameter specifies one or more email addresses that migration status reports are sent to.
        [Parameter(Mandatory=$true)]
        [string]
        $NotificationEmails,

        # The TargetDeliveryDomain parameter specifies the FQDN of the external email address created in the source forest for
        # the mail-enabled user when the migration batch is complete.
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TargetDeliveryDomain
    )

    $CSVFile = New-TemporaryFile
    Add-Content -Path $CSVFile -Value 'EmailAddress'
    Add-Content -PassThru $CSVFile -Value $EmailAddress

    $MigrationEndpoint =  Get-MigrationEndpoint -Type 'ExchangeRemoteMove'
    $OnboardingBatch = New-MigrationBatch -Name $FullName -AutoComplete -Autostart -Confirm:$false -CSVData ([System.IO.File]::ReadAllBytes($CSVFile)) -NotificationEmails $NotificationEmails -SourceEndpoint $MigrationEndpoint.Identity -TargetDeliveryDomain $TargetDeliveryDomain -TimeZone $TimeZone
    $Message = 'Creating migrationbatch for {0}.' -f $OnboardingBatch.Identity
    Write-Log -Message $Message -Severity 'Debug'

    Remove-Item $CSVFile  
}

Export-ModuleMember Connect-Online, Disconnect-Online, Get-GraphUser, Get-MigrationBatches, Move-Mailbox
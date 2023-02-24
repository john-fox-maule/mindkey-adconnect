<#
    .SYNOPSIS
    SYNOPSIS

    .LINK
    Connect to Exchange servers using remote PowerShell https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell?view=exchange-ps
#>

function Connect-ExchangeOnPremise {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]
        $ServerFQDN
    )

    $ConnectionUri = 'http://{0}/PowerShell/' -f $ServerFQDN
    $Script:ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos
    Import-PSSession $Script:ExchangeSession -DisableNameChecking
    
}

function Disconnect-ExchangeOnPremise {
    Remove-PSSession $Script:ExchangeSession
}

function Enable-EmployeeMailbox {
    [CmdletBinding()]
    param (
        # The Identity parameter specifies the user or InetOrgPerson object that you want to mailbox-enable.
        [Parameter(Mandatory=$true)]
        [string]
        $UserIdParameter,

        # The Database parameter specifies the Exchange database that contains the new mailbox.
        [Parameter(Mandatory=$false)]
        [string]
        $DatabaseIdParameter,

        # The Language parameter specifies the language for the mailbox.
        [Parameter(Mandatory=$false)]
        [string]
        $Language = 'da-DK'
    )

    Enable-Mailbox -Identity $UserIdParameter -Database $DatabaseIdParameter
    Set-Mailbox -Identity $UserIdParameter -Languages $Language
    Set-MailboxRegionalConfiguration -Identity $UserIdParameter -Language $Language -DateFormat $null -TimeFormat $null -LocalizeDefaultFolderName
}

function Get-EmployeeMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Identity
    )

    # Is the mailbox OnPremise
    $EmployeeMailbox = Get-Mailbox -Identity $Identity

    return $EmployeeMailbox
}

Export-ModuleMember Connect-ExchangeOnPremise, Disconnect-ExchangeOnPremise, Enable-EmployeeMailbox, Get-EmployeeMailbox
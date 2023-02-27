<#
    .SYNOPSIS
    Updates Active Directory attributes from MindKey.

    .DESCRIPTION
    Updates Active Directory attributes from MindKey like name, phonenumber, office, department etc.

    .PARAMETER ClientCertificateThumbprint
    The Thumbprint of the client certificate you got from MindKey.

    .PARAMETER CustomerId
    The CustomerId you got from MindKey.

    .PARAMETER ReportTerminatedEmployees
    Send an email with employess that will be terminated within 30 days.

    .PARAMETER ServiceUrl
    BaseAddress of MindKey REST api.

    .EXAMPLE
    MindKeyADConnect -ClientCertificateThumbprint 'dee30d71ec3bef0ac4322f2eba4751f943abbd65' -CustomerId '3abedcd4-953c-40aa-a07b-ee251d3ec085' -ServiceUrl 'https://integration2.mindkey.com/'

    .INPUTS
    None. You cannot pipe objects to MindKeyADConnect.

    .OUTPUTS
    None.

    .NOTES
    Version:        1.0
    Author:         John Fox Maule
    Creation date:  2022/02/06

    .LINK
    MindKey Integrations API https://apidocs.mindkey.com/en-us/about-integrations-api.html

    .LINK
    PowerShell Reference ActiveDirectory https://docs.microsoft.com/en-us/powershell/module/activedirectory/get-aduser

    .LINK
    PowerShell Reference ActiveDirectory https://docs.microsoft.com/en-us/powershell/module/activedirectory/set-aduser

#>

#---------------------------------------------------------------------------
#
#  Parameters
#
#---------------------------------------------------------------------------

[CmdletBinding()]
param (
    # MindKey Certificate ThumbPrint
    [Parameter(Mandatory=$false)]
    [string]
    $ClientCertificateThumbprint,

    # MindKey CustomerId 
    [Parameter(Mandatory=$false)]
    [string]
    $CustomerId,

    # Parameter help description
    [Parameter(Mandatory=$false)]
    [switch]
    $HideTerminatedEmployees = $false,

    # Parameter help description
    [Parameter(Mandatory=$false)]
    [Switch]
    $ReportTerminatedEmployees = $false,

    # BaseAddress for MindKey REST
    [Parameter(Mandatory=$false)]
    [string]
    $ServiceUrl
)

#---------------------------------------------------------------------------
#
#  Imports
#
#---------------------------------------------------------------------------

Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement
Import-Module Microsoft.Graph.Users
Import-Module (Join-Path $PSScriptRoot -ChildPath 'Exchange-Online.psm1') -Force
Import-Module (Join-Path $PSScriptRoot -ChildPath 'Exchange-OnPremise.psm1') -Force
Import-Module (Join-Path $PSScriptRoot -ChildPath 'MindKey.psm1') -Force

#---------------------------------------------------------------------------
#
#  Constants
#
#---------------------------------------------------------------------------

$ApplicationName = 'MindKey Service'

$CharacterSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'.ToCharArray()

# By default, debug messages are not displayed in the console, but you can
# display them by setting $DebugPreference variable.
$DebugPreference = 'Continue'

# By default, information messages are not displayed in the console, but you can
# display them by setting $InformationPreference variable.
$InformationPreference = 'Continue'

#---------------------------------------------------------------------------
#
#  Functions
#
#---------------------------------------------------------------------------

function Get-LastRuntime {
    $RegistryPath = 'HKCU:\SOFTWARE\Bugfinder Software\MindKey AD Connect\1.0'
    $RegistryName = 'Ticks'

    if (-not (Test-Path -Path $RegistryPath)) {
        New-Item -Path $RegistryPath -Force
    } 
        
    if ($null -eq (Get-Item -Path $RegistryPath).GetValue($RegistryName)) {
        $Ticks = (Get-Date -AsUTC).Ticks
        New-ItemProperty -Path $RegistryPath -Name $RegistryName -PropertyType Qword -Value $Ticks

        return Get-Date -Date $Ticks
    }
    
    $Ticks = Get-ItemPropertyValue -Path $RegistryPath -Name $RegistryName
    Set-ItemProperty -Path $RegistryPath -Name $RegistryName -Value (Get-Date -AsUTC).Ticks

    return Get-Date -Date $Ticks
}
function Get-RandomPassword {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]
        $Length = 12
    )

    $Rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $Bytes = New-Object byte[]($Length)
    $Rng.GetBytes($Bytes)

    $Result = New-Object char[]($Length)

    for ($i = 0 ; $i -lt $Length ; $i++) {
        $Result[$i] = $CharacterSet[$Bytes[$i]%$CharacterSet.Length]
    }

    return (-join $Result)
}

function Find-ByUserPrincipalName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $UserPrincipalName,

        [Parameter(Mandatory=$false)]
        [string[]]
        $Properties = '*'
    )

    $ADUser = Get-ADUser -Filter {userPrincipalName -eq $UserPrincipalName} -Properties $Properties

    return $ADUser
}

function Find-ManagerById {
    [CmdletBinding()]
    param (
        # ID on the employee’s manager.
        [Parameter(Mandatory=$true)]
        [ValidateLength(1, 10)]
        [string]
        $ManagerId
    )

    $Manager = Get-MKEmployee -EmployeeId $ManagerId

    if ($Manager.Email) {
        $ManagerEmail = $Manager.Email
        $ADUser = Get-ADUser -Filter {userPrincipalName -eq $ManagerEmail} -Properties ObjectGUID, Manager

        # Return Manager DistinguishedName
        return $ADUser.DistinguishedName
    }
}

function Find-ManagerByOrganizationId {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [ValidateLength(1, 10)]
        [string]
        $OrganizationId
    )

    $Organization = Get-MKOrganization -OrganizationId $OrganizationId
    if ($Organization.ManagerPositionId) {
        $ManagerPosition = Get-MKPositionVersion -PositionId $Organization.ManagerPositionId

        if ($ManagerPosition.EmployeeID) {
            $Manager = Get-MKEmployee -EmployeeId $ManagerPosition.EmployeeID

            if ($Manager.Email) {
                $ManagerEmail = $Manager.Email
                $ADUser = Get-ADUser -Filter { userPrincipalName -eq $ManagerEmail } -Properties ObjectGUID, Manager

                # Return Manager DistinguishedName
                return $ADUser.DistinguishedName
            }
        }
    }
}

function Format-PhoneNumber {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]
        $PhoneNumber,

        # CountryCode
        [Parameter(Mandatory=$false)]
        [String]
        $CountryCode = '+45'
    )

    if ($PhoneNumber.Length -lt 8) {
        $Message = 'Phonenumber {0} is shorter than 8 characters' -f $PhoneNumber
        Write-Log -Message $Message -Severity 'Warning'
    } else {
        # Remove whitespaces and prefix
        $PhoneNumber = $PhoneNumber.Replace(' ', '');
        $PhoneNumber = $PhoneNumber.Replace($CountryCode, '');

        # Format it nicely
        $PhoneNumber = ('{0} {1} {2} {3} {4}' -f $CountryCode,
                                                 $PhoneNumber.Substring(0, 2),
                                                 $PhoneNumber.Substring(2, 2),
                                                 $PhoneNumber.Substring(4, 2),
                                                 $PhoneNumber.Substring(6, 2))
    }

    return $PhoneNumber
}

function Hide-TerminatedEmployees {
    $YesterDay = (Get-Date -AsUTC).AddDays(-1)
    $TerminatedEmployees = Get-MKEmployee -PositionMode 'TerminatedInnerJoin' | Where-Object { $_.PositionVersion_ValidTo -lt $YesterDay}

    if ($TerminatedEmployees.Count -gt 0) {
        $TerminatedEmployees | ForEach-Object {
            $Message = ('User {0} should be hidden from the Exchange Address list' -f $_.Name_FullName)
            Write-Log -Message $Message -Severity 'Information'
        }
    }
}

function New-ADUsersFromMindKey {
    Get-ADDomainController -Discover -Site $Script:Site -ForceDiscover
    $FutureEmployees = Get-MKEmployee -PositionMode 'FutureInnerJoin'

    if ($FutureEmployees.Count -gt 0) {

        # Create OnPremise emails first otherwise cmdlets conflicts with Exchange-Online.
        $FutureEmployees | ForEach-Object {
            $ADUser = Find-ByUserPrincipalName -UserPrincipalName $_.Email

            # User exists in Active Directory but does not have email.
            if ($ADUser -and $_.PreventAutoEnable -eq $false -and -not $ADUser.EmailAddress ) {
                $SamAccountNameSplitted = $_.Email -split '@'
                $SamAccountName = $SamAccountNameSplitted[0]

                $Message = ('User with EmployeeID {0} does not have a mailbox. Creating mailbox for {1}.' -f $_.EmployeeID, $_.Name_FullName)
                Write-Log -Message $Message -Severity 'Information'

                # Create Mailbox.
                Connect-ExchangeOnPremise -ServerFQDN $Script:ExchangeServer
                Enable-EmployeeMailbox -UserIdParameter $SamAccountName -DatabaseIdParameter $Script:MailboxDatabase
                Disconnect-ExchangeOnPremise
            }

        }

        Connect-Online -CertificateThumbprint $Script:CertificateThumbprint -ClientID $ClientID -Organization $Organization -TenantId $TenantId
        $FutureEmployees | ForEach-Object {
            $ADUser = Find-ByUserPrincipalName -UserPrincipalName $_.Email

            # User exists in Active Directory and has email.
            if ($ADUser -and $_.PreventAutoEnable -eq $false -and $ADUser.EmailAddress) {
                $Message = 'Checking if user {0} is migrated to Office 365.' -f $_.Name_FullName
                Write-Log -Message $Message -Severity 'Debug'

                $CountUser = Get-GraphUser -UserPrincipalName $ADUser.EmailAddress
                if ($CountUser -eq 1) {
                    $Message = 'User {0} is synced to AAD.' -f $_.Name_FullName
                    Write-Log -Message $Message -Severity 'Debug'

                    $MailUserFilter = 'UserPrincipalName -eq ''{0}''' -f $ADUser.EmailAddress
                    $MailUser = Get-MailUser -Filter $MailUserFilter

                    if ($MailUser.Count -eq 1) {
                        $Message = 'User {0} is OnPremise, checking If migration is running.' -f $MailUser
                        Write-Log -Message $Message -Severity 'Information'

                        $MigrationBatches = Get-MigrationBatches

                        if ($MigrationBatches.Count -eq 0) {
                            $Message = 'Creating MigrationBatch for user ''{0}''.' -f $MailUser
                            Write-Log -Message $Message -Severity 'Information'
                            Move-Mailbox -FullName $_.Name_FullName -EmailAddress $ADUser.EmailAddress -NotificationEmails $Script:NotificationEmails -TargetDeliveryDomain $TargetDeliveryDomain
                        }

                        $MigrationBatches | ForEach-Object {
                            if ($_.Identity -eq $MailUser) {
                                $Message = 'User ''{0}'' status is ''{1}''.' -f $MailUser, $_.Status
                                Write-Log -Message $Message -Severity 'Debug'
                            } else {
                                $Message = 'Creating MigrationBatch for user ''{0}''.' -f $_.Name_FullName
                                Write-Log -Message $Message -Severity 'Information'
                                Move-Mailbox -FullName $MailUser -EmailAddress $ADUser.EmailAddress -NotificationEmails $Script:NotificationEmails -TargetDeliveryDomain $TargetDeliveryDomain
                            }
                        }
                    }
                }
            }

            if (-not $ADUser -and $_.PreventAutoEnable -eq $false) {
                
                $OtherAttributes = @{}

                $AccountPassword = Get-RandomPassword | ConvertTo-SecureString -AsPlainText -Force

                # Get the location.
                if ($_.PositionVersion_LocationId) {
                    $Location = Get-MKLocation -LocationId $_.PositionVersion_LocationId
                }

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers General tab
                #
                #---------------------------------------------------------------------------

                $GivenName = $_.Name_FirstName
                $Surname = $_.Name_LastName
                $DisplayName = $_.Name_FullName
                $Description = ('MindKey AD ({0} on {1})' -f $_.CreatedBy, $_.CreatedDateTime)

                if ($Location) {
                    $OtherAttributes.Add('physicalDeliveryOfficeName', $Location.Name)
                }

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers Address tab
                #
                #---------------------------------------------------------------------------
                
                $Location = Get-MKLocation -LocationId $_.PositionVersion_LocationId
                if ($Location) {
                    $OtherAttributes.Add('StreetAddress', $Location.Address_Street)
                    $OtherAttributes.Add('l', $Location.Address_City)
                    $OtherAttributes.Add('PostalCode', $Location.Address_ZipPostalCode)
                }

                # We always want Denmark
                $OtherAttributes.Add('co', 'Denmark')
                $OtherAttributes.Add('c', 'DK')
                $OtherAttributes.Add('countryCode', '208')

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers Account tab
                #
                #---------------------------------------------------------------------------

                $UserPrincipalName = $_.Email
                $SamAccountNameSplitted = $_.Email -split '@'
                $SamAccountName = $SamAccountNameSplitted[0]

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers Attribute Editor tab
                #
                #---------------------------------------------------------------------------

                if ($_.EmployeeID) {
                    $OtherAttributes.Add('employeeID', $_.EmployeeID)
                }

                if ($_.Name_MiddleName) {
                    $OtherAttributes.Add('middleName', $_.Name_MiddleName)
                }

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers Organization tab
                #
                #---------------------------------------------------------------------------
                
                $Title = $_.PositionVersion_Title
                $Department = Get-MKOrganization -OrganizationId $_.PositionVersion_OrganizationId
                if ($Department) {
                    $OtherAttributes.Add('department', $Department.Name)
                }

                # Manager
                $ManagerPosition = Get-MKPositionVersion -PositionId $_.PositionVersion_ReportsTo
                if ($ManagerPosition.EmployeeID) {
                    $Manager = Get-MKEmployee -EmployeeId $ManagerPosition.EmployeeID
                    $ReportsToDescription = ('{0} {1}' -f $Manager.Title, $Manager.Name_FullName)
                    $ManagerDistinguishedName  = Find-ManagerById -ManagerId $ManagerPosition.EmployeeID
                    if ($ManagerDistinguishedName) {
                        $OtherAttributes.Add('Manager', $ManagerDistinguishedName)
                    }
                } else {
                    $ReportsToDescription = ''
                }

                #---------------------------------------------------------------------------
                #
                #
                #
                #---------------------------------------------------------------------------

                $Template = 'MM/dd/yyyy HH:mm:ss'
                $Culture = [CultureInfo]'da-DK'
                $ValidFrom = [DateTime]::ParseExact($_.PositionVersion_ValidFrom, $Template, $null).ToString('d MMMM, yyyy', $Culture)
                
                $Message = ('{0} starts as {1} on {2} at the office {3} in ''{4}''' -f $DisplayName, $Title, $ValidFrom, $Location.Name, $Department.Name)
                Write-Log -Message $Message -Severity 'Information'
               
                $BodyTable = ('<table>
                                 
                                <caption>User account autocreated in Active Directory</caption>

                                 <tr>
                                   <td><b>Start date:</b></td>
                                   <td>{0}</td>
                                 </tr>

                                 <tr>
                                   <td><b>First name:</b></td>
                                   <td>{1}</td>
                                 </tr>

                                 <tr>
                                   <td><b>Last name:</b></td>
                                   <td>{2}</td>
                                 </tr>
 
                                 <tr>
                                   <td><b>Display name:</b></td>
                                   <td>{3}</td>
                                 </tr>

                                 <tr>
                                   <td><b>Office:</b></td>
                                   <td>{4}</td>
                                 </tr>

                                 <tr>
                                   <td><b>E-mail:</b></td>
                                   <td>{5}</td>
                                 </tr>

                                 <tr>
                                   <td><b>Job Title:</b></td>
                                   <td>{6}</td>
                                 </tr>

                                 <tr>
                                   <td><b>Department:</b></td>
                                   <td>{7}</td>
                                 </tr>

                                 <tr>
                                   <td><b>Manager:</b></td>
                                   <td>{8}</td>
                                 </tr>

                               </table>' -f $ValidFrom, $GivenName, $_.Name_LastName, $DisplayName, $Location.Name, $_.Email, $Title, $Department.Name, $ReportsToDescription)
                $Body = ('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>a</title></head><body>{0}</body></html>' -f $BodyTable)
                $Subject = ('INFO: Staff - {0}' -f $Location.Name)
                Send-UserMail -Content $Body -Subject $Subject
                                
                # Log stuff.
                $Message = ('User with EmployeeID {0} does not exist in Active Directory. Creating user {1}.' -f $_.EmployeeID, $_.Name_FullName)
                Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
                Write-Log -Message $Message -Severity 'Information'
                Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'

                # Create the user in Active Directory.
                New-ADUser -DisplayName $DisplayName `
                           -Enabled $true `
                           -AccountPassword $AccountPassword `
                           -GivenName $GivenName `
                           -Surname $Surname `
                           -Name $_.Name_FullName `
                           -OtherAttributes $OtherAttributes `
                           -UserPrincipalName $UserPrincipalName `
                           -SamAccountName $SamAccountName `
                           -Description $Description `
                           -Title $Title `
                           -Path $Script:OUPath
            }
        }
        Disconnect-Online
    }
    
}

function Out-EmployeeEquipment {

    $Employees = Get-MKEmployee # | Select-Object EmployeeID, Name_FullName, Location_Name, OrganizationName | Export-csv -Path C:\temp\emp.csv -Encoding utf8
    return $Employees
}
function Out-TerminatedEmployees {

    $Template = 'MM/dd/yyyy HH:mm:ss'
    $Culture = [CultureInfo]'da-DK'

    $FutureEmployees = Get-MKEmployee -PositionMode 'FutureInnerJoin'

    $FutureTable = '<h1>New employees</h1>'
    $FutureTable += '<table border="1">'
    $FutureTable += '<tr bgcolor="#2ff3e0">'
    $FutureTable += '<th style="text-align:left">Job Title</th>'
    $FutureTable += '<th style="text-align:left">Display name</th>'
    $FutureTable += '<th style="text-align:left">Office</th>'
    $FutureTable += '<th style="text-align:left">Start date</th>'
    $FutureTable += '<th style="text-align:left">Manager</th>'
    $FutureTable += '</tr>'

    $FutureEmployees | ForEach-Object {
        $ValidFrom  = [DateTime]::ParseExact($_.PositionVersion_ValidFrom, $Template, $null).ToString('d. MMMM, yyyy', $Culture)
        $Location = Get-MKLocation -LocationId $_.PositionVersion_LocationId

        # Manager
        $ManagerPosition = Get-MKPositionVersion -PositionId $_.PositionVersion_ReportsTo
        $ManagerFullName = ''
        if ($ManagerPosition.Count -gt 0) {
            $Manager = Get-MKEmployee -EmployeeId $ManagerPosition[0].EmployeeID
            $ManagerFullName = $Manager.Name_FullName
        }

        $FutureTable += '<tr>'
        $FutureTable += ('<td>{0}</td>' -f $_.PositionVersion_Title)
        $FutureTable += ('<td>{0}</td>' -f $_.Name_FullName)
        $FutureTable += ('<td>{0}</td>' -f $Location.Name)
        $FutureTable += ('<td>{0}</td>' -f $ValidFrom)
        $FutureTable += ('<td>{0}</td>' -f $ManagerFullName)
        $FutureTable += '</tr>'
    }
    $FutureTable += '</table>'

    if ($FutureEmployees.Count -eq 0) {
        $FutureTable = ''
    }

    $DaysFromNow = (Get-Date -AsUTC).AddDays(30)
    $TerminatedEmployees = Get-MKEmployee -PositionMode 'TerminatedInnerJoin' |
        Sort-Object -Property PositionVersion_ValidTo, Name_FullName |
        Where-Object { $_.PositionVersion_ValidTo -lt $DaysFromNow}

    $BodyTable = '<h1>Resignations</h1>'
    $BodyTable += '<table border="1">'
    $BodyTable += '<tr bgcolor="#f8d210">'
    $BodyTable += '<th style="text-align:left">Job Title</th>'
    $BodyTable += '<th style="text-align:left">Display name</th>'
    $BodyTable += '<th style="text-align:left">Office</th>'
    $BodyTable += '<th style="text-align:left">End date</th>'
    $BodyTable += '<th style="text-align:left">Reason</th>'
    $BodyTable += '</tr>'

    $TerminatedEmployees | ForEach-Object {
        $ValidTo = [DateTime]::ParseExact($_.PositionVersion_ValidTo, $Template, $null).ToString('d. MMMM, yyyy', $Culture)

        $BodyTable += '<tr>'
        $BodyTable += ('<td>{0}</td>' -f $_.Title)
        $BodyTable += ('<td>{0}</td>' -f $_.Name_FullName)
        $BodyTable += ('<td>{0}</td>' -f $_.Location_Name)
        $BodyTable += ('<td>{0}</td>' -f $ValidTo)
        $BodyTable += ('<td>{0}</td>' -f $_.PositionVersion_TerminateReasonCodeId)
        $BodyTable += '</tr>'
    }

    $BodyTable += '</table>'

    $PositionsTable = '<h1>Position changes</h1>'
    $PositionsTable += '<table border="1">'
    $PositionsTable += '<tr bgcolor="#fa26a0">'
    $PositionsTable += '<th style="text-align:left">Job Title</th>'
    $PositionsTable += '<th style="text-align:left">Valid from</th>'
    $PositionsTable += '<th style="text-align:left">Valid to</th>'
    $PositionsTable += '<th style="text-align:left">Display name</th>'
    $PositionsTable += '<th style="text-align:left">Office</th>'
    $PositionsTable += '<th style="text-align:left">Created DateTime</th>'
    $PositionsTable += '<th style="text-align:left">Created by</th>'
    $PositionsTable += '<th style="text-align:left">Modified DateTime</th>'
    $PositionsTable += '<th style="text-align:left">Modified by</th>'
    $PositionsTable += '</tr>'

    $LastWeek = (Get-Date -AsUTC).AddDays(-7)
    $Positions = Get-MKPositionVersion -FromDate $LastWeek
    $Positions | ForEach-Object {
        if ($_.EmployeeID) {
            $Employee = Get-MKEmployee -EmployeeId $_.EmployeeID
            
            if ($Employee) {
                if ($null -ne $_.ValidFrom) {
                    $ValidFrom = [DateTime]::ParseExact($_.ValidFrom, $Template, $null).ToString('d. MMMM, yyyy', $Culture)
                } else {
                    $ValidFrom = ''
                }

                if ($null -ne $_.ValidTo) {
                    $ValidTo = [DateTime]::ParseExact($_.ValidTo, $Template, $null).ToString('d. MMMM, yyyy', $Culture)
                } else {
                    $ValidTo = ''
                }

                if ($Employee.Location_Name) {
                    $Location_Name = $Employee.Location_Name
                } else {
                    $Location_Name = ''
                }
                $CreatedDateTime = [DateTime]::ParseExact($_.CreatedDateTime, $Template, $null).ToString('d. MMMM, yyyy, HH:mm:ss', $Culture)
                $ModifiedDateTime = [DateTime]::ParseExact($_.ModifiedDateTime, $Template, $null).ToString('d. MMMM, yyyy, HH:mm:ss', $Culture)
                
                $PositionsTable += '<tr>'
                $PositionsTable += ('<td>{0}</td>' -f $_.Title)
                $PositionsTable += ('<td>{0}</td>' -f $ValidFrom)
                $PositionsTable += ('<td>{0}</td>' -f $ValidTo)
                $PositionsTable += ('<td>{0}</td>' -f $Employee.Name_FullName)
                $PositionsTable += ('<td>{0}</td>' -f $Location_Name)
                $PositionsTable += ('<td>{0}</td>' -f $CreatedDateTime)
                $PositionsTable += ('<td>{0}</td>' -f $_.CreatedBy)
                $PositionsTable += ('<td>{0}</td>' -f $ModifiedDateTime)
                $PositionsTable += ('<td>{0}</td>' -f $_.ModifiedBy)
                $PositionsTable += '</tr>'
            }
        }
    }
    $PositionsTable += '</table>'

    if ($Positions.count -eq 0) {
        $PositionsTable = ''
    }

    $Body = ('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>a</title></head><body>{0}{1}{2}</body></html>' -f $FutureTable, $BodyTable, $PositionsTable)
    $Subject = 'INFO: Staff - new employees and resignations'
    Send-UserMail -Content $Body -Subject $Subject
}

function Send-UserMail {
    [CmdletBinding()]
    param (
        # Parameter help description
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Content,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Subject
    )

    $ReplyToArray = @()
    $ReplyTo | ForEach-Object {
        $ReplyToArray += @{
            EmailAddress = @{
                Address = $_.Address
                Name = $_.Name
            }
        }
    }

    $ToRecipientsArray = @()
    $ToRecipients | ForEach-Object {
        $ToRecipientsArray += @{
            EmailAddress = @{
                Address = $_.Address
                Name = $_.Name
            }
        }
    }
    
    $Params = @{
        Message = @{
            Subject = $Subject
            Body = @{
                ContentType = 'HTML'
                Content = $Content
            }
            From = @{
                EmailAddress = @{
                    Address = $Script:From.Address
                    Name = $Script:From.Name
                }
            }
            ReplyTo = $ReplyToArray
            Sender = @{
                EmailAddress = @{
                    Address = $Script:Sender.Address
                    Name = $Script:Sender.Name
                }

            }
            ToRecipients = $ToRecipientsArray
        }
        SaveToSentItems = 'false'
    }
    Send-MgUserMail -UserId $Script:Sender.Address -BodyParameter $Params
}

function Sync-MindKeyLocation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [datetime]
        $FromDate
    )

    if ($FromDate) {
        $LocationCollection = Get-MKLocation -FromDate $FromDate
    } else {
        $LocationCollection = Get-MKLocation
    }

    if ($LocationCollection.Count -gt 0) {
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Write-Log -Message '--  There are modfied Locations.' -Severity 'Information'
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        $LocationCollection | ForEach-Object {
            $Message = 'LocationId: {0} has changed, Name={1}, Street={2}, City={3}, Zip={4}' -f $_.LocationId, $_.Name, $_.Address_Street, $_.Address_City, $_.Address_ZipPostalCode
            Write-Log -Message $Message -Severity 'Information'

            $EmployeeCollection = Get-MKEmployee -LocationId $_.LocationId
            $EmployeeCollection | ForEach-Object {
                Sync-MindKeyEmployee -EmployeeId $_.EmployeeID
            }
        }
    }
}
function Sync-MindKeyEmployee {
    [CmdletBinding()]
    param (
        # MindKey EmployeeId
        [Parameter(Mandatory=$false)]
        [string]
        $EmployeeId,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [datetime]
        $CustomDate,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [datetime]
        $FromDate
    )

    # Fetch employees
    if ($EmployeeId) {
        $Employees = Get-MKEmployee -EmployeeId $EmployeeId
    } elseif ($FromDate) {
        $Employees = Get-MKEmployee -FromDate $FromDate
    } elseif ($CustomDate) {
        $Employees = Get-MKEmployee -CustomDate $CustomDate
    } else { 
        $Employees = Get-MKEmployee
    }

    if ($Employees.Count -gt 0) {
        $Message = ('--  Processing {0} employee(s)' -f $Employees.Count)
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Write-Log -Message '--' -Severity 'Information'
        Write-Log -Message $Message -Severity 'Information'
        Write-Log -Message '--' -Severity 'Information'
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
    }

    $Employees | ForEach-Object {
        $Message = ('--  Processing {0}, EmployeeID {1}' -f $_.Name_FullName, $_.EmployeeID)
        Write-Log -Message $Message -Severity 'Information'
        $Message = ('--  CreatedBy: {0}, CreatedDateTime {1}' -f $_.CreatedBy, $_.CreatedDateTime)
        Write-Log -Message $Message -Severity 'Information'
        $Message = ('--  ModifiedBy {0}, ModifiedDateTime {1}' -f $_.ModifiedBy, $_.ModifiedDateTime)
        Write-Log -Message $Message -Severity 'Information'
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'

        $UserPrincipalName = $_.Email

        # Does the user have an email.
        if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            $Message = ('User with EmployeeID {0} does not have an email address in MindKey can not process employee.' -f $_.EmployeeID)
            Write-Log -Message $Message -Severity 'Warning'
            Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        } else {
            $EmailSplit = $UserPrincipalName -split '@'
            $Domain = $EmailSplit[$EmailSplit.Count - 1]  
            $HasCompanyMail = $Domain -eq $Script:CompanyDomain          
        }

        if ($false -eq $HasCompanyMail) {
            $Message = ('{0} does not have an @{1} email address in MindKey can not process employee.' -f $_.Name_FullName, $Script:CompanyDomain)
            Write-Log -Message $Message -Severity 'Warning'
            Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        }

        if ($UserPrincipalName -and $HasCompanyMail) {
            $ADUser = Find-ByUserPrincipalName -UserPrincipalName $UserPrincipalName

            # Can the user be found in Active Directory.
            if ($ADUser) {
                $ObjectGUID = $ADUser['ObjectGUID'].Value
                $UserAttributes = @{}
                
                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers Attribute Editor tab
                #
                #---------------------------------------------------------------------------

                $UserAttributes.Add('EmployeeID', $_.EmployeeID)

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers General tab
                #
                #---------------------------------------------------------------------------

                #$UserAttributes.Add('GivenName', $_.Name_FirstName)
                #$Surname = ('{0} {1}' -f $_.Name_MiddleName, $_.Name_LastName).TrimStart()
                #$UserAttributes.Add('Surname', $Surname)
                #$UserAttributes.Add('DisplayName', $_.Name_FullName)
                $UserAttributes.Add('physicalDeliveryOfficeName', $_.Location_Name)
                # If the work phone number is filled out in MindKey sync it.
                if ([string]::IsNullOrWhiteSpace($_.WorkPhoneNumber_LocalNumber) -eq $false) {
                    $WorkPhoneNumber = Format-PhoneNumber -PhoneNumber $_.WorkPhoneNumber_LocalNumber
                    $UserAttributes.Add('telephoneNumber', $WorkPhoneNumber)
                }

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers Address tab
                #
                #---------------------------------------------------------------------------

                $LocationId = $_.LocationId
                if ($LocationId) {
                    $Location = Get-MKLocation -LocationId $LocationId
                    if ($Location) {
                        $UserAttributes.Add('StreetAddress', $Location.Address_Street)
                        $UserAttributes.Add('l', $Location.Address_City)
                        $UserAttributes.Add('PostalCode', $Location.Address_ZipPostalCode)
                        $UserAttributes.Add('co', 'Denmark')
                        $UserAttributes.Add('c', 'DK')
                        $UserAttributes.Add('countryCode', '208')
                    }
                }

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers Account tab
                #
                #---------------------------------------------------------------------------

                $AccountExpirationDate = $ADUser.AccountExpirationDate
                $PreventAutoDisable = $_.$PreventAutoDisable
                if ($_.ValidTo -and (-not $PreventAutoDisable)) {
                    $Message = 'ValidTo = {0}' -f $_.ValidTo
                    Write-Log $Message -Severity 'Debug'
                    [datetime]$ValidToPlusOneDay = $_.ValidTo
                    $ValidToPlusOneDay = $ValidToPlusOneDay.AddDays(1)

                    # Account expires set to end of 28 February 2022 will become 1. march 2022 so add 1 day
                    if ($ValidToPlusOneDay -ne $AccountExpirationDate) {
                        $Message = 'Setting ''Account expires to End of'' {0}' -f $_.ValidTo.ToLongDateString()
                        Write-Log -Message $Message -Severity 'Information'
                        Set-ADAccountExpiration -Identity $ObjectGUID -DateTime $ValidToPlusOneDay -WhatIf
                    }
                } else {
                    # Has ValidTo been set before but is now cleared
                    if ($AccountExpirationDate) {
                        # Has this been set manually then ALERT
                        $Message = 'Clearing ''Account expires from End of'' {0}' -f $AccountExpirationDate.ToLongDateString()
                        Clear-ADAccountExpiration -Identity $ObjectGUID -WhatIf
                    }
                }

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers Telephones tab
                #
                #---------------------------------------------------------------------------

                # If the mobile phone number is filled out in MindKey sync it
                if ([string]::IsNullOrWhiteSpace($_.MobilePhoneNumber_LocalNumber) -eq $false) {
                    $MobilePhoneNumber = Format-PhoneNumber -PhoneNumber $_.MobilePhoneNumber_LocalNumber
                    $UserAttributes.Add('mobile', $MobilePhoneNumber)
                }

                #---------------------------------------------------------------------------
                #
                #  Active Directory Users and Computers Organization tab
                #
                #---------------------------------------------------------------------------
                
                $UserAttributes.Add('Title', $_.Title)
                $UserAttributes.Add('Department', $_.OrganizationName)
                $UserAttributes.Add('Company', $_.Dimension2_Description)
                $ManagerId = $_.ManagerId
                if ($ManagerId) {
                    $ManagerDistinguishedName = Find-ManagerById -ManagerId $ManagerId
                    $UserAttributes.Add('Manager', $ManagerDistinguishedName)
                }

                # Setup hashtables for attributes we want to remove, add and replace.
                $RemoveHashTable = @{}
                $AddHashTable = @{}
                $ReplaceHashTable = @{}

                # Loop through the properties and compare them.
                foreach ($Key in $UserAttributes.Keys) {
                    $MindKeyProperty = $UserAttributes[$Key]
                    $ADProperty = $ADUser[$Key].Value

                    # Does the property needs to be updated in anyway.
                    if ($MindKeyProperty -ne $ADProperty) {
                        
                        # Process the difference
                        if ([string]::IsNullOrWhiteSpace($MindKeyProperty) -and [string]::IsNullOrWhiteSpace($ADProperty)) {
                            # Do nothing it could be empty in one place and have a whitespace in the other
                        } elseif ($null -eq $ADProperty) {
                            $Message = ('Adding [''{0}'', ''{1}''] to AddHashTable because ''{0}'' is null' -f $Key, $MindKeyProperty)
                            Write-Log -Message $Message -Severity 'Debug'
                            $AddHashTable.Add($Key, $MindKeyProperty)
                        } elseif ($null -eq $MindKeyProperty) {
                            $Message = ('Adding [''{0}'', ''{1}''] to RemoveHashTable because MindKey field is null or whitespace' -f $Key, $ADProperty)
                            Write-Log -Message $Message -Severity 'Debug'
                            $RemoveHashTable.Add($Key, $ADProperty)
                        } else {
                            $Message = ('Adding [''{0}'', ''{1}''] to ReplaceHashTable because ''{1}'' is not equal to ''{2}''' -f $Key, $MindKeyProperty, $ADProperty)
                            Write-Log -Message $Message -Severity 'Debug'
                            $ReplaceHashTable.Add($Key, $MindKeyProperty)
                        }
                    }
                    else {
                        $Message = ('Property [''{0}'', ''{1}''] are in sync between Active Directory and MindKey.' -f $Key, $ADProperty)
                        Write-Log -Message $Message -Severity 'Debug'
                    }
                }

                # Order is Remove, Add, Replace and Clear
                $Message = ('User with ObjectGUID: {0} have attributes that needs to be' -f $ObjectGUID)
                if ($RemoveHashTable.Count -gt 0) {
                    $MessageRemoved = ('{0} {1}' -f $Message, 'removed')
                    Write-Log -Message $MessageRemoved -Severity 'Debug'
                    Set-ADUser -Identity $ObjectGUID -Remove $RemoveHashTable
                }

                if ($AddHashTable.Count -gt 0) {
                    $MessageAdded = ('{0} {1}' -f $Message, 'added')
                    Write-Log -Message $MessageAdded -Severity 'Debug'
                    Set-ADUser -Identity $ObjectGUID -Add $AddHashTable
                }

                if ($ReplaceHashTable.Count -gt 0) {
                    $MessageReplaced = ('{0} {1}' -f $Message, 'replaced')
                    Write-Log -Message $MessageReplaced -Severity 'Debug'
                    Set-ADUser -Identity $ObjectGUID -Replace $ReplaceHashTable
                }
            }
            Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        }
    }

}

function Sync-MindKeyManager {
    [CmdletBinding()]
    param (
        #
        [Parameter(Mandatory=$false)]
        [datetime]
        $FromDate
    )

    if ($FromDate) {
        $OrganizationCollection = Get-MKOrganization -FromDate $FromDate
    } else {
        $OrganizationCollection = Get-MKOrganization
    }

    if ($OrganizationCollection.Count -gt 0) {
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Write-Log -Message '--  There are modfied Organizations.' -Severity 'Information'
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
    }

    $OrganizationCollection | ForEach-Object {
        if (-not $_.ManagerPositionId) {
            Write-Log -Message 'Contact HR department missing manager on department' -Severity 'Warning'
        } else {
            $ManagerPosition = Get-MKPositionVersion -PositionId $_.ManagerPositionId
            if ($ManagerPosition.Count -eq 1 -and $ManagerPosition.EmployeeID) {
                $Manager = Find-ManagerById -ManagerId $ManagerPosition.EmployeeId
            } elseif ($ManagerPosition.Count -gt 1) {
                $Now = Get-Date -AsUTC
                foreach ($Position in $ManagerPosition) {
                    if ($Position.EmployeeID) {
                        if (($Now -ge $Position.ValidFrom) -and ($Now -le $Position.ValidTo)) {
                            $Manager = Find-ManagerById -ManagerId $Position.EmployeeID
                        } elseif (($Now -ge $Position.ValidFrom) -and ($null -eq $Position.ValidTo)) {
                            $Manager = Find-ManagerById -ManagerId $Position.EmployeeID
                        }
                    }
                }
            }
            $Message = '--  Processing department {0} {1}: {2}' -f $_.OrganizationId, $_.Name, $Manager
            Write-Log -Message $Message -Severity 'Debug'
            $Message = ('--  CreatedBy: {0}, CreatedDateTime {1}' -f $_.CreatedBy, $_.CreatedDateTime)
            Write-Log -Message $Message -Severity 'Debug'
            $Message = ('--  ModifiedBy {0}, ModifiedDateTime {1}' -f $_.ModifiedBy, $_.ModifiedDateTime)
            Write-Log -Message $Message -Severity 'Debug'

            $EmployeeCollection = Get-MKEmployee -OrganizationId $_.OrganizationId
            $EmployeeCollection | ForEach-Object {
                Sync-MindKeyEmployee -EmployeeId $_.EmployeeID
            }
        }

    }
}

function Sync-MindKeyPosition {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [datetime]
        $FromDate,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [string]
        $EmployeeID
    )

    if ($EmployeeID) {
        $PositionsCollection = Get-MKPositionVersion -EmployeeId $EmployeeID
    } else {
        $PositionsCollection = Get-MKPositionVersion -FromDate $FromDate
    }

    if ($PositionsCollection.Count -gt 0) {
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Write-Log -Message '--  There are modfied positions, syncing the positions to AD.'
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
    }

    $Now = Get-Date -AsUTC
    $PositionsCollection | ForEach-Object {
        if (($Now -ge $_.ValidFrom) -and ($Now -le $_.ValidTo)) {
            Sync-MindKeyEmployee -EmployeeId $_.EmployeeID
        } elseif (($Now -ge $_.ValidFrom) -and ($null -eq $_.ValidTo)) {
            Sync-MindKeyEmployee -EmployeeId $_.EmployeeID
        }
        
    }
}

function Import-MindKeySettings {
    [CmdletBinding()]
    param (
        # The name of the appsettings json file.
        [Parameter(Mandatory=$false)]
        [string]
        $FileName = 'appsettings.json'
    )

    $AppSettingsPath = Join-Path -Path $PSScriptRoot -ChildPath $FileName
    $AppSettings = Get-Content -Path $AppSettingsPath | ConvertFrom-Json

    if (-not $ClientCertificateThumbprint) {
        $Script:ClientCertificateThumbprint = $AppSettings.MindKey.ClientCertificateThumbprint
    }

    if (-not $CustomerId) {
        $Script:CustomerId = $AppSettings.MindKey.CustomerId
    }

    if (-not $ServiceUrl) {
        $Script:ServiceUrl = $AppSettings.MindKey.ServiceUrl
    }

    # AAD section
    $Script:ClientId = $AppSettings.AAD.ClientID
    $Script:Organization = $AppSettings.AAD.Organization
    $Script:TenantId = $AppSettings.AAD.TenantId
    $Script:CertificateThumbprint = $AppSettings.AAD.CertificateThumbprint

    # ActiveDirectory section
    $Script:OUPath = $AppSettings.ActiveDirectory.OUPath
    $Script:Site = $AppSettings.ActiveDirectory.Site

    # Exchange section
    $Script:ExchangeServer = $AppSettings.Exchange.ExchangeServer
    $Script:MailboxDatabase = $AppSettings.Exchange.MailboxDatabase

    # Logging section
    $Script:LoggingPath = $AppSettings.Logging.Path

    # Mail 
    $Script:CompanyDomain = $AppSettings.Mail.CompanyDomain
    $Script:From = $AppSettings.Mail.From
    $Script:ReplyTo = $AppSettings.Mail.ReplyTo
    $Script:Sender = $AppSettings.Mail.Sender
    $Script:To = $AppSettings.Mail.To
    $Script:ToRecipients = $AppSettings.Mail.ToRecipients

    # Migration section
    $Script:NotificationEmails = $AppSettings.Migration.NotificationEmails
    $Script:TargetDeliveryDomain = $AppSettings.Migration.TargetDeliveryDomain

    # MindKey Section
    $Script:SearchConditionEmail = $AppSettings.MindKey.SearchConditionEmail
    $Script:TimeZone = $AppSettings.MindKey.TimeZone
}

#---------------------------------------------------------------------------
#
#  Main
#
#---------------------------------------------------------------------------

Import-MindKeySettings

# Initialize important module variables.
Set-LoggingPath -LoggingPath $Script:LoggingPath
Set-SearchConditionEmail -SearchConditionEmail $Script:SearchConditionEmail

Connect-MindKeyAccount -BaseAddress $ServiceUrl -CertificateThumbprint $ClientCertificateThumbprint -CustomerId $CustomerId -ApplicationName $ApplicationName

if ($ReportTerminatedEmployees.IsPresent) {
    Connect-MgGraph -ClientID $ClientID -TenantId $TenantId -CertificateThumbprint $Script:CertificateThumbprint
    Out-TerminatedEmployees
    Disconnect-MgGraph
} else {
    # Create new users
    New-ADUsersFromMindKey

    $LastRuntime = Get-LastRuntime
    #$LastRunTime = Get-Date -Year 2023 -Month 2 -Day 23 -Hour 12 -Minute 33 -AsUTC
    if ($null -ne $LastRuntime) {
        Write-Log -Message $LastRuntime -Severity 'Information'

        #---------------------------------------------------------------------------
        #
        #  Synchronize change locations
        #
        #---------------------------------------------------------------------------
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Write-Log -Message '--  Sync-MindKeyLocation' -Severity 'Information'
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Sync-MindKeyLocation -FromDate $LastRuntime

        #---------------------------------------------------------------------------
        #
        #  Synchronize managers
        #
        #---------------------------------------------------------------------------
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Write-Log -Message '--  Sync-MindKeyManager' -Severity 'Information'
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Sync-MindKeyManager -FromDate $LastRuntime

        #---------------------------------------------------------------------------
        #
        #  Synchronize changed positions
        #
        #---------------------------------------------------------------------------
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Write-Log -Message '--  Sync-MindKeyPosition' -Severity 'Information'
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Sync-MindKeyPosition -FromDate $LastRuntime

        #---------------------------------------------------------------------------
        #
        #  Synchronize employees
        #
        #---------------------------------------------------------------------------        
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Write-Log -Message '--  Employee-Sync' -Severity 'Information'
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
        Sync-MindKeyEmployee -FromDate $LastRuntime
        Write-Log -Message '------------------------------------------------------------------------------------------------' -Severity 'Information'
    } else {
        Write-Log -Message '-- ALERT WHAT THE FOX --' -Severity 'Error'
        #Sync-MindKeyEmployee
    }

}
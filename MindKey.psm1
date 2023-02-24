function Connect-MindKeyAccount {

    [CmdletBinding()]
    param (
        # BaseUrl of the MindKey REST api
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullorEmpty()]
        [string]
        $BaseAddress,

        # Specifies the digital public key certificate (X509) of a user account that has permission to send the request.
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string]
        $CertificateThumbprint,

        # The MindKey CustomerId
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string]
        $CustomerId,

        # The name of the application
        [Parameter(Mandatory=$false)]
        [string]
        $ApplicationName = 'MindKey Service',

        # The timezone
        [Parameter(Mandatory=$false)]
        [string]
        $TimeZone = 'Romance Standard Time',

        # The Systemuser.
        [Parameter(Mandatory=$false)]
        [string]
        $SystemUser = 'MindKey Service',

        # The culture.
        [Parameter(Mandatory=$false)]
        [string]
        $Culture = 'en-US',

        # The language.
        [Parameter(Mandatory=$false)]
        [string]
        $Language = 'en-US'
    )

    $ContentType = 'application/json; charset=utf-8'

    $Script:Headers = @{}
    $Script:Headers.Add('ApplicationName', $ApplicationName)
    $Script:Headers.Add('TimeZone', $TimeZone)
    $Script:Headers.Add('SystemUser', $SystemUser)
    $Script:Headers.Add('Culture', $Culture)
    $Script:Headers.Add('Language', $Language)

    # Specifies the client certificate that's used for a secure web request
    $Certificate = $null
    $Response = $null

    if ($CertificateThumbprint -and $CustomerId) {

        if ($CertificateThumbprint) {
            $Certificate = Get-ChildItem -Path Cert:\CurrentUser\My\ | Where-Object {$_.Thumbprint -eq $CertificateThumbprint}
        }

        if ($CustomerId) {
            $Script:BaseUri = ('{0}/api/{1}' -f $ServiceUrl, $CustomerId)
        } else {
            $Script:BaseUri = ('{0}/api' -f $ServiceUrl)
        }

        $LoginUri = ('{0}/system/login?ApplicationName={1}' -f $Script:BaseUri, $ApplicationName)

        try {
            $Response = Invoke-WebRequest -Uri $LoginUri -Method Post -Headers $Script:Headers -ContentType $ContentType -Certificate $Certificate -SessionVariable 'Session' -SslProtocol Tls12
            $Script:MindKeySession = $Session
            $StatusCode = $Response.StatusCode
        } catch {
            $StatusCode = $_.Exception.Response.StatusCode.value__
            Write-Log -Message "Failed.... $StatusCode" -Severity 'Error'
        }
    }
}

function Get-MKEmployee {
    [CmdletBinding()]
    param (
        # The employee’s id. The field is used as key in MindKey.
        [Parameter(Mandatory=$false)]
        [ValidateLength(1, 10)]
        [string]
        $EmployeeId,

        # The employee’s first name, middle name and surname combined.
        [Parameter(Mandatory=$false)]
        [ValidateLength(3, 150)]
        [string]
        $FullName,

        # ID on the employee’s current unit.
        [Parameter(Mandatory=$false)]
        [string]
        $OrganizationId,

        # ID of location.
        [Parameter(Mandatory=$false)]
        [string]
        $LocationId,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [datetime]
        $FromDate,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [datetime]
        $CustomDate,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [ValidateSet('CurrentOuterJoin', 'CurrentInnerJoin', 'TerminatedInnerJoin', 'FutureInnerJoin')]
        [string]
        $PositionMode = 'CurrentOuterJoin'
    )

    if ($EmployeeId) {
        # Find by the employee’s id.'
        $SearchCondition = @( @{ column = 'EmployeeId'; value = $EmployeeId } )
    } elseif ($FullName) {
        # Find by the fullname.
        $SearchCondition = @( @{ column = 'Name_FullName'; value = $FullName } )
    } elseif ($OrganizationId) {
        # Find by OrganizationId.
        $SearchCondition = @( @{ column = 'OrganizationId'; value = $OrganizationId } )
    } elseif ($LocationId) {
        # Find by the LocationId.
        $SearchCondition = @(
            @{ column = 'LocationId'; value = $LocationId; },
            @{ column = 'Email'; value = $Script:SearchConditionEmail; conditionOperator = 'LIKE'; logicalOperator = 'AND' } 
        )
    } else {
        # Find only company emails.
        $SearchCondition = @( @{ column = 'Email'; value = $Script:SearchConditionEmail; conditionOperator = 'LIKE' } )
    }

    $PostParameters = @{ 'columns' = $EmployeeColumns }
    $PostParameters.Add('searchCondition', $SearchCondition)

    if ($FromDate) {
        $CustomSearchCondition = @( @{ condition = 'FromDate'; value = $FromDate } )
        $PostParameters.Add('customSearchCondition', $CustomSearchCondition)
    }

    if ($CustomDate) {
        $CustomSearchCondition = @( @{ condition = 'FromDate'; value = $CustomDate } )
        $PostParameters.Add('customSearchCondition', $CustomSearchCondition)
    }

    if ($PositionMode) {
        $PostParameters.Add('positionMode', $PositionMode)
        $PostParameters.Add('positionColumns', $PositionVersionColumns)
        
        # This mode returns employees, where a positionVersion record exists.
        if ($PositionMode -eq 'CurrentInnerJoin') {
            $TransDate = Get-Date -AsUTC
            $CustomSearchCondition = @( @{ condition = "TransDate"; value = $TransDate; onlyParameter = $true } )
        }
        
        # This mode returns terminated employees since a given date.
        if ($PositionMode -eq 'TerminatedInnerJoin') {
            $TransDate = (Get-Date -AsUTC).AddDays(-1)
            $CustomSearchCondition = @( @{ condition = "TransDate"; value = $TransDate; onlyParameter = $true } )
            $PostParameters.Add('customSearchCondition', $CustomSearchCondition)
        }
        
        # This mode returns employees with a future employment from a given date.
        if ($PositionMode -eq 'FutureInnerJoin') {
            $TransDate = (Get-Date -AsUTC).AddDays(-1)
            #$SearchCondition = @( @{ column = 'Email'; value = $Script:SearchConditionEmail; conditionOperator = 'LIKE' } )
            #$PostParameters.Add('searchCondition', $SearchCondition)
            $CustomSearchCondition = @(
                @{ condition = "TransDate"; value = $TransDate; onlyParameter = $true }
            )
            $PostParameters.Add('customSearchCondition', $CustomSearchCondition)
        }

    }

    $Uri = 'employees/find'
    $Employees = Invoke-MindKeyRequest -Uri $Uri -Method 'Post' -InputObject $PostParameters

    return $Employees
}

function Get-MKEmployeeEquipment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [ValidateLength(1, 10)]
        [string]
        $EmployeeId,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [ValidateLength(1, 20)]
        [string]
        $EquipmentName,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [ValidateLength(1, 50)]
        [string]
        $Description,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [datetime]
        $IssuedDate
    )

    if ($EmployeeId) {
        # Find by the employee’s id.'
        $SearchCondition = @( @{ column = 'EmployeeId'; value = $EmployeeId } )
    } elseif ($EquipmentName) {
        # Find by the equipment name.
        $SearchCondition = @( @{ column = 'EquipmentName'; value = $EquipmentName } )
    } elseif ($Description) {
        # Find by the equipment description.
        $SearchCondition = @( @{ column = 'Description'; value = $Description } )
    } elseif ($IssuedDate) {
        # Find by the equipment issued date.
        $SearchCondition = @( @{ column = 'IssuedDate'; value = $IssuedDate } )
    }

    $PostParameters = @{ 'columns' = $EmployeeEquipmentColumns }
    $PostParameters.Add('searchCondition', $SearchCondition)

    $lol = ConvertTo-Json -InputObject $PostParameters

    $Uri = 'employeeequipments/find'
    $EmployeesEquipment = Invoke-MindKeyRequest -Uri $Uri -Method 'Post' -InputObject $PostParameters

    return $EmployeesEquipment
}

function Get-MKLocation {
    [CmdletBinding()]
    param (
        # ID of location.
        [Parameter(Mandatory=$false)]
        [ValidateLength(1, 10)]
        [string]
        $LocationId,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [datetime]
        $FromDate
    )

    $PostParameters = @{ 'columns' = $LocationColumns }
    # Find by LocationId
    if ($LocationId) {
        $SearchCondition = @( @{ column = 'LocationId'; value = $LocationId } )
    }

    if ($FromDate) {
        $CustomSearchCondition = @( @{ condition = 'FromDate'; value = $FromDate } )
        $PostParameters.Add('customSearchCondition', $CustomSearchCondition)
    }

    if ($SearchCondition) {
        $PostParameters.Add('searchCondition', $SearchCondition)
    }

    $Uri = 'locations/find'
    $LocationCollection = Invoke-MindKeyRequest -Uri $Uri -Method 'Post' -InputObject $PostParameters

    return $LocationCollection
}

function Get-MKOrganization {
    [CmdletBinding()]
    param (
        # Unit ID.
        [Parameter(Mandatory=$false)]
        [ValidateLength(1, 10)]
        [string]
        $OrganizationId,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [datetime]
        $FromDate
    )

    $PostParameters = @{ 'columns' = $OrganizationColumns }

    # Find by OrganizationId.
    if ($OrganizationId) {
        $SearchCondition = @( @{ column = 'OrganizationId'; value = $OrganizationId } )
    }

    if ($FromDate) {
        $CustomSearchCondition = @( @{ condition = 'FromDate'; value = $FromDate } )
        $PostParameters.Add('customSearchCondition', $CustomSearchCondition)
    }

    if ($SearchCondition) {
        $PostParameters.Add('searchCondition', $SearchCondition)
    }

    $Uri = 'organizations/find'
    $OrganizationCollection = Invoke-MindKeyRequest -Uri $Uri -Method 'Post' -InputObject $PostParameters

    return $OrganizationCollection
}

function Get-MKPositionVersion {
    [CmdletBinding()]
    param (
        # ID on position i MindKey.
        [Parameter(Mandatory=$false)]
        [string]
        $PositionId,

        # The employee’s ID.
        [Parameter(Mandatory=$false)]
        [string]
        $EmployeeId,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [datetime]
        $FromDate
    )

    $PostParameters = @{ 'columns' = $PositionVersionColumns }

    # Find by PositionId
    if ($PositionId) {
        $SearchCondition = @( @{ column = 'PositionId'; value = $PositionId } )
    }

    # Find by EmployeeId
    if ($EmployeeId) {
        $SearchCondition = @( @{ column = 'EmployeeId'; value = $EmployeeId } )
    }

    if ($FromDate) {
        $CustomSearchCondition = @( @{ condition = 'FromDate'; value = $FromDate } )
        $PostParameters.Add('customSearchCondition', $CustomSearchCondition)
    }

    if ($SearchCondition) {
        $PostParameters.Add('searchCondition', $SearchCondition)
    }

    $Uri = 'positionversions/find'
    $PositionVersionCollection = Invoke-MindKeyRequest -Uri $Uri -Method 'Post' -InputObject $PostParameters

    return $PositionVersionCollection
}

function Invoke-MindKeyRequest {
    [CmdletBinding()]
    param (
        # The Uri
        [Parameter(Mandatory=$true, Position=0)]
        [string]
        $Uri,

        # Method 
        [Parameter(Mandatory=$true)]
        [ValidateSet("Delete", "Get", "Post", "Put")]
        [string]
        $Method,

        # An optional InputObject
        [Parameter(Mandatory=$false)]
        $InputObject
    )

    if ($null -eq $Script:MindKeySession) {
 -Message "MindKeySession not established" -Severity 'Error'
    }

    if ($InputObject -and $Method -eq 'Get') {
        $Content = $InputObject
    } elseif ($InputObject) {
        $Content = ConvertTo-Json -InputObject $InputObject
    } else {
        $Content = $null
    }

    $ContentType = 'application/json; charset=utf-8'

    try {
        $Uri = ("$script:BaseUri/{0}" -f $Uri)
        Write-Log "Calling $Uri" -Severity 'Debug'
        $Reponse = Invoke-WebRequest -Uri $Uri -Method $Method -Headers $Script:Headers -WebSession $Script:MindKeySession -ContentType $ContentType -Body $Content -SslProtocol Tls12
        $ResponseContent = $Reponse.Content | ConvertFrom-Json
        $StatusCode = $Reponse.StatusCode
    } catch {
        $StatusCode = $_.Exception.Response.StatusCode.value__
        Write-Log -Message "HTTP Status Code: $StatusCode" -Severity 'Error'
    }

    return $ResponseContent
}

function Set-LoggingPath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string]
        $LoggingPath
    )

    $Script:LoggingPath = $LoggingPath
}

function Set-SearchConditionEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullorEmpty()]
        [string]
        $SearchConditionEmail
    )

    $Script:SearchConditionEmail = $SearchConditionEmail
}

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $Message,

        # Parameter help description
        [Parameter(Mandatory=$false)]
        [ValidateSet('Error', 'Warning', 'Information', 'Debug')]
        [string]
        $Severity = 'Information'
    )

    $LogDateTime = Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'
    $LogDate = Get-Date -Format 'yyyyMMdd'
    $LogFile = ('MindKey-{0}.log' -f $LogDate)
    $LogFile = Join-Path -Path $Script:LoggingPath -ChildPath $LogFile

    $Message = ('{0} {1} {2}' -f $LogDateTime, $Severity.PadRight(11), $Message)
    Add-Content -Path $LogFile -Value $Message

    switch ($Severity) {
        'Error' { Write-Error -Message $Message }
        'Warning' { Write-Warning -Message $Message}
        'Debug' { Write-Debug -Message $Message }
        default { Write-Information -MessageData $Message}
    }
}

$EmployeeColumns = "EmployeeId,
                    Name_FirstName,
                    Name_MiddleName,
                    Name_LastName,
                    Name_FullName,
                    Email,
                    WorkPhoneNumber_LocalNumber,
                    MobilePhoneNumber_LocalNumber,
                    PreventAutoEnable,
                    PreventAutoDisable,
                    LocationId,
                    Location_Name,
                    Office,
                    ManagerId,
                    OrganizationId,
                    OrganizationName,
                    Title,
                    ActualValidTo,
                    ValidTo,
                    Dimension2_Description,
                    CreatedBy,
                    CreatedDateTime,
                    ModifiedBy,
                    ModifiedDateTime"

$EmployeeEquipmentColumns = "EmployeeId,
                    EquipmentName,
                    Description,
                    IssuedDate,
                    ReturnDate,
                    Qty,
                    QtyValue,
                    Note,
                    RegistrationDate,
                    CreatedBy,
                    CreatedDateTime,
                    ModifiedBy,
                    ModifiedDateTime"

$LocationColumns = "LocationId,
                    Name,
                    Address_Street,
                    Address_City,
                    Address_ZipPostalCode,
                    Address_CountryRegion,
                    CreatedBy,
                    CreatedDateTime,
                    ModifiedBy,
                    ModifiedDateTime"

$OrganizationColumns = "OrganizationId,
                        Name,
                        ManagerPositionId,
                        CreatedBy,
                        CreatedDateTime,
                        ModifiedBy,
                        ModifiedDateTime"

$PositionVersionColumns = "PositionId,
                           Title,
                           ReportsTo,
                           EmploymentCategoryName,
                           TerminateReasonCodeId,
                           LocationId,
                           OrganizationId,
                           EmployeeId,
                           ValidFrom,
                           ValidTo,
                           Dimension2Id,
                           CreatedBy,
                           CreatedDateTime,
                           ModifiedBy,
                           ModifiedDateTime"

Export-ModuleMember Connect-MindKeyAccount, Get-MKEmployee, Get-MKEmployeeEquipment, Get-MKLocation, Get-MKOrganization, Get-MKPositionVersion, Write-Log
Export-ModuleMember -Function Set-LoggingPath
Export-ModuleMember -Function Set-SearchConditionEmail
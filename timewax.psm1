$APIUri = 'https://api.timewax.com/'
$Token = $null
$ValidUntil = $null
$DatabaseConnection = $null
$ResourceTablePresent = $false
$TimeEntryTablePresent = $false
$ProjectTablePresent = $false
$ProjectBreakDownTablePresent = $false
$InvoiceLineTablePresent = $false
$CalendarEntryTablePresent = $false

function Get-TimeWaxToken {
    [cmdletbinding(DefaultParameterSetName='New')]
    param (
        [Parameter(Mandatory, ParameterSetName='New')]
        [pscredential]
        [System.Management.Automation.CredentialAttribute()] $Credential,

        [Parameter(Mandatory, ParameterSetName='New')]
        [ValidateNotNullOrEmpty()]
        [String] $ClientName,

        [Parameter(ParameterSetName='Current')]
        [Switch] $Current
    )
    process {
        if ($Current) {
            return [pscustomobject]@{
                Token = $script:Token
                ValidUntil = $script:ValidUntil
            }
        } else {
            $TokenUri = $script:APIUri + 'authentication/token/get/'
            $Body = [xml]('<request> 
                <client>{0}</client> 
                <username>{1}</username>
                <password>{2}</password> 
            </request>' -f $ClientName,$Credential.UserName,$Credential.GetNetworkCredential().Password)
            try {
                $Response = (Invoke-RestMethod -Uri $TokenUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
            } catch {
                Write-Error -ErrorRecord $_ -ErrorAction Stop
            }
            if ($Response.valid -eq 'no') {
                Write-Error -Message "Did not acquire a token. Exception: $($Response.errors.'#cdata-section')" -ErrorAction Stop
            } else {
                Set-Variable -Scope 1 -Name Token -Value $Response.token
                Set-Variable -Scope 1 -Name ValidUntil -Value (ConvertDateTime $Response.validUntil)
                Get-TimeWaxToken -Current
            }
        }
    }
}

#object should be usable on pipeline to map ResourceCode to filters for other Get functions
function Get-TimeWaxResource {
    [cmdletbinding(DefaultParameterSetName='Named')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName='Named')]
        [Alias('email','fullname','ResourceCode')]
        [ValidateNotNullOrEmpty()]
        [String] $Name,

        [Parameter(ParameterSetName='List')]
        [Switch] $List
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        if ($List) {
            $ResourceUri = $script:APIUri + 'resource/list/'
        } else {
            $ResourceUri = $script:APIUri + 'resource/get/'
        }
    } process {
        if ($List) {
            $Body = [xml]('<request> 
             <token>{0}</token> 
            </request>' -f $script:Token)
        } else {
            $Body = [xml]('<request> 
             <token>{0}</token> 
             <resource>{1}</resource> 
            </request>' -f $script:Token,$Name)
        }
        $Response = (Invoke-RestMethod -Uri $ResourceUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            if ($List) {
                foreach ($r in $Response.resources.resource) {
                    $outputObj = $r | ConvertXMLElement
                    $outputObj.PSObject.TypeNames.Insert(0,'TimeWax.Resource')
                    Write-Output -InputObject $outputObj
                }
            } else {
                $outputObj = $Response.resource | ConvertXMLElement
                $outputObj.PSObject.TypeNames.Insert(0,'TimeWax.Resource')
                Write-Output -InputObject $outputObj
            }
        }
    }
}

function Get-TimeWaxTimeEntry {
    [CmdletBinding(DefaultParameterSetName='List')]
    param (
        [Parameter(Mandatory, ParameterSetName='Resource', ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('Resource','Code')]
        [String] $ResourceCode,

        [Parameter(Mandatory, ParameterSetName='Project')]
        [ValidateNotNullOrEmpty()]
        [String] $Project,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [datetime] $From = [datetime]::Now.AddMonths(-1),

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [datetime] $To = [datetime]::Now,

        [Switch] $onlyApprovedEntries
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        $TimeURI = $script:APIUri + 'time/entries/list/'
        $dateFrom = $From.ToString('yyyyMMdd')
        $dateTo = $To.ToString('yyyyMMdd')
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token>
         <dateFrom>{1}</dateFrom>
         <dateTo>{2}</dateTo>
        </request>' -f $script:Token, $dateFrom, $dateTo)

        if ($ResourceCode) {
            $child = $Body.CreateElement("resource")
            $child.InnerText = $ResourceCode
            [void] $body.DocumentElement.AppendChild($child)
        }
        if ($Project) {
            $child = $Body.CreateElement("project")
            $child.InnerText = $Project
            [void] $body.DocumentElement.AppendChild($child)
        }
        if ($onlyApprovedEntries) {
            $child = $Body.CreateElement("onlyApprovedEntries")
            $child.InnerText = 'yes'
            [void] $body.DocumentElement.AppendChild($child)
        }
        $Response = (Invoke-RestMethod -Uri $TimeURI -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            foreach ($e in $Response.Entries.entry) {
                $outputObj = $e | ConvertXMLElement
                $outputObj.PSObject.TypeNames.Insert(0,'TimeWax.TimeEntry')
                Write-Output -InputObject $outputObj
            }
        }
    }
}

function Get-TimeWaxProject {
    [CmdletBinding(DefaultParameterSetName='List')]
    param (
        [Parameter(ParameterSetName='Named')]
        [ValidateNotNullOrEmpty()]
        [Alias('Code','name')]
        [String] $Project
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        if ($PSCmdlet.ParameterSetName -eq 'List') {
            $ProjectUri = $script:APIUri + 'project/list/'
        } else {
            $ProjectUri = $script:APIUri + 'project/get/'
        }
    } process {
        $Body = [xml]('<request> 
            <token>{0}</token> 
        </request>' -f $script:Token)

        if ($Project) {
            $child = $Body.CreateElement("project")
            $child.InnerText = $Project
            [void] $body.DocumentElement.AppendChild($child)
        }

        $Response = (Invoke-RestMethod -Uri $ProjectUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            if ($PSCmdlet.ParameterSetName -eq 'List') {
                foreach ($P in $Response.projects.project) {
                    $outputObj = $P | ConvertXMLElement
                    $outputObj.PSObject.TypeNames.Insert(0,'TimeWax.Project')
                    Write-Output -InputObject $outputObj
                }
            } else {
                $outputObj = $Response.project | ConvertXMLElement
                $outputObj.PSObject.TypeNames.Insert(0,'TimeWax.Project')
                Write-Output -InputObject $outputObj
            }
        }
    }
}

function Get-TimeWaxProjectBreakdown {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('Code')]
        [String] $Project
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        $BreakdownUri = $script:APIUri + 'project/breakdown/list/'
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token>
         <project>{1}</project>
        </request>' -f $script:Token,$Project)

        $Response = (Invoke-RestMethod -Uri $BreakdownUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            foreach ($B in $Response.breakdowns.breakdown) {
                $outputObj = $B | ConvertXMLElement
                $outputObj.PSObject.TypeNames.Insert(0,'TimeWax.ProjectBreakDown')
                Write-Output -InputObject $outputObj
            }
        }
    }
}

function Get-TimeWaxInvoiceLine {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [datetime] $From = [datetime]::Now.AddMonths(-1),

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [datetime] $To = [datetime]::Now
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        $InvoiceUri = $script:APIUri + 'invoices/lines/list/'
        $dateFrom = $From.ToString('yyyyMMdd')
        $dateTo = $To.ToString('yyyyMMdd')
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token>
         <dateFrom>{1}</dateFrom>
         <dateTo>{2}</dateTo>
        </request>' -f $script:Token, $dateFrom, $dateTo)

        $Response = (Invoke-RestMethod -Uri $InvoiceUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            foreach ($I in $Response.invoices.invoice) {
                $outputObj = $I | ConvertXMLElement
                $outputObj.PSObject.TypeNames.Insert(0,'TimeWax.InvoiceLine')
                Write-Output -InputObject $outputObj
            }
        }
    }
}

function Get-TimeWaxBudgetCost {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('Code','Name')]
        [String] $Project
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        $BudgetUri = $script:APIUri + 'budgetcost/list/'
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token>
         <project>{1}</project>
        </request>' -f $script:Token, $Project)

        $Response = (Invoke-RestMethod -Uri $BudgetUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            foreach ($B in $Response.budgetcosts.budgetcost) {
                $B | ConvertXMLElement
            }
        }
    }
}

function Get-TimeWaxPurchaseInvoice {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('Code','Name')]
        [String] $Project
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        $PurchaseInvoiceUri = $script:APIUri + 'purchaseinvoices/list/'
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token>
         <project>{1}</project>
        </request>' -f $script:Token, $Project)

        $Response = (Invoke-RestMethod -Uri $PurchaseInvoiceUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            foreach ($P in $Response.purchaseInvoices.purchaseInvoice) {
                $P | ConvertXMLElement
            }
        }
    }
}

function Get-TimeWaxCalculation {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [datetime] $From = [datetime]::Now.AddMonths(-1),

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [datetime] $To = [datetime]::Now
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        $CalculationUri = $script:APIUri + 'calculation/list/'
        $dateFrom = $From.ToString('yyyyMMdd')
        $dateTo = $To.ToString('yyyyMMdd')
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token>
         <dateFrom>{1}</dateFrom>
         <dateTo>{2}</dateTo>
        </request>' -f $script:Token, $dateFrom, $dateTo)

        $Response = (Invoke-RestMethod -Uri $CalculationUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            foreach ($C in $Response.calculations.calculation) {
                $C | ConvertXMLElement
            }
        }
    }
}

function Get-TimeWaxCalendarEntry {
    [CmdletBinding(DefaultParameterSetName='List')]
    param (
        [Parameter(Mandatory, ParameterSetName='Resource', ValueFromPipelineByPropertyName)]
        [Alias('Resource','Code')]
        [ValidateNotNullOrEmpty()]
        [String] $ResourceCode,

        [Parameter(Mandatory, ParameterSetName='Id')]
        [ValidateNotNullOrEmpty()]
        [String] $Id,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [datetime] $From = [datetime]::Now.AddMonths(-1),

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [datetime] $To = [datetime]::Now
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        $CalendarUri = $script:APIUri + 'calendar/entries/list/'
        $dateFrom = $From.ToString('yyyyMMdd')
        $dateTo = $To.ToString('yyyyMMdd')
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token>
         <dateFrom>{1}</dateFrom>
         <dateTo>{2}</dateTo>
        </request>' -f $script:Token, $dateFrom, $dateTo)

        if ($ResourceCode) {
            $child = $Body.CreateElement("resource")
            $child.InnerText = $ResourceCode
            [void] $body.DocumentElement.AppendChild($child)
        }
        if ($Id) { #API apparently does not filter on this https://support.timewax.com/hc/nl/articles/203496036-calendar-entries-list
            $child = $Body.CreateElement("id")
            $child.InnerText = $Id
            [void] $body.DocumentElement.AppendChild($child)
        }
        Write-Verbose -Message "$($Body | fc | Out-String)"
        $Response = (Invoke-RestMethod -Uri $CalendarUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            foreach ($e in $Response.Entries.entry) {
                $outputObj = $e | ConvertXMLElement
                $outputObj.PSObject.TypeNames.Insert(0,'TimeWax.CalendarEntry')
                Write-Output -InputObject $outputObj
            }
        }
    }
}

function Get-TimeWaxCompany {
    [CmdletBinding(DefaultParameterSetName='List')]
    param (
        [Parameter(ParameterSetName='Named')]
        [ValidateNotNullOrEmpty()]
        [String] $Name,

        [Parameter(ParameterSetName='Code')]
        [ValidateNotNullOrEmpty()]
        [String] $Code
    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        if ($PSCmdlet.ParameterSetName -eq 'List') {
            $CompanyUri = $script:APIUri + 'company/list/'
        } else {
            $CompanyUri = $script:APIUri + 'company/get/'
        }
        
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token>
        </request>' -f $script:Token)
        if ($Name) {
            $child = $Body.CreateElement("company")
            $child.InnerText = $Name
            [void] $body.DocumentElement.AppendChild($child)
        }
        if ($Code) {
            $child = $Body.CreateElement("company")
            $child.InnerText = $Code
            [void] $body.DocumentElement.AppendChild($child)
        }
        $Response = (Invoke-RestMethod -Uri $CompanyUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } 
        if ($PSCmdlet.ParameterSetName -eq 'List') {
            foreach ($C in $Response.companies.company) {
                $C | ConvertXMLElement
            }
        } else {
            $Response.company | ConvertXMLElement
        }
    }
}

function Connect-TimeWaxDatabase {
    [CmdletBinding(DefaultParameterSetName='New')]
    param (
        [Parameter(Mandatory,ParameterSetName='New')]
        [PSCredential]
        [System.Management.Automation.CredentialAttribute()] $Credential,

        [Parameter(Mandatory,ParameterSetName='New')]
        [ValidateNotNullOrEmpty()]
        [String] $Server,

        [Parameter(Mandatory,ParameterSetName='New')]
        [ValidateNotNullOrEmpty()]
        [String] $Database,

        [Parameter(ParameterSetName='Current')]
        [Switch] $Current
    )
    process {
        if ($Current) {
            return $script:DatabaseConnection
        } else {
            if ($script:DatabaseConnection.State -eq 'Open') {
                Write-Warning -Message "Only one open database connection is supported at a time. Currently connected to: $($script:DatabaseConnection.DataSource)"
            }
            $Connection = New-Object -TypeName System.Data.SQLClient.SQLConnection
            $Connection.ConnectionString = "server='$Server';database='$Database';User Id=$($Credential.UserName); Password=$($Credential.GetNetworkCredential().Password);"
            try {
                $Connection.Open()
            } catch {
                Write-Error -ErrorRecord $_ -ErrorAction Stop
            }
            Set-Variable -Name DatabaseConnection -Value $Connection -Scope 1
            Connect-TimeWaxDatabase -Current
        }
    }
}

function Disconnect-TimeWaxDatabase {
    if ($script:DatabaseConnection) {
        $script:DatabaseConnection.Close()
        $script:DatabaseConnection.Dispose()
        Set-Variable -Name DatabaseConnection -Value $null -Scope 1
    }
}

function Write-TimeWaxDatabase {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [PSCustomObject] $InputObject
    )
    begin {
        if (-not $script:DatabaseConnection.State -eq 'Open') {
            Write-Error -Message 'A database connection does not exist, run Connect-TimeWaxDatabase first' -ErrorAction Stop
        }
    } process {
        if ($InputObject.pstypenames.Contains('TimeWax.Resource')) {
            $TableName = 'Resource'
            $TableVar = $TableName + "TablePresent"
        } elseif ($InputObject.pstypenames.Contains('TimeWax.TimeEntry')) {
            $TableName = 'TimeEntry'
            $TableVar = $TableName + "TablePresent"
        } elseif ($InputObject.pstypenames.Contains('TimeWax.Project')) {
            $TableName = 'Project'
            $TableVar = $TableName + "TablePresent"
        } elseif ($InputObject.pstypenames.Contains('TimeWax.ProjectBreakDown')) {
            $TableName = 'ProjectBreakDown'
            $TableVar = $TableName + "TablePresent"
        } elseif ($InputObject.pstypenames.Contains('TimeWax.InvoiceLine')) {
            $TableName = 'InvoiceLine'
            $TableVar = $TableName + "TablePresent"
        } elseif ($InputObject.pstypenames.Contains('TimeWax.CalendarEntry')) {
            $TableName = 'CalendarEntry'
            $TableVar = $TableName + "TablePresent"
        } else {
            Write-Error -Message 'Unknown Inputtype, abort' -ErrorAction Stop
        }

        if (-not (Get-Variable -Name $TableVar -Scope 1).value) {
            if (TestTableExists -TableName $TableName) {
                Write-Verbose -Message "$TableName Table already exists"
                Set-Variable -Name $TableVar -Value $true -Scope 1
            } else {
                Write-Verbose -Message "$TableName Table does not exist, creating"
                $InputObject | CreateTableFromObject -TableName $TableName | InvokeSQLQuery
                Set-Variable -Name $TableVar -Value $true -Scope 1
            }
        }
        $InputObject | CreateInsertFromObject -TableName $TableName | InvokeSQLQuery
        # start inserting data here
    }
}

#region private functions
function TestTableExists {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)]
        [String] $TableName
    )
    $SQLQuery = "if exists (select * from sys.objects where object_id = OBJECT_ID(N'[dbo].[{0}]') AND type in (N'U'))
    select 1 else select 0" -f $TableName
    Write-Debug -Message $SQLQuery
    [bool](InvokeSQLQuery -SQLQuery $SQLQuery).Column1
}

function CreateTableFromObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)]
        [PSCustomObject] $InputObject,

        [Parameter(Mandatory,ValueFromPipeline)]
        [String] $TableName
    )
    $Columns = $InputObject | Get-Member -MemberType NoteProperty | ForEach-Object -Process {
        $datatype = $_.Definition.split(' ')[0]
        if ($datatype -eq 'String' -or $datatype -eq 'object') {
            $type = 'varchar(max)'
        } elseif ($datatype -eq 'datetime') {
            $type = 'datetime'
        } elseif ($datatype -eq 'timespan') {
            $type = 'time'
        } elseif ($datatype -eq 'bool') {
            $type = 'bit'
        }
        $columnname = $_.Name
        '{0} {1},' -f $columnname,$type
    }
    $SQLQuery = 'CREATE TABLE {0}(InsertID int IDENTITY(1,1) PRIMARY KEY, {1})' -f $TableName,($Columns | Out-String)
    Write-Output -InputObject $SQLQuery
}

function CreateInsertFromObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,ValueFromPipeline)]
        [PSCustomObject] $InputObject,

        [Parameter(Mandatory,ValueFromPipeline)]
        [String] $TableName
    )
    process {
        $columns = [string]::Empty
        $values = [string]::Empty
        $InputObject | Get-Member -MemberType NoteProperty | ForEach-Object -Process {
            $columns += "$($_.Name),"
            $value = "$($InputObject.($_.Name) -as [String])"
            if (-not $value) {
                $values += 'NULL,'
            } else {
                $values += "`'$value`',"
            }
        }
        $SQLQuery = 'INSERT INTO {0}({1}) VALUES ({2})' -f $TableName,$columns.TrimEnd(','),$values.TrimEnd(',')
        Write-Output -InputObject $SQLQuery
    }
}

function InvokeSQLQuery {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [String] $SQLQuery
    )
    $Command = New-Object System.Data.SQLClient.SQLCommand
    $Command.Connection = $script:DatabaseConnection
    $Command.CommandText = $SQLQuery
    try {
        $Reader = $Command.ExecuteReader()
        $Datatable = New-Object System.Data.DataTable
        $Datatable.Load($Reader)
        Write-Output -InputObject $Datatable
    } catch {
        Write-Error -ErrorRecord $_
    }
}

function TestAuthenticated {
    if (($null -ne $script:Token) -and ($script:ValidUntil -gt [datetime]::Now)) {
        return $true
    } else {
        return $false
    }
}

function ConvertDateTime {
    param (
        [String] $InputString
    )
    if ($InputString.Length -eq 15) {
        $Convert = $InputString.Split('T') -join " "
        return [datetime]::ParseExact($Convert,'yyyyMMdd HHmmss',$null)
    } elseif ($InputString.Length -eq 8) {
        return [datetime]::ParseExact($InputString,'yyyyMMdd',$null)
    } elseif ($InputString.Length -eq 10) {
        return [datetime]::ParseExact($InputString,'dd-MM-yyyy',$null)
    } else {
        [datetime]::Parse($InputString)
    }
}

function ConvertXMLElement {
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Xml.XmlElement] $Element
    )
    $OutObj = New-Object -TypeName hashtable
    $Element | Get-Member -MemberType Properties | ForEach-Object -Process {
        $Key = $_.Name
        if ($null -ne $Element.($_.Name)) {
            $Value = $Element.($_.Name).ToString()
        } else {
            $Value = $null
        }
        if ([string]::Empty -eq $Value) {
            $value = $null
        }
        if ($Key -like '*date*' -and $null -ne $value) {
            try {
                $OutObj.Add($Key,(ConvertDateTime $Value))
            } catch {
                $OutObj.Add($Key,$Value)
            }
        } elseif ($Value -eq 'yes' -or $Value -eq 'ja') {
            $OutObj.Add($Key,$true) 
        } elseif ($Value -eq 'no' -or $Value -eq 'nee') {
            $OutObj.Add($Key,$false) 
        } elseif ($Key -like '*time*'-and $null -ne $Value) {
            try {
                $OutObj.Add($Key,[timespan]::Parse($Value)) 
            } catch {
                $OutObj.Add($Key,$Value)
            }
        } elseif ($Key -like '*hour*' -and $null -ne $value) {
            try {
                $OutObj.Add($Key,[int]$Value)
            } catch {
                $OutObj.Add($Key,$Value)
            }
        } else {
            $OutObj.Add($Key,$Value)
        }
    }
    [PSCustomObject] $OutObj
}
#endregion private functions
#Export-ModuleMember -Function *-TimeWax*
Export-ModuleMember -Function *
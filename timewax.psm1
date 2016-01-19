$APIUri = 'https://api.timewax.com/'
$Token = $null
$ValidUntil = $null

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
            $Response = (Invoke-RestMethod -Uri $TokenUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
            if ($Response.valid -eq 'no') {
                Write-Error -Message "Did not acquire a token. Exception: $($Response.errors.'#cdata-section')" -ErrorAction Stop
            } else {
                Set-Variable -Scope 1 -Name Token -Value $Response.token
                Set-Variable -Scope 1 -Name ValidUntil -Value (ConvertDateTime $Response.validUntil)
            }
        }
    }
}

#should output custom object instead of XML elements.
#//TODO: Create c# class to serialize to
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
                    Write-Output -InputObject $r
                }
            } else {
                Write-Output -InputObject $Response.resource
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
                Write-Output -InputObject $e
            }
        }
    }
}

function Get-TimeWaxProject {
    [CmdletBinding()]
    param (

    )
    begin {
        if (-not (TestAuthenticated)) {
            Write-Error -Message "Token was not valid or not found. Run Get-TimeWaxToken" -ErrorAction Stop
        }
        $ProjectUri = $script:APIUri + 'project/list/'
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token> 
        </request>' -f $script:Token)
        $Response = (Invoke-RestMethod -Uri $ProjectUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            foreach ($P in $Response.projects.project) {
                Write-Output -InputObject $P
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
                Write-Output -InputObject $e
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
        $CompanyUri = $script:APIUri + 'company/list/'
    } process {
        $Body = [xml]('<request> 
         <token>{0}</token>
        </request>' -f $script:Token)
        $Response = (Invoke-RestMethod -Uri $CompanyUri -Method Post -Body $Body -ContentType application/xml -UseBasicParsing).response
        if ($Response.valid -eq 'no') {
            Write-Error -Message "$($Response.errors.'#cdata-section')" -ErrorAction Stop
        } else {
            foreach ($C in $Response.companies.company) {
                if (($PSCmdlet.ParameterSetName -eq 'Named') -and ($C.name -ne $Name)) {
                    continue
                } elseif (($PSCmdlet.ParameterSetName -eq 'Code') -and ($C.code -ne $Code)) {
                    continue
                }
                Write-Output -InputObject $C
            }
        }
    }
}

#region private functions
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
    $Convert = $InputString.Split('T') -join " "
    return [datetime]::ParseExact($Convert,'yyyyMMdd HHmmss',$null)
}
#endregion private functions
Export-ModuleMember -Function *-TimeWax*
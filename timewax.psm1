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
                Write-Error -Message "Did not acquire a token. Exception: $($Response.errors)" -ErrorAction Stop
            } else {
                Set-Variable -Scope 1 -Name Token -Value $Response.token
                Set-Variable -Scope 1 -Name ValidUntil -Value $Response.validUntil
            }
        }
    }
}

function Get-TimeWaxResource {
    [cmdletbinding(DefaultParameterSetName='Named')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName='Named')]
        [Alias('email','fullname')]
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
            Write-Error -Message "$($Response.error)" -ErrorAction Stop
        } else {
            if ($List) {
                foreach ($r in $Response.resources) {
                    Write-Output -InputObject $r.resource
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
        [Parameter(Mandatory, ParameterSetName='Resource')]
        [ValidateNotNullOrEmpty()]
        [String] $Resource,

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
        if ($Resource) {
            $child = $Body.CreateElement("resource")
            $child.InnerText = $Resource
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
            Write-Error -Message "$($Response.errors)" -ErrorAction Stop
        } else {
            foreach ($e in $Response.Entries) {
                Write-Output -InputObject $e.entry
            }
        }
    }
}

#private functions
function TestAuthenticated {
    if ($null -ne $script:Token) {
        #//TODO: Build in check for validity time
        return $true
    } else {
        return $false
    }
}

Export-ModuleMember -Function *-TimeWax*
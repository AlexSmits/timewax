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
            $TokenUri = $APIUri + 'authentication/token/get/'
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
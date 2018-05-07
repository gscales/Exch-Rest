function Set-EXRItemCategory {
    [CmdletBinding()]
    param (

        [parameter(ValueFromPipeline = $True)]
        [psobject[]]$Item,
		
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,

        [Parameter(Position = 2, Mandatory = $false)]
        [String]
        $ItemId,
		
        [Parameter(Position = 3, Mandatory = $false)]
        [psobject]
        $AccessToken,

        [Parameter(Position = 4, Mandatory = $false)]
        [Object[]]
        $Categories	
        
        
    )
    Process {
        Write-Host $Item
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        } 
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL = $EndPoint + "('" + $MailboxName + "')/messages"
        if ($Item -eq $null) {
            $RequestURL = $EndPoint + "('" + $MailboxName + "')/messages" + "('" + $ItemId + "')"
        }
        else {
            $RequestURL = $EndPoint + "('" + $MailboxName + "')/messages" + "('" + $Item.Id + "')"
        }
        $Update = "{`r`n`"categories`":["
        $first = $true
        foreach ($Item in $Categories) {
            if (!$first) {
                $Update += ",`r`n"
            }
            else {$first = $false}        
            $Update += "`"" + $Item + "`""
        }
        $Update += "]}`r`n"
        return Update-EXRItem -ItemURI $RequestURL -MailboxName $MailboxName -AccessToken $AccessToken -details $Update
    }
}

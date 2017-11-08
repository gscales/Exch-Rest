function  Find-EXRRooms {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [String]$filter      
    )
    Begin{
        
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName          
        }        
        $HttpClient =  Get-EXRHTTPClient -MailboxName $MailboxName
        $EndPoint =  Get-EXREndPoint -AccessToken $AccessToken -Segment "users"
        $RequestURL =  $EndPoint + "('" + $MailboxName + "')/findrooms?`$Top=1000"
        if(![String]::IsNullOrEmpty($filter)){
                $RequestURL =  $EndPoint + "('" + $MailboxName + "')/findrooms?`$Top=1000&`$filter=" + $filter
        }
        do{
            $JSONOutput = Invoke-EXRRestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
            foreach ($Message in $JSONOutput.Value) {
                Write-Output $Message
            }           
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }while(![String]::IsNullOrEmpty($RequestURL))     
    }
}

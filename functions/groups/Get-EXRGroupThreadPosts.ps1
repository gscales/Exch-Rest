function Get-EXRGroupThreadPosts {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
		
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $AccessToken,

        [Parameter(Position = 2, Mandatory = $false)]
        [psobject]
        $Group,
		
        [Parameter(Position = 3, Mandatory = $false)]
        [psobject]
        $ThreadId,

        [Parameter(Position = 4, Mandatory = $false)]
        [Int]
        $Top = 1000,

        [Parameter(Position = 5, Mandatory = $false)]
        [psobject]
        $PropList,

        [Parameter(Position = 6, Mandatory = $false)]
        [string]
        $Filter,

        [Parameter(Position = 5, Mandatory = $false)]
        [string]
        $Extensions


    )
    Process {
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        }
        $PropList = Get-EXRKnownProperty -PropList $PropList -PropertyName "PR_BODY_HTML"
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "groups"
        $RequestURL = $EndPoint + "('" + $Group.Id + "')/Threads('" + $ThreadId + "')/Posts?`$Top=$Top"
        if(![String]::IsNullorEmpty($Filter)){
            $Filter = "`&`$filter=" + [System.Web.HttpUtility]::UrlEncode($Filter)
            $RequestURL += $Filter
        }
        if($PropList){
            $Props = Get-EXRExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
            $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            if($Extensions){
                $RequestURL += ",Extensions(`$filter=Id eq '$Extensions')"
            }
        }else{
            $RequestURL += "&`$expand=Extensions(`$filter=Id eq '$Extensions')"
        }
        do {
            $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName 
            foreach ($Message in $JSONOutput.Value) {
                $ItemURI =  $EndPoint + "('" + $Group.Id + "')/Threads('" + $ThreadId + "')/Posts" + "('" + $Message.Id + "')"
                add-Member -InputObject $Message -NotePropertyName ItemURI -NotePropertyValue $ItemURI
                Expand-ExtendedProperties -Item $Message
                Write-Output $Message
            }
            $RequestURL = $JSONOutput.'@odata.nextLink'
        }
        while (![String]::IsNullOrEmpty($RequestURL))	
    }
}

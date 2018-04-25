function Get-EXRLastInboxEmail {
    [CmdletBinding()]
    param( 
        [Parameter(Position = 0, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 3, Mandatory = $false)] [String]$First,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$Unread,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Focused,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$Other,
        [Parameter(Position = 12, Mandatory = $false)] [switch]$ReturnSize,
        [Parameter(Position = 13, Mandatory = $false)] [switch]$ReturnAttachments,
        [Parameter(Position=14, Mandatory=$false)] [switch]$ReturnSentiment

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
        $Filter = $null
        if ([String]::IsNullOrEmpty($First)) {
            $First = 1
        }    
        if ($Unread.IsPresent) {
            if ($Filter -eq $null) {
                $Filter = "isread eq false"
            }    
        }
        $ClientFilter = $null
        $ClientFilterTop = $null
        if ($Focused.IsPresent) {
            $ClientFilter = "" | Select Property, Value, Operator
            $ClientFilter.Property = "inferenceClassification"
            $ClientFilter.Value = "focused"
            $ClientFilter.Operator = "eq"
            $ClientFilterTop = $First 
            $First = 100;
        }
        if ($Other.IsPresent) {
            $ClientFilter = "" | Select Property, Value, Operator
            $ClientFilter.Property = "inferenceClassification"
            $ClientFilter.Value = "other"
            $ClientFilter.Operator = "eq"
            $ClientFilterTop = $First 
            $First = 100;
        }
        Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder Inbox -TopOnly:$true -Top $First -Filter $Filter -ClientFilter $ClientFilter -ClientFilterTop $ClientFilterTop -ReturnAttachments:$ReturnAttachments.IsPresent -ReturnSize:$ReturnSize.IsPresent -ReturnSentiment:$ReturnSentiment.IsPresent
    }
}
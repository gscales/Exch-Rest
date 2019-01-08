function Get-EXRLastInboxEmail {
    [CmdletBinding()]
    param( 
        [Parameter(Position = 0, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 3, Mandatory = $false)] [String]$First,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$Unread,
        [Parameter(Position = 5, Mandatory = $false)] [switch]$Focused,
        [Parameter(Position = 6, Mandatory = $false)] [switch]$Other,
        [Parameter(Position = 7, Mandatory = $false)] [switch]$ReturnSize,
        [Parameter(Position = 8, Mandatory = $false)] [switch]$ReturnAttachments,
        [Parameter(Position = 9, Mandatory = $false)] [switch]$ReturnBody,
        [Parameter(Position=10, Mandatory=$false)] [switch]$ReturnSentiment,
        [Parameter(Position = 11, Mandatory = $false)] [String]$BodyFormat,
	[Parameter(Position = 12, Mandatory = $false)] [switch]$hasAttachment


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
            if ([String]::IsNullOrEmpty($Filter)) {
                $Filter = "isread eq false"
            }else{
                $Filter += " isread eq false"
            }    
        }
	if($hasAttachment.IsPresent){
        if ([String]::IsNullOrEmpty($Filter)) {
            $Filter = "hasAttachments eq true"
        }else{
            $Filter += " hasAttachments eq true"
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
        Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder Inbox -TopOnly:$true -Top $First -Filter $Filter -ClientFilter $ClientFilter -ClientFilterTop $ClientFilterTop -ReturnAttachments:$ReturnAttachments.IsPresent -ReturnSize:$ReturnSize.IsPresent -ReturnSentiment:$ReturnSentiment.IsPresent -returnBody:$ReturnBody.IsPresent -BodyFormat $BodyFormat
    }
}
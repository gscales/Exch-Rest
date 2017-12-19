function Get-EXRLastInboxEmail{
    [CmdletBinding()]
    param( 
        [Parameter(Position=0, Mandatory=$false)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=3, Mandatory=$false)] [String]$Top,
        [Parameter(Position=4, Mandatory=$false)] [switch]$Unread,
        [Parameter(Position=5, Mandatory=$false)] [switch]$Focused,
        [Parameter(Position=6, Mandatory=$false)] [switch]$Other

    )
    Process{
    if($AccessToken -eq $null)
    {
        $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
        if($AccessToken -eq $null){
            $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
        }                 
    }
    if([String]::IsNullOrEmpty($MailboxName)){
        $MailboxName = $AccessToken.mailbox
    } 
    $Filter= $null
    if([String]::IsNullOrEmpty($Top)){
        $Top = 1
    }    
    if($Unread.IsPresent){
        if($Filter -eq $null){
            $Filter = "isread eq false"
        }    
    }
    $ClientFilter = $null
    $ClientFilterTop = $null
    if($Focused.IsPresent){
        $ClientFilter = "" | Select Property,Value,Operator
        $ClientFilter.Property = "inferenceClassification"
        $ClientFilter.Value = "focused"
        $ClientFilter.Operator = "eq"
        $ClientFilterTop = $Top 
        $Top = 1000;
    }
    if($Other.IsPresent){
        $ClientFilter = "" | Select Property,Value,Operator
        $ClientFilter.Property = "inferenceClassification"
        $ClientFilter.Value = "other"
        $ClientFilter.Operator = "eq"
        $ClientFilterTop = $Top 
        $Top = 1000;
    }
    if($Top -eq 1){
	    $Items = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder Inbox -TopOnly:$true -Top $Top -Filter $Filter -ClientFilter $ClientFilter -ClientFilterTop $ClientFilterTop
	    Get-EXREmail -ItemRESTURI $Items[0].ItemRESTURI -MailboxName $MailboxName -AccessToken $AccessToken
    }
    else{
        Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder Inbox -TopOnly:$true -Top $Top -Filter $Filter -ClientFilter $ClientFilter -ClientFilterTop $ClientFilterTop
    }
    }
}
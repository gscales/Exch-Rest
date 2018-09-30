function Invoke-EXRFillMailboxFolder {
   	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$FolderPath,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 2, Mandatory = $false)]
		[psobject]
        $AccessToken,
        
        [Parameter(Position = 3, Mandatory = $false)]
		[String]
		$MaxItems 
		
	)
	Process
	{
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
       $MaxItemCount = 100
       if(![String]::IsNullOrEmpty($MaxItems)){
            $MaxItemCount = [Int]::Parse($MaxItems)
       }
       $Folder = Get-EXRFolderFromPath -MailboxName $MailboxName -AccessToken $AccessToken -FolderPath $FolderPath
       $rssItemCount =0
       $XMLDoc = $null
       $rssPageNumber = 0
       for($itCount=0;$itCount -lt $MaxItemCount;$itCount++){
          if($rssItemCount -eq 0){
              $rssPageNumber++
              $Content = Invoke-WebRequest -uri https://blogs.msdn.microsoft.com/feed/?page=$rssPageNumber
              $XMLDoc = [XML]$Content.Content
              
          }
          $XMLDoc.rss.channel.Item[$rssItemCount].Title
          $rssItemCount++
          if($rssItemCount -eq 10){$rssItemCount = 0}
          Start-Sleep -Seconds 2
       }



        
    }
}

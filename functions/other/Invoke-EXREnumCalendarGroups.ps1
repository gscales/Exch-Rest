function Invoke-EXREnumCalendarGroups
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[psobject]
		$AccessToken
	)
	Begin
	{
		
		if ($AccessToken -eq $null)
		{
			$AccessToken = Get-EXRAccessToken -MailboxName $MailboxName
		}
		$HttpClient = Get-HTTPClient -MailboxName $MailboxName
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"
		$RequestURL = $EndPoint + "/" + $MailboxName + "/CalendarGroups"
		$JsonObject = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		foreach ($Group in $JsonObject.Value)
		{
			Write-Host ("GroupName : " + $Group.Name)
			$GroupId = $Group.Id.ToString()
			$RequestURL = $EndPoint + "/" + $MailboxName + "/CalendarGroups('$GroupId')/Calendars"
			$JsonObjectSub = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Calendar in $JsonObjectSub.Value)
			{
				Write-Host $Calendar.Name
			}
			$RequestURL = $EndPoint + "/" + $MailboxName + "/CalendarGroups('$GroupId')/MailFolders"
			$JsonObjectSub = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
			foreach ($Calendar in $JsonObjectSub.Value)
			{
				Write-Host $Calendar.Name
			}
			
		}
		
		
	}
}

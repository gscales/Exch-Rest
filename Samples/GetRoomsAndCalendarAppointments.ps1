$MailboxName = "gscales@datarumble.com"
$Token = Get-AccessToken -MailboxName $MailboxName  -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl urn:ietf:wg:oauth:2.0:oob -ResourceURL graph.microsoft.com -beta  
$rptCollection = @()
Find-Rooms -Mailbox $MailboxName  -AccessToken $Token | foreach-object{
	$RoomAddress =  $_.address
	Get-CalendarView -MailboxName $_.address -StartTime (Get-Date) -EndTime (Get-Date).AddDays(1) -AccessToken $Token  | foreach-object{
		$rptObj = "" | Select Room,Organizer,Subject,Start,End
		$rptObj.Room = $RoomAddress 
                $rptObj.Organizer = $_.organizer.emailAddress.name
		$rptObj.Subject = $_.subject
		$rptObj.Start = [DateTime]::Parse($_.start.datetime).ToString("yyyy-MM-ddTHH:mm:ss")
		$rptObj.End =  [DateTime]::Parse($_.end.datetime).ToString("yyyy-MM-ddTHH:mm:ss")
		$rptCollection += $rptObj

	}
}
$rptCollection

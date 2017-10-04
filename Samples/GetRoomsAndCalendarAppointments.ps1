
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
$tableStyle = @" 
<style> 
BODY{background-color:white;} 
TABLE{border-width: 1px; 
  border-style: solid; 
  border-color: black; 
  border-collapse: collapse; 
} 
TH{border-width: 1px; 
  padding: 10px; 
  border-style: solid; 
  border-color: black; 
  background-color:#66CCCC 
} 
TD{border-width: 1px; 
  padding: 2px; 
  border-style: solid; 
  border-color: black; 
  background-color:white 
} 
</style> 
"@  
    
$body = @" 
<p style="font-size:25px;family:calibri;color:#ff9100">  
$TableHeader  
</p>  
"@  
$rptCollection |  ConvertTo-Html -head $tableStyle –body $body| Out-File "c:\temp\RoomReport.html"

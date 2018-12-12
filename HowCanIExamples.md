# How Can I Examples #

How can I can the last Email in my Mailbox ?

    Get-EXRLastInboxEmail

How can I can I get the Last Email in a Shared Mailbox ?

    Get-EXRLastInboxEmail -MailboxName mailbox@domain.com

How can I show the Last email from the Focused Inbox

    Get-EXRLastInboxEmail -MailboxName mailbox@domain.com -Focused

How can I read the body of the Last email as Text

    Get-EXRLastInboxEmail -ReturnBody -BodyFormat Text | Select Body | fl

How can I move the Last email in my Inbox to another folder

	$LastMail = Get-EXRLastInboxEmail 	
	Move-EXRMessage -ItemURI $LastMail.ItemRESTURI -TargetFolder \aa

How can I get the Size of the Inbox

    Get-EXRWellKnownFolder -FolderName Inbox | Select displayName,totalItemCount,FolderSize
    

How can I Mark the last email in the Focused Inbox as Read

    Get-EXRLastInboxEmail -Focused | Invoke-EXRMarkEmailAsRead

How can I Send a Skype for Business Message

	Connect-EXRSK4B  #Only needs to be done once per session
    Send-EXRSK4BMessage -Subject "Message Subject" -Message "The isthe Message" -ToSipAddress jcool@datarumble.com

How can I check a users Skype for Business Presence 

	Connect-EXRSK4B  #Only needs to be done once per session
    Get-EXRSK4BPresence -TargetUser e5tmp5@datarumble.com

How can I export my Contacts to a CSV file

    Export-EXRContactFolderToCSV -FileName c:\contacts\myContacts.csv

How can I get this weeks Calendar Appointments 

    Get-EXRNamedCalendarView -MailboxName gscales@datarumble.com

How can I get this Months Calendar Appointments 

     Get-EXRNamedCalendarView -MailboxName gscales@datarumble.com -StartTime (Get-Date).Date -EndTime (Get-Date).AddMonths(1)

How can I get a Report of Mailbox sizes
    
    Get-EXRMailboxUsage
	
How can I get the last 10 deleted Items in a Mailbox and show the Folder they where deleted From (utilizes the LAPFID)

    Get-EXRDeletedItems -MessageCount 10 -ReturnLastActiveParentFolderPath | Select LastActiveParentFolderPath,SenderEmailAddress,Subject

How can I show the SPF,DKIM,DMARC on the last 10 messages in the JunkEmail Folder

     Get-EXRWellKnownFolderItems -WellKnownFolder JunkEmail -MessageCount 10 -ReturnInternetMessageHeaders -ProcessAntiSPAMHeaders | Select Subject,SPF,DKIM,DMARC | fl


   
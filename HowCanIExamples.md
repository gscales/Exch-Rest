# How Can I Examples #

How can I can the last Email in my Mailbox ?

    Get-EXRLastInboxEmail

How can I can I get the Last Email in a Shared Mailbox ?

    Get-EXRLastInboxEmail -MailboxName mailbox@domain.com

How can I show the Last email from the Focused Inbox

    Get-EXRLastInboxEmail -MailboxName mailbox@domain.com -Focused

How can I Mark the last email in the Focused Inbox as Read

    Get-EXRLastInboxEmail -Focused | Invoke-EXRMarkEmailAsRead

How can I export my Contacts to a CSV file

    Export-EXRContactFolderToCSV -FileName c:\contacts\myContacts.csv

How can I get this weeks Calendar Appointments 

    Get-EXRNamedCalendarView -MailboxName gscales@datarumble.com

How can I get a Report of Mailbox sizes
    
    Get-EXRMailboxUsage
	
   
With Version 2.7 of the Exch-REST Module you can now use the 

Start-EXRMailClient cmdlet which will start the small mail client which will use the underlying Exch-REST Module cmldets to open and browse Mailbox Mail folders, Read Items from a Folder you select in the TreeView, Show Internet Message Headers or Download attachments from messages in the datagrid eg



Getting Messages from the Inbox



Showing the content of a Message



or show the Message Headers




You can either start the Mailbox client using Start-EXRMailClient cmdlet with no parameters or pass in a Mailbox and AccessToken you maybe working with and it will automatically enumerate the Folders from that Mailbox using the token passed in.

To use the Read Message form on a message you have found in the cmdline through using another cmdlet eg to open the last message in the Inbox you could use


$Items = Get-WellKnownFolderItems -MailboxName gscales@datarumble.com -AccessToken $AccessToken -WellKnownFolder Inbox -TopOnly:$true -Top 1
Invoke-EXRReadEmail -ItemRESTURI $Items[0].ItemRESTURI -AccessToken $AccessToken -MailboxName gscales@datarumble.com
To Start the New Message form just use

Invoke-EXRNewMessagesForm -MailboxName gscales@datarumble.com -AccessToken $AccessToken




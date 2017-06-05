To send a Message you need to at least provide a Subject and recipients in either the To, CC or BCC properties 

Pre Requsites that you have obtained an Access Token using of the Access Token method from https://github.com/gscales/Exch-Rest/blob/master/README.md

Example 1 Send a Message to one recipient

Send-MessageREST -MailboxName user@demaildomain.com  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address user@ydomain.com) -Subject test123 -Body "test 123"

Example 2 Send a Message as Another user that a Mailbox has been delegaete rights to either via SendAS or SendOnBehalf

Send-MessageREST -MailboxName user@dmaildomain.com  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address user@ydomain.com) -Subject test123 -Body "test 123" -SenderEmailAddress (New-EmailAddress -Address SendingAs@dmaildomain.com

Example 3 Send a Message and set the ReplyTo address

Send-MessageREST -MailboxName user@dmaildomain.com  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address user@ydomain.com) -Subject test123 -Body "test 123" -ReplyTo (New-EmailAddress -Address replyTo@ydomain.com) 

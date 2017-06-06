To send a Message you need to at least provide a Subject and recipients in either the To, CC or BCC properties 

Pre Requsites that you have obtained an Access Token using of the Access Token method from https://github.com/gscales/Exch-Rest/blob/master/README.md

Switches

```
 -SaveToSentItems : Saves the message being sent to the SentItems folder of calling account -MailboxName
 -ShowRequest : Used for debuging outputs the REST message that is being sent to the server
```
Example 1 Send a Message to one recipient
````
Send-MessageREST -MailboxName user@demaildomain.com  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address user@ydomain.com) -Subject test123 -Body "test 123"
```
Example 2 Send a Message as Another user that a Mailbox has been delegaete rights to either via SendAS or SendOnBehalf
```
Send-MessageREST -MailboxName user@dmaildomain.com  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address user@ydomain.com) -Subject test123 -Body "test 123" -SenderEmailAddress (New-EmailAddress -Address SendingAs@dmaildomain.com
```
Example 3 Send a Message and set the ReplyTo address
```
Send-MessageREST -MailboxName user@dmaildomain.com  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address user@ydomain.com) -Subject test123 -Body "test 123" -ReplyTo (New-EmailAddress -Address replyTo@ydomain.com) 
```
Example 4 Send a Message to one recipient and one attachment
````
Send-MessageREST -MailboxName user@demaildomain.com  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address user@ydomain.com) -Subject test123 -Body "test 123" -Attachments @("c:\temp\excelattachment.csv")
```
Example 4 Send a Message with two recipients and two attachments
```
$ToRecipients = @(New-EmailAddress -Address user1@ydomain.com)
$ToRecipients += New-EmailAddress -Address user2@ydomain.com
$Attachments = @("c:\temp\excelattachment1.csv")
$Attachments += "c:\temp\excelattachment2.csv"
Send-MessageREST -MailboxName user@demaildomain.com  -AccessToken $AccessToken -ToRecipients $ToRecipients -Subject test123 -Body "test 123" -Attachments $Attachments
```
Example 5 Send a Message to one recipient and save it to SentItems (default is not to save)
```
Send-MessageREST -MailboxName user@demaildomain.com  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address user@ydomain.com) -Subject test123 -Body "test 123" -SaveToSentItems
```
Example 6 Send a Message to one To recipient and one BCC recipient and save it to SentItems 
```
Send-MessageREST -MailboxName user@demaildomain.com  -AccessToken $AccessToken -ToRecipients @(New-EmailAddress -Address user@ydomain.com) -BCCRecipients @(New-EmailAddress -Address bccRecip@ydomain.com) -Subject test123 -Body "test 123" -SaveToSentItems
```

# Searching for Items using the Exch-REST Module #

Within an Office365 Mailbox there are number of different methods for performing a Search using the REST API for particular items within a particular Mailbox folder or across all folder in a mailbox. Depending on the properties you want to search on and the amount of data you expect to return in your search you should select the most appropriate method to use.

## Using filters (or Restrictions) ##

Restrictions are the standard way in which searches are normally carried out in Outlook and Exchange, restrictions are property based searches on one or more properties and may involve a number of operators such as equals, greatorthan, lessthan, contains (for substrings) or startswith. Because the raw content is queried this type of search can take longer to perform if the underlying folder has a larger number of items or the restriction being used is very complicated. A common issue that can occur using filters is the Query will timeout or will be throttled because of the overall load it produces on the server.

## Using KQL (or Exchange Search) ##

Exchange Search is the full text indexing feature of Exchange which indexes contents automatically (asynchronously) as they arrive in a Mailbox.This type of Search utilizes the KQL (Keyword Query Logic) that is also used with in Share-point for query processing. This is the faster and more efficient method of searching a Mailbox because it utilizes Full Text Indexes rather then needing to query the raw content and should be the preferred starting point for search. However there are some limitations on the properties that can be searched using this method and also on the amount of results that will be returned by a query. 

Property restrictions : As Exchange only indexes certain properties for performance reason only certain properties can be searched using KQL a full list of the indexed properties the can be used can be found [https://technet.microsoft.com/en-us/library/jj983804(v=exchg.150).aspx](https://technet.microsoft.com/en-us/library/jj983804(v=exchg.150).aspx)

Maximum Result Sets: Exchange limits the maximum result set of Exchange search queries to 250 Items. with OnPrem Servers this value is adjustable via the MaxHitsForFullTextIndexSearches property [https://support.microsoft.com/en-us/help/3093866/the-number-of-search-results-can-t-be-more-than-250-when-you-search-em](https://support.microsoft.com/en-us/help/3093866/the-number-of-search-results-can-t-be-more-than-250-when-you-search-em) however in Exchange OnLine (Office365) you can't adjust that value and need to work under the 250 item ceiling.

## Using SearchFolders ##

Search Folders are Special Mailbox folders that contain no items but return linked Items based on a predefined search criteria (usually the same as you would use in Filters). Because these are constantly updated asynchronously by the Exchange Server they offer better performance when retrieving items based on a search criteria then doing a normal search using a filter. Search Folders are a good choice if you have static queries that don't require any dynamic input or you need to have a search that spans multiple folders. You do need to be aware that overuse of Search Folders can cause poor performance and they are updated by a background process so new items may not be instantly available as apposed to querying the Mailbox folder directly. Some more background information on Search Folders can be found

[https://blogs.msdn.microsoft.com/dgoldman/2008/07/01/microsoft-exchange-and-search-folders/](https://blogs.msdn.microsoft.com/dgoldman/2008/07/01/microsoft-exchange-and-search-folders/)
[https://blogs.msdn.microsoft.com/webdav_101/2015/05/03/ews-best-practices-searching/](https://blogs.msdn.microsoft.com/webdav_101/2015/05/03/ews-best-practices-searching/)

Currently there is no way of creating a SearchFolder programmaticly using the REST API but this can be done using EWS. 

## Search Examples using the Search-EXRMessage cmdlet##

Several of the cmdlets allow the entering of both the Filter and Search criteria however the cmdlet specifically designed for search is the Search-EXRMessage cmdlet

**Using Raw KQL**

The -KQL Parameter allows you enter raw KQL to be used as a Mailbox Search if you do use this approach you should take care that some characters need to be escaped correct. Eg when search for an exact phrase in KQL the phrase should be encolesd in double quotes. When doing this in REST these double quotes need to be escape. Eg so if you where search for email with the subject has the following keywords you could use **-KQL "subject:'termone termtwo'"** where if you want to search for a phrase you should use  **-KQL 'subject:\"this is a phrase\""**

**Specifying the folders to Search**

By default the Search-EXRMessage cmdlet will search all the Mail folders in a Mailbox using the AllItems Search Folder if no folder is specified in the cmdline. To limit the Search to a particular folder you can specify the folder using the -FolderPath eg 

Search-EXRMessage -FolderPath \Inbox -KQL 'subject:\"Happy Days\"'

or use the  -Wellknown parameter to specify one of the default folder such as Inbox or SentItems eg

Search-EXRMessage -Wellknow SentItems -KQL 'subject:\"Happy Days\"'

**Returning Attachments details within Search Results**

Returning attachment details when search for objects requires extra requests be made to the server which will affect the performance of search greatly so this isn't done by default. However if its import for you to get this information because of the type of search that you are trying to do then you can use the -ReturnAttachments swich. Not when using the -AttachmentKQL switch this is submitted by default.


**Finding a Message from a Particular Sender**

Find a Message from a Particular Sender by Name or Email Address using a Filter

by Email

    Search-EXRMessage -MailboxName user@domain -from 'James@domain.com'  | select Subject,SenderEmailAddress,SenderName,FolderPath



Find a messages from a Particular Sender by Name or Email Address using KQL

by Name

    Search-EXRMessage -MailboxName user@domain -kql 'From:James'  | select Subject,SenderEmailAddress,SenderName,FolderPath

by Email Address

    Search-EXRMessage -MailboxName user@domain -kql 'From:James@domain.com'  | select Subject,SenderEmailAddress,SenderName,FolderPath

**Find a Message from a Particular Sender on a Particular Date**

using a Filter

    Search-EXRMessage -MailboxName user@domain -ReceivedtimeFrom ([DateTime]::Parse('2018-01-01')) -ReceivedtimeTo ([DateTime]::Parse('2018-01-02')) -from 'james@domain.com' | select Subject,SenderEmailAddress,SenderName,FolderPath

using KQL

    Search-EXRMessage -MailboxName user@domain -ReceivedtimeFromKQL ([DateTime]::Parse('2018-01-01')) -ReceivedtimeToKQL ([DateTime]::Parse('2018-01-02')) -KQL 'from:james@domain.com' | select Subject,SenderEmailAddress,SenderName,FolderPath

**Find a Message with a Particular Subject (Exact Match)**

using a Filter

    Search-EXRMessage -MailboxName user@domain -subject 'And the Subject is' | select Subject,SenderEmailAddress,SenderName,FolderPath

using KQL

    Search-EXRMessage -MailboxName user@domain -SubjectKQL 'And the Subject is' | select Subject,SenderEmailAddress,SenderName,FolderPath

**Find a Message where the Subject contains a phrase**

using a Filter

    Search-EXRMessage -MailboxName user@domain -SubjectContains 'And the Subject' | select Subject,SenderEmailAddress,SenderName,FolderPath

using KQL

    Search-EXRMessage -MailboxName user@domain -KQL "Subject:'And the Subject'" | select Subject,SenderEmailAddress,SenderName,FolderPath


**Find a Message where the Subject starts with a phrase**
using a Filter

    Search-EXRMessage -MailboxName user@domain -SubjectStartsWith 'And the Subject' | select Subject,SenderEmailAddress,SenderName,FolderPath

**Find a Message where the Body contains a phrase**

using a Filter

    Search-EXRMessage -MailboxName user@domain -BodyContains 'Summer Sales' | select Subject,SenderEmailAddress,SenderName,FolderPath

using KQL

    Search-EXRMessage -MailboxName user@domain -BodyKQL "Sumber Sale" | select Subject,SenderEmailAddress,SenderName,FolderPath


**Find Attachments**

You can search for attachments using the full attachment name or a partial name like the extension if you want to find all email with word documents for example.

Using KQL

    Search-EXRMessage -MailboxName user@domain -AttachmentKQL ".doc" -ReturnAttachments 

## Using SearchFolders ##

To use a Search folder you first need to know it FolderId value, SearchFolder should be available under the NON_IPM_Subtree Finder folder. To get the available search folders you can use in a Mailbox you can use the Get-EXRSearchFolders Cmldet eg 

    Get-EXRSearchFolders -MailboxName user@domain.com

To Restrict the results to a particular name SearchFolder you can use

    Get-EXRSearchFolders -MailboxName user@domain.com -FolderName 'Voice Mail'

To enumerate Items from the Voice Mail SearchFolder use

    $VoiceMail = Get-EXRSearchFolders -MailboxName user@domain.com -FolderName 'Voice Mail'
    Get-EXRFolderItems  -MailboxName user@domain.com -Folder $voicemail
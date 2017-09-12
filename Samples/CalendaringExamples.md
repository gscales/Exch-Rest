# Calendaring Documentation and Examples

The Exch-Rest module contains the following Calendaring Cmdlets to help peform one of many calendaring operations using REST

#First Step before you can use any of the calendaring cmdlets is to Create an AcessToken to use for Authentication see https://github.com/gscales/Exch-Rest#authentication

# Recurrence

When you want to create a recurring Appointment you need to create a recurrence structure that can be used in one of the Calendaring cmdlets. The recurrence structure is documented https://github.com/microsoftgraph/microsoft-graph-docs/blob/master/api-reference/beta/resources/patternedrecurrence.md and https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar#PatternedRecurrence . The Get-Recurrence cmdlet creates a recurrence structure that you can then feed into the other event cmdlets
```
Get-Recurrence Parameters

-RecurrenceTimeZone (TimeZone for the Recurrence your create eg for the local Timezone [TimeZoneInfo]::Local.Id
-PatternType Recurrence patternType valid values "daily","weekly","absoluteMonthly","relativeMonthly", "absoluteYearly"," relativeYearly"
-PatternInterval The number of times the given recurrence type between occurrences.
-PatternMonth The month that the item occurs on. This is a number from 1 to 12.
-PatternDayOfMonth which day of the month the recurrance falls on when using absoluteMonthly
-PatternFirstDayOfWeek First day of week of pattern valid values "sunday","monday","tuesday","wednesday", "thursday","friday","saturday"
-PatternDaysOfWeek Which days of the week the recurrence falls on should be an array of days in lowercase
-PatternIndex valid values "first","second","third","fourth", "last"      
-RangeType valid values "noend"," enddate","numbered"
-RangeStartDate DateTime of the start of the recurrence range mandatory for every recurrance
-RangeEndDate DateTime of the end of the recurrence range 
-RangeNumberOfOccurrences number of recurrences when RangeType is numbered 

```


Examples of Creating the Recurrence structure using the Get-Recurrence cmdlet
```
#Creates a Recurrance Structure to be used in Calendaring Cmdlets 
#Daily recurrance eveny day
$Recurrence = Get-Recurrence -PatternType daily -PatternFirstDayOfWeek monday -PatternIndex first -RangeType noend -RangeStartDate (Get-Date)

#Create a weekly recurrance for every Monday End in two months
$days = @()
$days+"monday"
$Recurrence = Get-Recurrence -PatternType weekly -PatternFirstDayOfWeek monday -PatternDaysOfWeek $days -PatternIndex first  -RangeType enddate -RangeStartDate (Get-Date) -RangeEndDate (Get-Date).AddMonth(2)
```


# New-HolidayEvent

Creates a new Holiday event (Basically creates an All Day event on one day)

```
#Creates a new all day event for Today on the default calendar
New-HolidayEvent -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Day (Get-Date) -Subject "Name of Holiday"

#Create a new all day event for Today on the Australia holidays calendar
New-HolidayEvent -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Day (Get-Date) -Subject "Name of Holiday" -CalendarName 'Australia holidays'

#Creates a new recurring all day event for the 5th July to remember the cats birthday each year for 10 years on the default calendar
$Recurrence = Get-Recurrence -PatternType absoluteYearly -PatternMonth 7 -PatternDayOfMonth 5 -PatternFirstDayOfWeek monday -PatternIndex first  -RangeType enddate -RangeStartDate ([DateTime]::Parse("2017-07-05")) -RangeEndDate ([DateTime]::Parse("2027-07-05"))
New-HolidayEvent -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Day [DateTime]::Parse("2017-07-05") -Subject "Cats Birthday" -Recurrence $Recurrence
```


# New-CalendarEventREST

Create a Single Appointment on the default Calendar for the 5th of July between 1 and 2pm
```
#Creates a new all day event for Today on the default calendar
New-CalendarEventREST -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Start ([DateTime]::Parse("2017-07-05 13:00")) -End ([DateTime]::Parse("2017-07-05 14:00")) -Subject "Name of Event"
```

Create a Recurring Appointment for each monday on the default Calendar between 1 and 2pm

```
$days = @()
$days+"monday"
$Recurrence = Get-Recurrence -PatternType weekly -PatternFirstDayOfWeek monday -PatternDaysOfWeek $days -PatternIndex first  -RangeType enddate -RangeStartDate ([DateTime]::Parse("2017-07-05 13:00")) -RangeEndDate ([DateTime]::Parse("2019-07-05 13:00"))
#Creates a new all day event for Today on the default calendar
New-CalendarEventREST -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Start ([DateTime]::Parse("2017-07-05 13:00")) -End ([DateTime]::Parse("2017-07-05 14:00")) -Subject "Monday Meeting" -Recurrence $Recurrence
```

Create a Single Appointment on a Secondary Calendar

```
#Creates a new all day event for Today on the a calendar called Secondary calendar
New-CalendarEventREST -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Start ([DateTime]::Parse("2017-07-05 13:00")) -End ([DateTime]::Parse("2017-07-05 14:00")) -Subject "Name of Event" -CalendarName 'Secondary calendar'
```

Create a Recurring Appointment on a Secondary Calendar

```
$days = @()
$days+"monday"
$Recurrence = Get-Recurrence -PatternType weekly -PatternFirstDayOfWeek monday -PatternDaysOfWeek $days -PatternIndex first  -RangeType enddate -RangeStartDate ([DateTime]::Parse("2017-07-05 13:00")) -RangeEndDate ([DateTime]::Parse("2019-07-05 13:00"))
#Creates a new all day event for Today on the default calendar
New-CalendarEventREST -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Start ([DateTime]::Parse("2017-07-05 13:00")) -End ([DateTime]::Parse("2017-07-05 14:00")) -Subject "Monday Meeting" -Recurrence $Recurrence -CalendarName 'Secondary calendar'
```

Create a Single Appointment on a Group Calendar Calendar

To create an appointment in a Group calendar you need to use the -Group switch and -GroupName to specify the name of the group

```
#Creates a new all event for the 5th of July from 13:00-14:00 on the a group calendar called 'Powershell Module'
New-CalendarEventREST -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Start ([DateTime]::Parse("2017-07-05 13:00")) -End ([DateTime]::Parse("2017-07-05 14:00")) -Subject "Name of Event" -group -groupname 'Powershell Module'
```


Create a Recurring Appointment on a Group Calendar Calendar

```
$days = @()
$days+"monday"
$Recurrence = Get-Recurrence -PatternType weekly -PatternFirstDayOfWeek monday -PatternDaysOfWeek $days -PatternIndex first  -RangeType enddate -RangeStartDate ([DateTime]::Parse("2017-07-05 13:00")) -RangeEndDate ([DateTime]::Parse("2019-07-05 13:00"))
#Creates a new all event for the recurring event on 5th of July from 13:00-14:00 on the a group calendar called 'Powershell Module'
New-CalendarEventREST -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Start ([DateTime]::Parse("2017-07-05 13:00")) -End ([DateTime]::Parse("2017-07-05 14:00")) -Subject "Name of Event" -group -groupname 'Powershell Module' -Recurrence $Recurrence
```

# Meetings

To create a Meeting you use the same REST message structure with the addition of the Attendees elements https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar#AttendeeBaseV2. Meeting Attendess can be Required, Optional or Resources. To pass attendees into the Meeting first create a collection
$Attendees = @()
 
Then Add Atteendess to that collection eg

$Attendees += (new-attendee -Name 'fred smith' -Address 'fred@datarumble.com' -type 'Required')
$Attendees += (new-attendee -Name 'barney jones' -Address 'barney@datarumble.com' -type 'Optional')

Then pass that in when creating a new Appontment

Create a Single Meeting with two attendees on the default Calendar for the 5th of July between 1 and 2pm
```
#Creates a new all day event for Today on the default calendar
$Attendees = @()
$Attendees += (new-attendee -Name 'fred smith' -Address 'fred@datarumble.com' -type 'Required')
$Attendees += (new-attendee -Name 'barney jones' -Address 'barney@datarumble.com' -type 'Optional')

New-CalendarEventREST -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Start ([DateTime]::Parse("2017-07-05 13:00")) -End ([DateTime]::Parse("2017-07-05 14:00")) -Subject "Name of Event" -Attendees $Attendees
```

# Get-DefaultCalendarFolder

Return the Default Calendar Folder for a Mailbox

```
Get-DefaultCalendarFolder -MailboxName Mailbox@domain.com -AccessToken $AccessToken

```

# Get-CalendarFolder

Return the Calendar folder based on the Name entered

```
Get-CalendarFolder -MailboxName Mailbox@domain.com -AccessToken $AccessToken -FolderName 'calendar name'

```

# Get-AllCalendarFolders
Return all Calendar folders

```
Get-AllCalendarFolders -MailboxName Mailbox@domain.com -AccessToken $AccessToken

```

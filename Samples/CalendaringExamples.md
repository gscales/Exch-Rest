# Calendaring Documenation and Examples

The Exch-Rest module contains the following Calendaring Cmdlets can help peform one of many calendaring operations

#First Step is to Create an AcessToken to use for Authentication see https://github.com/gscales/Exch-Rest#authentication

# Recurrence

When you want to create a recurring Appointment you need to create a recurrence structure that can be used in one of the Calendaring cmdlets. The recurrence structure is documented https://github.com/microsoftgraph/microsoft-graph-docs/blob/master/api-reference/beta/resources/patternedrecurrence.md and https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar#PatternedRecurrence
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


Examples of Creating the Recurrence structure use the Get-Recurrence cmdlet
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

Create a Single Appointment on Secondary Calendar

```
#Creates a new all day event for Today on the default calendar
New-CalendarEventREST -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Start ([DateTime]::Parse("2017-07-05 13:00")) -End ([DateTime]::Parse("2017-07-05 14:00")) -Subject "Name of Event" -CalendarName 'Secondary calendar'
```

Create a Recurring Appointment on Secondary Calendar

```
$days = @()
$days+"monday"
$Recurrence = Get-Recurrence -PatternType weekly -PatternFirstDayOfWeek monday -PatternDaysOfWeek $days -PatternIndex first  -RangeType enddate -RangeStartDate ([DateTime]::Parse("2017-07-05 13:00")) -RangeEndDate ([DateTime]::Parse("2019-07-05 13:00"))
#Creates a new all day event for Today on the default calendar
New-CalendarEventREST -MailboxName mailbox@datarumble.com -AcessToken $AccessToken -Start ([DateTime]::Parse("2017-07-05 13:00")) -End ([DateTime]::Parse("2017-07-05 14:00")) -Subject "Monday Meeting" -Recurrence $Recurrence -CalendarName 'Secondary calendar'
```

Create a Single Appointment on a Group Calendar Calendar

Create a Recurring Appointment on a Group Calendar Calendar

Create a Single or Recurring Meeting 

# Get-DefaultCalendar

Return the Default Calendar Folder for a Mailbox

# Get-Calendar

Return the Calendar folder based on the Name entered




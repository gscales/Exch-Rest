function Get-EXRRecurrence
{
	param (
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$RecurrenceTimeZone,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[ValidateSet("daily", "weekly", "absoluteMonthly", "relativeMonthly", "absoluteYearly", " relativeYearly")]
		[string]
		$PatternType,
		
		[Parameter(Position = 3, Mandatory = $false)]
		[Int]
		$PatternInterval,
		
		[Parameter(Position = 4, Mandatory = $false)]
		[Int]
		$PatternMonth,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[Int]
		$PatternDayOfMonth,
		
		[Parameter(Position = 6, Mandatory = $true)]
		[ValidateSet("sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday")]
		[string]
		$PatternFirstDayOfWeek,
		
		[Parameter(Position = 7, Mandatory = $false)]
		[psobject]
		$PatternDaysOfWeek,
		
		[Parameter(Position = 8, Mandatory = $true)]
		[ValidateSet("first", "second", "third", "fourth", "last")]
		[string]
		$PatternIndex,
		
		[Parameter(Position = 9, Mandatory = $true)]
		[ValidateSet("noend", "enddate", "numbered")]
		[string]
		$RangeType,
		
		[Parameter(Position = 10, Mandatory = $true)]
		[datetime]
		$RangeStartDate,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[datetime]
		$RangeEndDate,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[Int]
		$RangeNumberOfOccurrences
	)
	Begin
	{
		$Recurrence = "" | Select-Object Pattern, Range, RecurrenceTimeZone
		$Pattern = "" | Select-Object Type, Interval, Month, DayOfMonth, DaysOfWeek, FirstDayOfWeek, Index
		$Range = "" | Select-Object  Type, StartDate, EndDate, NumberOfOccurrences
		if ([String]::IsNullOrEmpty($RecurrenceTimeZone))
		{
			$RecurrenceTimeZone = [TimeZoneInfo]::Local.Id
		}
		$Range.NumberOfOccurrences = 0
		$Pattern.Interval = 1
		$Pattern.Month = 0
		$Pattern.DayOfMonth = 0
		$Range.EndDate = "0001-01-01"
		$Recurrence.Pattern = $Pattern
		$Recurrence.Pattern.Type = $PatternType
		$Recurrence.Pattern.Interval = $PatternInterval
		if ($Recurrence.Pattern.Interval -eq 0)
		{
			$Recurrence.Pattern.Interval = 1
		}
		$Recurrence.Pattern.Month = $PatternMonth
		$Recurrence.Pattern.DayOfMonth = $PatternDayOfMonth
		$Recurrence.Pattern.DaysOfWeek = $PatternDaysOfWeek
		$Recurrence.Pattern.FirstDayOfWeek = $PatternFirstDayOfWeek
		$Recurrence.Pattern.Index = $PatternIndex
		$Recurrence.Range = $Range
		$Recurrence.Range.Type = $RangeType
		$Recurrence.Range.StartDate = $RangeStartDate.ToString("yyyy-MM-dd")
		if ($RangeEndDate -ne $null)
		{
			$Recurrence.Range.EndDate = $RangeEndDate.ToString("yyyy-MM-dd")
		}
		$Recurrence.Range.NumberOfOccurrences = $RangeNumberOfOccurrences
		$Recurrence.RecurrenceTimeZone = $RecurrenceTimeZone
		return, $Recurrence
	}
}

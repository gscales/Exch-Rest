function CreateFlatList
{
	param (
		[Parameter(Position = 0, Mandatory = $true)]
		[psobject]
		$EmailAddress
	)
	Begin
	{
		
		$FlatListEntry = new-object System.IO.MemoryStream
		$EntryOneOffid = "00000000812B1FA4BEA310199D6E00DD010F540200000190" + [BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes(($EmailAddress.Name + "`0"))).Replace("-", "") + [BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes(("SMTP" + "`0"))).Replace("-", "") + [BitConverter]::ToString([System.Text.UnicodeEncoding]::Unicode.GetBytes(($EmailAddress.Address + "`0"))).Replace("-", "")
		$FlatListEntryBytes = HexStringToByteArray($EntryOneOffid)
		$FlatListEntry.Write([BitConverter]::GetBytes($FlatListEntryBytes.Length), 0, 4);
		$FlatListEntry.Write($FlatListEntryBytes, 0, $FlatListEntryBytes.Length);
		$InnerLength += $FlatListEntryBytes.Length
		$Modulsval = $FlatListEntryBytes.Length % 4;
		$PadingValue = 0;
		if ($Modulsval -ne 0)
		{
			$PadingValue = 4 - $Modulsval;
			for ($AddPading = 0; $AddPading -lt $PadingValue; $AddPading++)
			{
				[Byte]$NullValue = 00;
				$FlatlistStream.Write($NullValue, 0, 1);
			}
		}
		$FlatListEntry.Position = 0
		$FlatListEntryBytes = $FlatListEntry.ToArray()
		$FlatList = new-object System.IO.MemoryStream
		$FlatList.Write([BitConverter]::GetBytes(1), 0, 4);
		$FlatList.Write([BitConverter]::GetBytes($FlatListEntryBytes.Length), 0, 4);
		$FlatList.Write($FlatListEntryBytes, 0, $FlatListEntryBytes.Length);
		$FlatList.Position = 0
		return, $FlatList.ToArray()
	}
}

Connect-EXRMailbox -MailboxName "gscales@datarumble.com"
$WebhookAddress = "https://outlook.office.com/webhook/....."
$GroupName = "A Team"
$Group = Get-EXRModernGroups -GroupName $GroupName
$Members = Get-EXRGroupMembers -GroupId $Group.id
$OOFMailboxCol = @()
$mtHash = @{}
foreach($Member in $Members){
    if(![String]::IsNullOrEmpty($Member.mail)){
        $OOFMailboxCol += $Member.mail
        $mtHash.Add($Member.mail,$Member.displayName)
    }  
}
$Mailtips = Get-EXRMailTips -Mailboxes $OOFMailboxCol -tips "automaticReplies"
$FactsColl = ($Mailtips | select @{Name='EmailAddress';Expression={$mtHash[$_.emailAddress.address]}},@{Name='OOF Message';Expression={$val =($_.automaticReplies.message -replace '<[^>]+>','').Trim();if([String]::IsNullOrEmpty($val)){"In"}else{"Out : " + $val}}})
Invoke-WebRequest -Uri $WebhookAddress -Method Post -ContentType 'Application/Json' -Body (New-EXRMessageCard -Facts $FactsColl -Summary "Team OOF Status" -Title "Team OOF Status")
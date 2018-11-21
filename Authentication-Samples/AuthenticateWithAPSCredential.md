    $PSCredential = Get-Credential
	Connect-Mailbox -MailboxName Bal@domain.com -Credential $PSCredentials -ResourceURL graph.microsoft.com -ClientId "d3590ed6-52b3-4102-aeff-aad2292ab01c"
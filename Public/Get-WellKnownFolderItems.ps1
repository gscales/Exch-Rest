function Get-WellKnownFolderItems{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [string]$WellKnownFolder,
        [Parameter(Position=4, Mandatory=$false)] [switch]$ReturnSize,
        [Parameter(Position=5, Mandatory=$false)] [string]$SelectProperties,
        [Parameter(Position=6, Mandatory=$false)] [string]$Filter,
        [Parameter(Position=7, Mandatory=$false)] [string]$Top,
        [Parameter(Position=8, Mandatory=$false)] [string]$OrderBy,
        [Parameter(Position=9, Mandatory=$false)] [bool]$TopOnly
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        if(![String]::IsNullorEmpty($Filter)){
            $Filter = "`&`$filter=" + $Filter
        }
        if(![String]::IsNullorEmpty($Orderby)){
            $OrderBy = "`&`$OrderBy=" + $OrderBy
        }
        $TopValue = "1000"    
        if(![String]::IsNullorEmpty($Top)){
            $TopValue = $Top
        }      
        if([String]::IsNullorEmpty($SelectProperties)){
            $SelectProperties = "`$select=ReceivedDateTime,Sender,Subject,IsRead"
        }
        else{
            $SelectProperties = "`$select=" + $SelectProperties
        }
        if($WellKnownFolder -ne $null)
        {
            $HttpClient =  Get-HTTPClient($MailboxName)
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL =  $EndPoint + "('" + $MailboxName + "')/MailFolders/" + $WellKnownFolder + "/messages/?" +  $SelectProperties + "`&`$Top=" + $TopValue + $Filter + $OrderBy
            
            if($ReturnSize.IsPresent){
                $PropName = "PropertyId"
                if($AccessToken.resource -eq "https://graph.microsoft.com"){
                        $PropName = "Id"
                }
                $RequestURL =  $EndPoint + "('" + $MailboxName + "')/MailFolders/" + $WellKnownFolder + "/messages/?`$select=ReceivedDateTime,Sender,Subject,IsRead`&`$Top=" + $TopValue + "`&`$expand=SingleValueExtendedProperties(`$filter=$PropName%20eq%20'Integer%200x0E08')" + $Filter + $OrderBy
            }
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
                foreach ($Message in $JSONOutput.Value) {
                    Add-Member -InputObject $Message -NotePropertyName ItemRESTURI -NotePropertyValue ($Folder.FolderRestURI + "/messages('" + $Message.Id + "')")
                    Write-Output $Message
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL) -band (!$TopOnly))     
       } 
   

    }
}

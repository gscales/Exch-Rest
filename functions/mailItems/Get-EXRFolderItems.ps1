function Get-EXRFolderItems{
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$false)] [psobject]$AccessToken,
        [Parameter(Position=2, Mandatory=$false)] [psobject]$Folder,
        [Parameter(Position=4, Mandatory=$false)] [switch]$ReturnSize,
        [Parameter(Position=5, Mandatory=$false)] [string]$SelectProperties,
        [Parameter(Position=6, Mandatory=$false)] [string]$Filter,
        [Parameter(Position=7, Mandatory=$false)] [string]$Top,
        [Parameter(Position=8, Mandatory=$false)] [string]$OrderBy,
        [Parameter(Position=9, Mandatory=$false)] [bool]$TopOnly,
        [Parameter(Position=10, Mandatory=$false)] [PSCustomObject]$PropList,
        [Parameter(Position=11, Mandatory=$false)] [string]$Search,
        [Parameter(Position=12, Mandatory=$false)] [switch]$TrackStatus
    )
    Begin{
        if($AccessToken -eq $null)
        {
              $AccessToken = Get-AccessToken -MailboxName $MailboxName          
        }  
        if(![String]::IsNullorEmpty($Filter)){
            $Filter = "`&`$filter=" + $Filter
        }
        if(![String]::IsNullorEmpty($Search)){
            $Search = "`&`$Search=" + $Search
        }
        if(![String]::IsNullorEmpty($Orderby)){
            $OrderBy = "`&`$OrderBy=" + $OrderBy
        }
        $TopValue = "1000"    
        if(![String]::IsNullorEmpty($Top)){
            $TopValue = $Top
        }      
        if([String]::IsNullorEmpty($SelectProperties)){
            $SelectProperties = "`$select=ReceivedDateTime,Sender,Subject,IsRead,hasAttachments"
        }
        else{
            $SelectProperties = "`$select=" + $SelectProperties
        }
        if($Folder -ne $null)
        {
            $HttpClient =  Get-HTTPClient -MailboxName $MailboxName
            $EndPoint =  Get-EndPoint -AccessToken $AccessToken -Segment "users"
            $RequestURL =  $EndPoint + "('" + $MailboxName + "')/MailFolders('" + $Folder.Id + "')/messages/?" +  $SelectProperties + "`&`$Top=" + $TopValue 
            $folderURI =  $EndPoint + "('" + $MailboxName + "')/MailFolders('" + $Folder.Id + "')"
            if($ReturnSize.IsPresent){
                if($PropList -eq $null){
                    $PropList = @()
                    $PidTagMessageSize = Get-TaggedProperty -DataType "Integer" -Id "0x0E08"  
                    $PropList += $PidTagMessageSize
                }
            }
            if($PropList -ne $null){
               $Props = Get-ExtendedPropList -PropertyList $PropList -AccessToken $AccessToken
               $RequestURL += "`&`$expand=SingleValueExtendedProperties(`$filter=" + $Props + ")"
            }
            $RequestURL += $Search + $Filter + $OrderBy
            do{
                $JSONOutput = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -TrackStatus $TrackStatus.IsPresent
                foreach ($Message in $JSONOutput.Value) {
                    Add-Member -InputObject $Message -NotePropertyName ItemRESTURI -NotePropertyValue ($folderURI  + "/messages('" + $Message.Id + "')")
                    Invoke-EXRParseExtendedProperties -Item $Message
                    Write-Output $Message
                }           
                $RequestURL = $JSONOutput.'@odata.nextLink'
            }while(![String]::IsNullOrEmpty($RequestURL) -band (!$TopOnly))     
       } 
   

    }
}

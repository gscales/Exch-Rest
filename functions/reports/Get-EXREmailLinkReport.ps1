
function Get-EXREmailLinkReport {
    [CmdletBinding()] 
    param (
        [Parameter(Position = 0, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$WellKnownFolder,
        [Parameter(Position = 2, Mandatory = $false)] [psobject]$Folder,
        [Parameter(Position = 3, Mandatory = $false)] [String]$FolderPath,
        [Parameter(Position = 3, Mandatory = $false)] [String]$MessageCount

    )
	
    process {
        $Props = @()
        $PR_BODY_HTML = Get-EXRTaggedProperty -DataType Binary -Id 0x1013
        $Props += $PR_BODY_HTML
        $Messages = Get-EXRWellKnownFolderItems -MailboxName $MailboxName -AccessToken $AccessToken -WellKnownFolder $WellKnownFolder -Folder $Folder -FolderPath $FolderPath -MessageCount $MessageCount -BatchReturnItems -SelectProperties Subject -PropList $Props
        $HrefPaths = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $Domains = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $Hrefs = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $BaseHrefs = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $Images = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        $ImageDomains = New-Object 'system.collections.generic.dictionary[string,PsObject]'
        foreach ($Message in $Messages) {
            Invoke-EXRParseEmailBodyLinks -Item $Message -UseExtendedProperty
            $MessageDomains = @{}      
            $MessageDomainsHrefPaths = @{}      
            $MessageDomainsHrefs = @{}  
            $MessageDomainsImages = @{}    
            $MessageImageDomains = @{} 
            if ($Message.ParsedLinks.HasBaseURL) {
                if (!$BaseHrefs.ContainsKey($Message.ParsedLinks.ParsedBaseURL)) {
                    $values = "" | Select BaseHref, Count
                    $values.BaseHref = $Message.ParsedLinks.ParsedBaseURL
                    $values.Count = 1
                    $BaseHrefs.add($Message.ParsedLinks.ParsedBaseURL, $values)
                }
                else {
                    $BaseHrefs[$Message.ParsedLinks.ParsedBaseURL].Count++
                }
            }
            foreach ($link in $Message.ParsedLinks.Links) {
                if (![String]::IsNullOrEmpty($link.DnsSafeHost)) {
                    if (![String]::IsNullOrEmpty($link.AbsolutePath)) {
                        $fpath = $link.host + "/" + $link.AbsolutePath
                        if (!$HrefPaths.ContainsKey($fpath)) {
                            $Counts = "" | select HrefPath, MessageCount, LinkCount
                            $Counts.HrefPath = $fpath
                            $Counts.MessageCount = 0
                            $Counts.LinkCount = 1
                            $HrefPaths.Add($fpath, $Counts)
                        }
                        else {
                            $HrefPaths[$fpath].LinkCount++
                        }
                        if (!$MessageDomainsHrefPaths.Contains($fpath)) {
                            $MessageDomainsHrefPaths.Add($fpath, 1)
                            $HrefPaths[$fpath].MessageCount++
                        }
                    }
                    if (![String]::IsNullOrEmpty($link.AbsoluteUri)) {
                        if (!$Hrefs.ContainsKey($link.AbsoluteUri)) {
                            $Counts = "" | select Href, MessageCount, LinkCount
                            $Counts.Href = $link.AbsoluteUri
                            $Counts.MessageCount = 0
                            $Counts.LinkCount = 1
                            $Hrefs.Add($link.AbsoluteUri, $Counts)
                        }
                        else {
                            $Hrefs[$link.AbsoluteUri].LinkCount++
                        }
                        if (!$MessageDomainsHrefs.Contains($link.AbsoluteUri)) {
                            $MessageDomainsHrefs.Add($link.AbsoluteUri, 1)
                            $Hrefs[$link.AbsoluteUri].MessageCount++
                        }
                    }
                    if (!$Domains.ContainsKey($link.DnsSafeHost)) {
                        $Counts = "" | select HostName, MessageCount, LinkCount
                        $Counts.MessageCount = 0
                        $Counts.LinkCount = 1
                        $Counts.HostName = $link.DnsSafeHost
                        $Domains.Add($link.DnsSafeHost, $Counts)
                    }
                    else {
                        $Domains[$link.DnsSafeHost].LinkCount++
                    }
                    if (!$MessageDomains.Contains($link.DnsSafeHost)) {
                        $MessageDomains.Add($link.DnsSafeHost, 1)
                        $Domains[$link.DnsSafeHost].MessageCount++
                    }
                }
            }
            foreach ($link in $Message.ParsedLinks.Images) {
                if (![String]::IsNullOrEmpty($link.AbsoluteUri)) {
                    if (!$Images.ContainsKey($link.AbsoluteUri)) {
                        $Counts = "" | select Src, MessageCount, LinkCount
                        $Counts.Src = $link.AbsoluteUri
                        $Counts.MessageCount = 0
                        $Counts.LinkCount = 1
                        $Images.Add($link.AbsoluteUri, $Counts)
                    }
                    else {
                        $Images[$link.AbsoluteUri].LinkCount++
                    }
                    if (!$MessageDomainsImages.Contains($link.AbsoluteUri)) {
                        $MessageDomainsImages.Add($link.AbsoluteUri, 1)
                        $Images[$link.AbsoluteUri].MessageCount++
                    }
                    if (![String]::IsNullOrEmpty($link.DnsSafeHost)) {
                        if (!$ImageDomains.ContainsKey($link.DnsSafeHost)) {
                            $Counts = "" | select HostName, MessageCount, LinkCount
                            $Counts.MessageCount = 0
                            $Counts.LinkCount = 1
                            $Counts.HostName = $link.DnsSafeHost
                            $ImageDomains.Add($link.DnsSafeHost, $Counts)
                        }
                        else {
                            $ImageDomains[$link.DnsSafeHost].LinkCount++
                        }
                        if (!$MessageImageDomains.Contains($link.DnsSafeHost)) {
                            $MessageImageDomains.Add($link.DnsSafeHost, 1)
                            $ImageDomains[$link.DnsSafeHost].MessageCount++
                        }
                    }
                }   
            }
        }
        $report = "" | Select Domains, hrefPaths, Hrefs, BaseHrefs, Images, ImageDomains
        $report.Domains = [Collections.Generic.List[PsObject]]$Domains.Values
        $report.HrefPaths = [Collections.Generic.List[PsObject]]$HrefPaths.Values
        $report.Hrefs = [Collections.Generic.List[PsObject]]$Hrefs.Values       
        $report.BaseHrefs = [Collections.Generic.List[PsObject]]$BaseHrefs.Values 
        $report.Images = [Collections.Generic.List[PsObject]]$Images.Values
        $report.ImageDomains = [Collections.Generic.List[PsObject]]$ImageDomains.Values
        return, $report
    }

}
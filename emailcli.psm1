# emailcli.psm1
# Import with: Import-Module .\emailcli.psm1
# Tom Villani, Ph.D.
#
# A simple powershell toolbox to read emails from outlook into the command line

Add-Type -AssemblyName System.Web


# Global variable to cache messages for the session
$script:CachedMessages = @()
# Track the parameters used for the last inbox query so interactive mode can reuse them
$script:LastInboxParameters = @{}
# Cache for archived email search results
$script:CachedArchiveMessages = @()
# Global configuration loaded from ~/.emailcli.json
$script:EmailCliConfig = $null

function Get-CachedMessages {
    <#
    .SYNOPSIS
    Returns the currently cached messages from the last Get-OutlookInbox call

    .DESCRIPTION
    Helper function to access cached messages for debugging or external use
    #>
    return $script:CachedMessages
}

function Get-EmailCliConfig {
    <#
    .SYNOPSIS
    Loads configuration from ~/.emailcli.json if available.

    .DESCRIPTION
    Attempts to load and parse the JSON configuration file.
    Returns a hashtable with configuration sections or empty hashtable if unavailable.
    Caches the result in $script:EmailCliConfig.
    #>
    [CmdletBinding()]
    param()

    # Return cached config if already loaded
    if ($null -ne $script:EmailCliConfig) {
        return $script:EmailCliConfig
    }

    # Determine config file path
    $configPath = Join-Path $env:USERPROFILE ".emailcli.json"

    # Check if file exists
    if (-not (Test-Path $configPath)) {
        Write-Verbose "Config file not found at $configPath"
        $script:EmailCliConfig = @{}
        return $script:EmailCliConfig
    }

    try {
        $configContent = Get-Content -Path $configPath -Raw -Encoding UTF8
        $config = ConvertFrom-Json -InputObject $configContent

        # Convert PSCustomObject to hashtable for easier access
        $script:EmailCliConfig = @{}
        foreach ($section in $config.PSObject.Properties) {
            $script:EmailCliConfig[$section.Name] = @{}
            foreach ($prop in $section.Value.PSObject.Properties) {
                $script:EmailCliConfig[$section.Name][$prop.Name] = $prop.Value
            }
        }

        Write-Verbose "Loaded configuration from $configPath"
        return $script:EmailCliConfig

    } catch {
        Write-Warning "Failed to load configuration from ${configPath}: $_"
        $script:EmailCliConfig = @{}
        return $script:EmailCliConfig
    }
}

function Get-ConfigValue {
    <#
    .SYNOPSIS
    Helper to retrieve a configuration value from a specific section.

    .PARAMETER Section
    The configuration section name (e.g., "Calendar", "Email")

    .PARAMETER Key
    The parameter name (e.g., "Days", "DaysBack")

    .PARAMETER Default
    Default value if config key not found
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Section,
        [Parameter(Mandatory=$true)]
        [string]$Key,
        $Default = $null
    )

    $config = Get-EmailCliConfig

    if ($config.ContainsKey($Section) -and $config[$Section].ContainsKey($Key)) {
        return $config[$Section][$Key]
    }

    return $Default
}

function Reset-EmailCliConfig {
    <#
    .SYNOPSIS
    Clears the cached configuration, forcing reload on next access.
    #>
    [CmdletBinding()]
    param()

    $script:EmailCliConfig = $null
    Write-Host "Configuration cache cleared." -ForegroundColor Green
}

function Set-ConfigValue {
    <#
    .SYNOPSIS
    Sets a configuration value and saves it to the config file.

    .DESCRIPTION
    Updates a configuration value in memory and persists it to ~/.emailcli.json.
    Creates the config file if it doesn't exist.

    .PARAMETER Section
    The configuration section name (e.g., "Calendar", "Email", "Interactive")

    .PARAMETER Key
    The parameter name (e.g., "Days", "DaysBack", "PreviewLength")

    .PARAMETER Value
    The value to set (can be string, number, or boolean)

    .EXAMPLE
    Set-ConfigValue -Section "Interactive" -Key "PreviewLength" -Value 800
    Sets the interactive preview length to 800 characters.

    .EXAMPLE
    Set-ConfigValue -Section "Email" -Key "Compact" -Value $true
    Enables compact mode by default for email viewing.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Section,
        [Parameter(Mandatory=$true)]
        [string]$Key,
        [Parameter(Mandatory=$true)]
        $Value
    )

    try {
        $configPath = Join-Path $env:USERPROFILE ".emailcli.json"

        # Load existing config or create new one
        if (Test-Path $configPath) {
            $configContent = Get-Content -Path $configPath -Raw -Encoding UTF8
            $config = ConvertFrom-Json -InputObject $configContent
        } else {
            # Create empty config structure
            $config = [PSCustomObject]@{}
        }

        # Ensure section exists
        if (-not ($config.PSObject.Properties.Name -contains $Section)) {
            $config | Add-Member -MemberType NoteProperty -Name $Section -Value ([PSCustomObject]@{})
        }

        # Set or update the value
        if ($config.$Section.PSObject.Properties.Name -contains $Key) {
            $config.$Section.$Key = $Value
        } else {
            $config.$Section | Add-Member -MemberType NoteProperty -Name $Key -Value $Value
        }

        # Save back to file
        $configJson = $config | ConvertTo-Json -Depth 10
        $configJson | Out-File -FilePath $configPath -Encoding UTF8 -Force

        # Clear the cache so the new value is picked up
        $script:EmailCliConfig = $null

        Write-Host "Configuration updated: $Section.$Key = $Value" -ForegroundColor Green
        Write-Host "Saved to: $configPath" -ForegroundColor Gray

    } catch {
        Write-Error "Failed to set configuration value: $_"
    }
}

# Load configuration file if available
Get-EmailCliConfig | Out-Null

function Clean-EmailBody {
    <#
    .SYNOPSIS
    Cleans up email body text for better CLI readability
    #>
    param(
        [string]$Body,
        [switch]$KeepLinks
    )

    # Decode SafeLinks URLs using a proper regex callback
    $pattern = 'https://[a-z0-9]+\.safelinks\.protection\.outlook\.com/\?url=([^&]+)&[^>\s]*'
    $Body = [regex]::Replace($Body, $pattern, {
        param($match)
        $encodedUrl = $match.Groups[1].Value
        # URL decode the actual destination
        [System.Web.HttpUtility]::UrlDecode($encodedUrl)
    })

    # Handle SendGrid and other tracking URLs by extracting the upn parameter or simplifying
    $Body = [regex]::Replace($Body, 'https://[^/]+\.ct\.sendgrid\.net/ls/click\?upn=[^>\s]+', '[SendGrid Link]')

    # Handle other common tracking domains with very long URLs
    $Body = [regex]::Replace($Body, 'https://[^/]+\.(?:mailchimp|constantcontact|aweber)\.com/[^>\s]{100,}', '[Tracking Link]')

    # Truncate any remaining URLs longer than 80 characters to just show domain
    if (-not $Body.Contains("docusign.net/Signing") -and -not $KeepLinks){
    	$Body = [regex]::Replace($Body, '<?(https?://([^/>\s]+)(/[^>\s]{60,}))>?', {
    		 param($match)
    		 $domain = $match.Groups[2].Value
    		 "<$domain/...>"
    	})	
    }
    

    # Remove NJII external email warning boilerplate
    $Body = $Body -replace '\[EXTERNAL EMAIL\]\s*CAUTION:\s*This email originated from outside of NJII\.\s*Do not click on links or open attachments unless you recognize the sender and know the content is safe\.', ''

    # Clean up common HTML entities if present
    $Body = $Body -replace '&nbsp;', ' '
    $Body = $Body -replace '&amp;', '&'
    $Body = $Body -replace '&lt;', '<'
    $Body = $Body -replace '&gt;', '>'
    $Body = $Body -replace '&quot;', '"'
    $Body = $Body -replace '&#39;', "'"

    # Remove trailing whitespace from each line
    $Body = ($Body -split "`n" | ForEach-Object { $_.TrimEnd() }) -join "`n"

    # Remove lines that contain only punctuation or whitespace (like single commas, periods, etc.)
    $Body = ($Body -split "`n" | Where-Object { $_ -notmatch '^\s*[,;.\-_|]\s*$' }) -join "`n"

    # Remove excessive blank lines (more than 2 consecutive newlines)
    # Do this after trimming so lines with only whitespace are now truly empty
    $Body = $Body -replace '(\r?\n\s*){3,}', "`n`n"

    return $Body.Trim()
}

function Split-EmailThread {
    <#
    .SYNOPSIS
    Splits an email body into the newest message content and the quoted history.

    .DESCRIPTION
    Detects common reply separators (On ... wrote, From/Sent headers, Original Message markers, Gmail quote divs)
    to isolate the most recent reply from older thread content.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Body
    )

    if ([string]::IsNullOrWhiteSpace($Body)) {
        return [pscustomobject]@{
            Latest    = ""
            History   = ""
            HasHistory = $false
        }
    }

    $normalized = $Body
    $gmailIndex = $normalized.IndexOf("gmail_quote")
    if ($gmailIndex -gt -1) {
        $splitText = '<div class="gmail_quote">'
        $pos = $normalized.IndexOf($splitText)
        if ($pos -gt -1) {
            $latest = $normalized.Substring(0, $pos)
            $history = $normalized.Substring($pos)
            return [pscustomobject]@{
                Latest = $latest.Trim()
                History = $history.Trim()
                HasHistory = $true
            }
        }
    }

    $patterns = @(
        '(?m)^\s*On .+ wrote:$',
        '(?m)^\s*From:\s.+\r?\n\s*Sent:',
        '(?m)^\s*From:\s.+\r?\n\s*To:',
        '(?m)^\s*-----Original Message-----',
        '(?m)^\s*Begin forwarded message:',
        '(?m)^\s*__+\s*Original Message\s*__+',
        '(?m)^\s*>+\s?On .+ wrote:$'
    )

    $earliestIndex = $null
    foreach ($pattern in $patterns) {
        $match = [regex]::Match($normalized, $pattern)
        if ($match.Success) {
            if ($earliestIndex -eq $null -or $match.Index -lt $earliestIndex) {
                $earliestIndex = $match.Index
            }
        }
    }

    if ($earliestIndex -eq $null -or $earliestIndex -le 0) {
        return [pscustomobject]@{
            Latest = $Body.Trim()
            History = ""
            HasHistory = $false
        }
    }

    $latestPart = $normalized.Substring(0, $earliestIndex).Trim()
    $historyPart = $normalized.Substring($earliestIndex).Trim()

    if ([string]::IsNullOrWhiteSpace($latestPart)) {
        return [pscustomobject]@{
            Latest = $Body.Trim()
            History = ""
            HasHistory = $false
        }
    }

    return [pscustomobject]@{
        Latest = $latestPart
        History = $historyPart
        HasHistory = -not [string]::IsNullOrWhiteSpace($historyPart)
    }
}

function Get-CategoryColorMap {
    <#
    .SYNOPSIS
    Builds a hashtable mapping category names to their console colors

    .DESCRIPTION
    Queries Outlook categories and maps each category name to an appropriate
    PowerShell console color based on the category's Outlook color value.

    .RETURNS
    Hashtable - Keys are category names, values are PowerShell console color names
    #>
    [CmdletBinding()]
    param()

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $categories = $namespace.Categories

        $colorMap = @{}

        foreach ($cat in $categories) {
            # Map Outlook color codes to PowerShell console colors
            $consoleColor = switch ($cat.Color) {
                0  { "White" }        # None
                1  { "Red" }          # Red
                2  { "DarkYellow" }   # Orange
                3  { "Yellow" }       # Peach
                4  { "Yellow" }       # Yellow
                5  { "Green" }        # Green
                6  { "Cyan" }         # Teal
                7  { "DarkYellow" }   # Olive
                8  { "Blue" }         # Blue
                9  { "Magenta" }      # Purple
                10 { "DarkRed" }      # Maroon
                11 { "Gray" }         # Steel
                12 { "DarkGray" }     # DarkSteel
                13 { "Gray" }         # Gray
                14 { "DarkGray" }     # DarkGray
                15 { "White" }        # Black (show as white on dark background)
                16 { "DarkRed" }      # DarkRed
                17 { "DarkYellow" }   # DarkOrange
                18 { "DarkYellow" }   # DarkPeach
                19 { "DarkYellow" }   # DarkYellow
                20 { "DarkGreen" }    # DarkGreen
                21 { "DarkCyan" }     # DarkTeal
                22 { "DarkYellow" }   # DarkOlive
                23 { "DarkBlue" }     # DarkBlue
                24 { "DarkMagenta" }  # DarkPurple
                25 { "DarkRed" }      # DarkMaroon
                default { "White" }   # Unknown
            }

            $colorMap[$cat.Name] = $consoleColor
        }

        return $colorMap
    }
    catch {
        # Return empty map on error
        return @{}
    }
}

# Cache the category color map for the session
$script:CategoryColorMap = $null

function Get-CachedCategoryColorMap {
    <#
    .SYNOPSIS
    Returns a cached category color map, building it if necessary
    #>
    if ($null -eq $script:CategoryColorMap) {
        $script:CategoryColorMap = Get-CategoryColorMap
    }
    return $script:CategoryColorMap
}

function Reset-CategoryColorCache {
    <#
    .SYNOPSIS
    Clears the cached category color map, forcing a refresh on next access

    .DESCRIPTION
    Use this if you have added new categories to Outlook during your session
    and want them to appear with proper colors in subsequent commands
    #>
    $script:CategoryColorMap = $null
    Write-Host "Category color cache cleared. Colors will refresh on next use." -ForegroundColor Green
}

function Write-ColoredCategories {
    <#
    .SYNOPSIS
    Writes category names in their associated colors

    .PARAMETER Categories
    Semicolon-separated string of category names

    .PARAMETER Prefix
    Optional prefix to write before categories (e.g., " [")

    .PARAMETER Suffix
    Optional suffix to write after categories (e.g., "]")

    .PARAMETER NoNewline
    If specified, does not add a newline after output
    #>
    [CmdletBinding()]
    param(
        [string]$Categories,
        [string]$Prefix = "",
        [string]$Suffix = "",
        [switch]$NoNewline
    )

    if ([string]::IsNullOrWhiteSpace($Categories)) {
        return
    }

    $colorMap = Get-CachedCategoryColorMap
    # Force array type to prevent single-element strings from being treated as character arrays
    $categoryList = @($Categories -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    if ($Prefix) {
        Write-Host $Prefix -NoNewline -ForegroundColor Gray
    }

    for ($i = 0; $i -lt $categoryList.Count; $i++) {
        $catName = $categoryList[$i]
        $color = if ($colorMap.ContainsKey($catName)) { $colorMap[$catName] } else { "White" }

        Write-Host $catName -NoNewline -ForegroundColor $color

        if ($i -lt $categoryList.Count - 1) {
            Write-Host ", " -NoNewline -ForegroundColor Gray
        }
    }

    if ($Suffix) {
        Write-Host $Suffix -NoNewline -ForegroundColor Gray
    }

    if (-not $NoNewline) {
        Write-Host ""
    }
}

function Get-SmtpAddress {
    <#
    .SYNOPSIS
    Resolves an email sender's SMTP address from a mail item

    .DESCRIPTION
    Extracts the sender's SMTP email address from an Outlook mail item.
    If the sender's address is in Exchange X.500 Distinguished Name format,
    this function attempts to resolve it to the actual SMTP address using
    the Outlook Object Model's GetExchangeUser() method.

    .PARAMETER MailItem
    The Outlook mail item object from which to extract the sender's SMTP address

    .RETURNS
    String - The sender's SMTP email address, or the original address if resolution fails

    .NOTES
    Exchange X.500 Distinguished Names have the format:
    /O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP.../CN=RECIPIENTS/CN=...

    This function will attempt to resolve these to proper SMTP addresses like:
    user@example.com
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.__ComObject]$MailItem
    )

    try {
        $address = $MailItem.SenderEmailAddress

        # Check if the address is in X.500 Distinguished Name format
        if ($address -like "/O=*" -or $address -like "/o=*") {
            try {
                # Get the sender's AddressEntry object
                $sender = $MailItem.Sender

                if ($null -ne $sender) {
                    # Try to get the Exchange user object
                    $exchUser = $sender.GetExchangeUser()

                    if ($null -ne $exchUser -and ![string]::IsNullOrWhiteSpace($exchUser.PrimarySmtpAddress)) {
                        # Successfully resolved to SMTP address
                        return $exchUser.PrimarySmtpAddress
                    }
                }
            }
            catch {
                # If resolution fails, fall back to original address
                # This is expected for some edge cases and is not an error
            }
        }

        # Return original address if not X.500 or if resolution failed
        return $address
    }
    catch {
        # If anything goes wrong, return a safe default
        return $MailItem.SenderEmailAddress
    }
}


function Get-OutlookInbox {
    <#
    .SYNOPSIS
    Lists emails from your Outlook inbox
    
    .PARAMETER IncludeRead
    Include read messages (default: false)
    
    .PARAMETER DaysBack
    Only show emails from the last N days (default: 7)
    
    .PARAMETER All
    Show all messages regardless of read status, replied status, or date

    .PARAMETER Replied
    Include messages you have replied to (by default, replied messages are hidden)

    .PARAMETER Flagged
    Show only flagged/marked messages
    
    .PARAMETER Category
    Filter by category/tag name (e.g., "Red Category", "Important")
    
    .PARAMETER From
    Filter by sender name or email (partial match)
    
    .PARAMETER HasAttachments
    Show only messages with attachments
    
    .PARAMETER Importance
    Filter by importance level: High, Normal, or Low
    
    .PARAMETER Subject
    Filter by subject line (partial match)
    
    .PARAMETER Limit
    Maximum number of messages to display (alias: -n)
    
    .PARAMETER Before
    Include messages older than this many days ago. Combine with -DaysBack to create a window.
    
    .PARAMETER StartDate
    Earliest ReceivedTime to include (inclusive). Cannot be combined with -DaysBack/-Before.
    
    .PARAMETER EndDate
    Latest ReceivedTime to include (inclusive). Cannot be combined with -DaysBack/-Before.
    
    .PARAMETER Plain
    Output plain text without colors or formatting (suitable for piping)

    .PARAMETER Full
    Output full email content for all messages in plain text mode (automatically enables -Plain)

    .PARAMETER Offset
    The offset to start at when displaying messages (0 is newest)
    #>
    [CmdletBinding()]
    param(
        [Alias("read")]
        [switch]$IncludeRead,
        [Alias("days")]
        [int]$DaysBack = 7,
        [switch]$All,
        [switch]$Replied,
        [switch]$Flagged,
        [string]$Category,
        [string]$From,
        [switch]$HasAttachments,
        [ValidateSet("High", "Normal", "Low")]
        [string]$Importance,
        [string]$Subject,
        [Alias("n")]
        [int]$Limit = 0,
        [int]$Before = 0,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [switch]$Plain,
        [switch]$Full,
        [int]$Offset = 0,
        [switch]$Oldest,
        [switch]$Compact,
        [switch]$Silent,
        [switch]$KeepLinks
    )

    # Apply configuration defaults if parameters not explicitly provided
    if (-not $PSBoundParameters.ContainsKey('DaysBack')) {
        $configDaysBack = Get-ConfigValue -Section "Email" -Key "DaysBack" -Default 7
        $DaysBack = $configDaysBack
    }

    if (-not $PSBoundParameters.ContainsKey('IncludeRead')) {
        $configIncludeRead = Get-ConfigValue -Section "Email" -Key "IncludeRead" -Default $false
        if ($configIncludeRead -eq $true) {
            $IncludeRead = $true
        }
    }

    if (-not $PSBoundParameters.ContainsKey('Limit')) {
        $configLimit = Get-ConfigValue -Section "Email" -Key "Limit" -Default 0
        if ($configLimit -gt 0) {
            $Limit = $configLimit
        }
    }

    if (-not $PSBoundParameters.ContainsKey('Compact')) {
        $configCompact = Get-ConfigValue -Section "Email" -Key "Compact" -Default $false
        if ($configCompact -eq $true) {
            $Compact = $true
        }
    }

    if (-not $PSBoundParameters.ContainsKey('KeepLinks')) {
        $configKeepLinks = Get-ConfigValue -Section "Email" -Key "KeepLinks" -Default $false
        if ($configKeepLinks -eq $true) {
            $KeepLinks = $true
        }
    }

    if (-not $PSBoundParameters.ContainsKey('Oldest')) {
        $configOldest = Get-ConfigValue -Section "Email" -Key "Oldest" -Default $false
        if ($configOldest -eq $true) {
            $Oldest = $true
        }
    }

    try {
        # Auto-enable Plain mode when Full is specified
        if ($Full) {
            $Plain = $true
        }

        if ($Before -lt 0) {
            throw "Before must be zero or a positive number of days."
        }

        $usingAbsoluteDates = $PSBoundParameters.ContainsKey("StartDate") -or $PSBoundParameters.ContainsKey("EndDate")

        if ($usingAbsoluteDates -and $PSBoundParameters.ContainsKey("DaysBack")) {
            throw "StartDate/EndDate cannot be combined with DaysBack."
        }

        if ($usingAbsoluteDates -and $PSBoundParameters.ContainsKey("Before")) {
            throw "StartDate/EndDate cannot be combined with Before."
        }

        if ($StartDate -and $EndDate -and $StartDate -gt $EndDate) {
            throw "StartDate must be earlier than or equal to EndDate."
        }

        # Snapshot of the parameters used for this inbox view
        $paramSnapshot = @{}
        if (-not $usingAbsoluteDates) {
            $paramSnapshot['DaysBack'] = $DaysBack
        }
        if ($IncludeRead)    { $paramSnapshot['IncludeRead']   = $true }
        if ($All)            { $paramSnapshot['All']           = $true }
        if ($Replied)        { $paramSnapshot['Replied']       = $true }
        if ($Flagged)        { $paramSnapshot['Flagged']       = $true }
        if ($Category)       { $paramSnapshot['Category']      = $Category }
        if ($From)           { $paramSnapshot['From']          = $From }
        if ($HasAttachments) { $paramSnapshot['HasAttachments']= $true }
        if ($Importance)     { $paramSnapshot['Importance']    = $Importance }
        if ($Subject)        { $paramSnapshot['Subject']       = $Subject }
        if ($Limit -gt 0)    { $paramSnapshot['Limit']         = $Limit }
        if ($Offset -gt 0)   { $paramSnapshot['Offset']        = $Offset }
        if ($Oldest)         { $paramSnapshot['Oldest']        = $true }
        if ($KeepLinks)      { $paramSnapshot['KeepLinks']     = $true }
        if ($Compact)        { $paramSnapshot['Compact']       = $true }
        if ($Silent)         { $paramSnapshot['Silent']        = $true }
        if ($Before -gt 0)   { $paramSnapshot['Before']        = $Before }
        if ($StartDate)      { $paramSnapshot['StartDate']     = $StartDate }
        if ($EndDate)        { $paramSnapshot['EndDate']       = $EndDate }

        $lowerBound = $null
        $upperBound = $null
        if (!$All) {
            if ($usingAbsoluteDates) {
                if ($StartDate) {
                    $lowerBound = [datetime]$StartDate
                }
                if ($EndDate) {
                    $upperBound = [datetime]$EndDate
                    if ($upperBound.TimeOfDay -eq [TimeSpan]::Zero) {
                        $upperBound = $upperBound.AddDays(1).AddTicks(-1)
                    }
                }
            } else {
                $referenceNow = Get-Date
                if ($DaysBack -gt 0) {
                    $lowerBound = $referenceNow.AddDays(-$DaysBack)
                }
                if ($Before -gt 0) {
                    $upperBound = $referenceNow.AddDays(-$Before)
                }
            }

            if ($lowerBound -and $upperBound -and $lowerBound -gt $upperBound) {
                throw "The computed date range is invalid (lower bound occurs after upper bound). Adjust -DaysBack/-Before or StartDate/EndDate."
            }
        }

        # Connect to Outlook
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6) # 6 = Inbox
        
        # Get messages
        $messages = $inbox.Items
        $messages.Sort("[ReceivedTime]", -not $Oldest) # Sort by received time, descending or ascending depending on flag.
        
        # Apply filters
        $filteredMessages = @()
        foreach ($msg in $messages) {
            # Skip if not a mail item
            if ($msg.Class -ne 43) { continue }
            
            # Filter by unread status (unless -Replied is specified, which implies showing read messages)
            if (!$IncludeRead -and !$All -and !$Replied -and $msg.UnRead -eq $false) { continue }

            # Filter by date using relative or absolute bounds
            if (!$All) {
                if ($lowerBound -and $msg.ReceivedTime -lt $lowerBound) { continue }
                if ($upperBound -and $msg.ReceivedTime -gt $upperBound) { continue }
            }

            # Filter by flagged status
            # FlagStatus: 0 = NoFlag, 1 = Complete, 2 = Flagged
            if ($Flagged -and $msg.FlagStatus -ne 2) { continue }

            # Filter by category
            if ($Category -and $msg.Categories -notlike "*$Category*") { continue }

            # Filter by sender
            if ($From) {
                $senderEmail = Get-SmtpAddress -MailItem $msg
                $senderMatch = $msg.SenderName -like "*$From*" -or $senderEmail -like "*$From*"
                if (!$senderMatch) { continue }
            }

            # Filter by attachments (exclude small images)
            if ($HasAttachments -and (Get-MeaningfulAttachmentCount -Attachments $msg.Attachments) -eq 0) { continue }

            # Filter by importance
            # Importance: 0 = Low, 1 = Normal, 2 = High
            if ($Importance) {
                $importanceValue = switch ($Importance) {
                    "Low" { 0 }
                    "Normal" { 1 }
                    "High" { 2 }
                }
                if ($msg.Importance -ne $importanceValue) { continue }
            }

            # Filter by subject
            if ($Subject -and $msg.Subject -notlike "*$Subject*") { continue }

            # Filter by replied status (check last to avoid performance hit)
            # Only check this property on messages that passed all other filters
            if (!$All) {
                try {
                    # Use MAPI property PR_LAST_VERB_EXECUTED to check if replied
                    # 102 = Reply, 103 = ReplyAll
                    $lastVerb = $msg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
                    $hasReplied = ($lastVerb -eq 102 -or $lastVerb -eq 103)

                    if ($Replied) {
                        # If -Replied flag is set, ONLY show messages that have been replied to
                        if (!$hasReplied) { continue }
                    } else {
                        # By default, exclude messages that have been replied to
                        if ($hasReplied) { continue }
                    }
                } catch {
                    # If property is not accessible, assume not replied
                    # If -Replied flag is set, skip messages where we can't determine replied status
                    if ($Replied) { continue }
                }
            }
            
            $filteredMessages += $msg
        }
        
        # Store total count before offset/limit
        $totalCount = $filteredMessages.Count
        
        # Apply offset if specified
        if ($Offset -gt 0) {
            if ($Offset -ge $filteredMessages.Count) {
                if (!$Plain) {
                    Write-Host "`nOffset ($Offset) is greater than or equal to the number of messages ($($filteredMessages.Count))." -ForegroundColor Yellow
                }
                return
            }
            $filteredMessages = $filteredMessages[$Offset..($filteredMessages.Count - 1)]
        }
        
        # Apply limit if specified
        if ($Limit -gt 0 -and $filteredMessages.Count -gt $Limit) {
            $filteredMessages = $filteredMessages[0..($Limit - 1)]
        }
        
        # Cache messages for later retrieval
        $script:CachedMessages = $filteredMessages
        $script:LastInboxParameters = $paramSnapshot
        
        # Display messages
        if ($filteredMessages.Count -eq 0) {
            if (!$Plain -and !$Silent) {
                Write-Host "`nNo messages found matching criteria." -ForegroundColor Yellow
            }
            return
        }

        # Skip display if Silent mode
        if ($Silent) {
            return
        }

        # Plain text output for piping
        if ($Plain) {
            $output = @()
            for ($i = 0; $i -lt $filteredMessages.Count; $i++) {
                $msg = $filteredMessages[$i]

                if ($Full) {
                    # Full email content output
                    $senderEmail = Get-SmtpAddress -MailItem $msg
                    $output += "=" * 80
                    $output += "MESSAGE $i"
                    $output += "From: $($msg.SenderName) <$senderEmail>"
                    $output += "To: $($msg.To)"
                    $output += "Subject: $($msg.Subject)"
                    $output += "Received: $($msg.ReceivedTime)"

                    # Add status tags
                    $statusTags = @()
                    if ($msg.UnRead) { $statusTags += "UNREAD" }

                    # Check if replied
                    try {
                        $lastVerb = $msg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
                        if ($lastVerb -eq 102 -or $lastVerb -eq 103) { $statusTags += "REPLIED" }
                    } catch { }

                    if ($msg.FlagStatus -eq 2) { $statusTags += "FLAGGED" }
                    $meaningfulAttCount = Get-MeaningfulAttachmentCount -Attachments $msg.Attachments
                    if ($meaningfulAttCount -gt 0) { $statusTags += "ATTACHMENTS:$meaningfulAttCount" }
                    if ($msg.Categories) { $statusTags += "CATEGORIES:$($msg.Categories)" }

                    if ($statusTags.Count -gt 0) {
                        $output += "Status: " + ($statusTags -join ", ")
                    }

                    if ($meaningfulAttCount -gt 0) {
                        $output += "Attachments:"
                        foreach ($att in $msg.Attachments) {
                            if (-not (Test-SmallImage -Attachment $att)) {
                                $output += "  - $($att.FileName) ($([math]::Round($att.Size / 1KB, 2)) KB)"
                            }
                        }
                    }

                    $output += ""
                    $output += Clean-EmailBody -Body $msg.Body -KeepLinks $KeepLinks
                    $output += ""
                } else {
                    # Summary output
                    $unread = if ($msg.UnRead) { "UNREAD" } else { "READ" }

                    # Check if replied using MAPI property
                    $hasReplied = $false
                    try {
                        $lastVerb = $msg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
                        $hasReplied = ($lastVerb -eq 102 -or $lastVerb -eq 103)
                    } catch { }
                    $repliedTag = if ($hasReplied) { "REPLIED" } else { "" }

                    $flaggedTag = if ($msg.FlagStatus -eq 2) { "FLAGGED" } else { "" }
                    $meaningfulAttCount = Get-MeaningfulAttachmentCount -Attachments $msg.Attachments
                    $attachments = if ($meaningfulAttCount -gt 0) { "ATTACHMENTS:$meaningfulAttCount" } else { "" }

                    $tags = @($unread, $repliedTag, $flaggedTag, $attachments) | Where-Object { $_ -ne "" }
                    $tagString = if ($tags.Count -gt 0) { "[" + ($tags -join ",") + "]" } else { "" }

                    $senderEmail = Get-SmtpAddress -MailItem $msg
                    $output += "[$i] $tagString $($msg.SenderName) <$senderEmail> | $($msg.Subject) | $($msg.ReceivedTime)"
                }
            }
            Write-Output ($output -join "`n")
            return
        }
        
        # Formatted display
        $displayText = "Found $totalCount message(s)"
        if ($Offset -gt 0 -or $Limit -gt 0) {
            $startIdx = $Offset
            $endIdx = $Offset + $filteredMessages.Count - 1
            $displayText += " (showing $startIdx-$endIdx)"
        }
        Write-Host "`n$displayText`n" -ForegroundColor Cyan
        
        for ($i = 0; $i -lt $filteredMessages.Count; $i++) {
            $msg = $filteredMessages[$i]

            # Clean and prepare preview
            $cleanBody = Clean-EmailBody -Body $msg.Body -KeepLinks $KeepLinks
            $preview = $cleanBody -replace "`r`n", " " -replace "`n", " "
            $previewLength = 80
            if ($Compact) {
            	$previewLength = 100
            }
            if ($preview.Length -gt $previewLength) {
                $preview = $preview.Substring(0, $previewLength) + "..."
            }

            if ($msg.UnRead) { 
                $unreadIndicator = "[UNREAD]" 
            } else { 
                $unreadIndicator = "[READ]"
            }

            # Check if replied using MAPI property
            $hasReplied = $false
            try {
                $lastVerb = $msg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
                $hasReplied = ($lastVerb -eq 102 -or $lastVerb -eq 103)
            } catch { }
            $repliedIndicator = if ($hasReplied) { " [REPLIED]" } else { "" }

            $meaningfulAttCount = Get-MeaningfulAttachmentCount -Attachments $msg.Attachments
            $attachmentIndicator = if ($meaningfulAttCount -gt 0) { " [ATT:$meaningfulAttCount]" } else { "" }
            $flagIndicator = if ($msg.FlagStatus -eq 2) { " [FLAG]" } else { "" }
            $importanceIndicator = if ($msg.Importance -eq 2) { " [!!]" } else { "" }

            # Resolve sender's SMTP address
            $senderEmail = Get-SmtpAddress -MailItem $msg

            if ($Compact) {
                # Compact 2-line format
                Write-Host "[$i] " -NoNewline -ForegroundColor Green
                Write-Host "$unreadIndicator" -NoNewline -ForegroundColor Yellow
                Write-Host "$repliedIndicator" -NoNewline -ForegroundColor Green
                Write-Host "$attachmentIndicator" -NoNewLine -ForegroundColor Magenta
                Write-Host "$flagIndicator$importanceIndicator" -NoNewline -ForegroundColor Red
                if ($msg.Categories) {
                    Write-ColoredCategories -Categories ([string]$msg.Categories) -Prefix " [" -Suffix "]" -NoNewline
                }
                Write-Host " $($msg.SenderName) - " -NoNewline -ForegroundColor Cyan
                Write-Host $msg.Subject -ForegroundColor White
                Write-Host "    $preview" -ForegroundColor DarkGray
            } else {
                # Original verbose format
                Write-Host "[$i] " -NoNewline -ForegroundColor Green
                Write-Host "$unreadIndicator" -NoNewline -ForegroundColor Yellow
                Write-Host "$repliedIndicator " -NoNewline -ForegroundColor Green
                Write-Host "$($msg.SenderName) <$senderEmail>" -NoNewline -ForegroundColor Cyan
                Write-Host "$attachmentIndicator" -NoNewline -ForegroundColor Magenta
                Write-Host "$flagIndicator$importanceIndicator" -NoNewline -ForegroundColor Red
                if ($msg.Categories) {
                    Write-ColoredCategories -Categories ([string]$msg.Categories) -Prefix " [" -Suffix "]" -NoNewline
                }
                Write-Host ""
                Write-Host "    Subject: " -NoNewline -ForegroundColor Gray
                Write-Host $msg.Subject -ForegroundColor White
                Write-Host "    Received: " -NoNewline -ForegroundColor Gray
                Write-Host $msg.ReceivedTime -ForegroundColor White
                Write-Host "    Preview: " -NoNewline -ForegroundColor Gray
                Write-Host $preview -ForegroundColor DarkGray
                Write-Host ""
            }
        }
        
        Write-Host "Use 'Read-OutlookMessage <number>' to read a message`n" -ForegroundColor Green
        
    }
    catch {
        Write-Error "Failed to access Outlook: $_"
        Write-Host "Make sure Outlook is installed and you're logged in." -ForegroundColor Yellow
    }
}


function Test-SmallImage {
    <#
    .SYNOPSIS
    Helper function to check if an attachment is a small image that should be skipped

    .PARAMETER Attachment
    The Outlook attachment object

    .PARAMETER ThresholdKB
    Size threshold in KB

    .RETURNS
    Boolean - True if this is a small image, False otherwise
    #>
    param(
        [Parameter(Mandatory=$true)]
        $Attachment,

        [int]$ThresholdKB = 15
    )

    $imageExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.ico', '.svg', '.webp', '.tiff', '.tif')
    $extension = [System.IO.Path]::GetExtension($Attachment.FileName).ToLower()

    if ($imageExtensions -contains $extension) {
        $sizeKB = $Attachment.Size / 1KB
        return ($sizeKB -lt $ThresholdKB)
    }

    return $false
}

function Get-MeaningfulAttachmentCount {
    <#
    .SYNOPSIS
    Count attachments excluding small images (signature icons, logos, etc.)

    .PARAMETER Attachments
    The Attachments collection from an Outlook message

    .PARAMETER ThresholdKB
    Size threshold in KB for small images. Default is 15KB.

    .RETURNS
    Integer - Count of meaningful attachments
    #>
    param(
        [Parameter(Mandatory=$true)]
        $Attachments,

        [int]$ThresholdKB = 15
    )

    $count = 0
    foreach ($att in $Attachments) {
        if (-not (Test-SmallImage -Attachment $att -ThresholdKB $ThresholdKB)) {
            $count++
        }
    }
    return $count
}

function Save-OutlookMessageAttachments {
    <#
    .SYNOPSIS
    Saves attachments from a message using the same filtering rules as Read-OutlookMessage.

    .PARAMETER Message
    The Outlook message object containing attachments.

    .PARAMETER AttachmentPath
    Destination directory. Defaults to current directory.

    .PARAMETER AttachmentFilter
    Optional include filter (partial match) for attachment names.

    .PARAMETER ExcludeFilter
    Optional exclude filter (partial match) for attachment names.

    .PARAMETER IncludeSmallImages
    Include small inline images when saving.

    .PARAMETER SmallImageThreshold
    Threshold (KB) that defines a small image.

    .PARAMETER Silent
    Suppress status output (useful for automated scenarios).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.__ComObject]$Message,
        [string]$AttachmentPath = ".",
        [string]$AttachmentFilter,
        [string]$ExcludeFilter,
        [switch]$IncludeSmallImages,
        [int]$SmallImageThreshold = 100,
        [switch]$Silent
    )

    if ($Message.Attachments.Count -eq 0) {
        if (-not $Silent) {
            Write-Host "No attachments to save." -ForegroundColor Yellow
        }
        return [pscustomobject]@{
            Saved        = 0
            SkippedSmall = 0
            Excluded     = 0
        }
    }

    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($AttachmentPath)

    if (!(Test-Path $resolvedPath)) {
        New-Item -ItemType Directory -Path $resolvedPath -Force | Out-Null
    }

    $savedCount = 0
    $skippedCount = 0
    $excludedCount = 0
    foreach ($att in $Message.Attachments) {
        if (-not $IncludeSmallImages -and (Test-SmallImage -Attachment $att -ThresholdKB $SmallImageThreshold)) {
            $skippedCount++
            continue
        }

        if ($AttachmentFilter -and $att.FileName -notlike "*$AttachmentFilter*") {
            continue
        }

        if ($ExcludeFilter -and $att.FileName -like "*$ExcludeFilter*") {
            $excludedCount++
            continue
        }

        $filepath = Join-Path $resolvedPath $att.FileName
        try {
            $att.SaveAsFile($filepath)
            $savedCount++
            if (-not $Silent) {
                Write-Host "Saved: $filepath" -ForegroundColor Green
            }
        } catch {
            if (-not $Silent) {
                Write-Host "Error saving $($att.FileName): $_" -ForegroundColor Red
            }
        }
    }

    if (-not $Silent) {
        if ($skippedCount -gt 0) {
            Write-Host "Skipped $skippedCount small image(s)" -ForegroundColor DarkGray
        }

        if ($excludedCount -gt 0) {
            Write-Host "Excluded $excludedCount attachment(s) matching exclusion filter" -ForegroundColor DarkGray
        }

        if ($AttachmentFilter -and $savedCount -eq 0) {
            Write-Host "No attachments matched filter: $AttachmentFilter" -ForegroundColor Yellow
        }
    }

    return [pscustomobject]@{
        Saved        = $savedCount
        SkippedSmall = $skippedCount
        Excluded     = $excludedCount
    }
}

function Read-OutlookMessage {
    <#
    .SYNOPSIS
    Reads a specific email message
    
    .PARAMETER Index
    The index number from Get-OutlookInbox
    
    .PARAMETER MarkAsRead
    Mark the message as read after viewing
    
    .PARAMETER SaveAttachments
    Switch to enable saving attachments

    .PARAMETER AttachmentPath
    Directory path where attachments should be saved. If not specified, defaults to current directory.
    Creates the directory if it doesn't exist.

    .PARAMETER AttachmentFilter
    Filter attachments by filename (partial match). Only attachments matching this pattern will be saved.
    If not specified, all attachments are saved.

    .PARAMETER ExcludeFilter
    Exclude attachments by filename (partial match). Attachments matching this pattern will be skipped.
    Can be used with AttachmentFilter for more precise control.

    .PARAMETER IncludeSmallImages
    Include small images (signature icons, logos) when saving attachments.
    By default, images smaller than SmallImageThreshold (15KB) are skipped.

    .PARAMETER SmallImageThreshold
    Size threshold in KB for small images. Default is 15KB.

    .PARAMETER Plain
    Output plain text without colors or formatting (suitable for piping)

    .PARAMETER Condense
    Remove single line breaks and condense text (preserves paragraph breaks)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [int]$Index,
        [switch]$MarkAsRead,

        [Alias("save")]
        [switch]$SaveAttachments,

        [Alias("out")]
        [string]$AttachmentPath = ".",

        [Alias("filter")]
        [string]$AttachmentFilter,

        [Alias("exclude")]
        [string]$ExcludeFilter,

        [switch]$IncludeSmallImages,

        [int]$SmallImageThreshold = 100,

        [switch]$Plain,

        [Alias("small")]
        [switch]$Condense,
        [switch]$KeepLinks,
        [switch]$ShowHistory
    )

    # Apply configuration defaults if parameters not explicitly provided
    if (-not $PSBoundParameters.ContainsKey('AttachmentPath')) {
        $configAttachmentPath = Get-ConfigValue -Section "ReadMessage" -Key "AttachmentPath" -Default "."
        $AttachmentPath = $configAttachmentPath
    }

    if (-not $PSBoundParameters.ContainsKey('SmallImageThreshold')) {
        $configThreshold = Get-ConfigValue -Section "ReadMessage" -Key "SmallImageThreshold" -Default 100
        $SmallImageThreshold = $configThreshold
    }

    if (-not $PSBoundParameters.ContainsKey('IncludeSmallImages')) {
        $configIncludeSmallImages = Get-ConfigValue -Section "ReadMessage" -Key "IncludeSmallImages" -Default $false
        if ($configIncludeSmallImages -eq $true) {
            $IncludeSmallImages = $true
        }
    }

    if ($script:CachedMessages.Count -eq 0) {
        Write-Host "No cached messages. Run Get-OutlookInbox first." -ForegroundColor Yellow
        return
    }
    
    if ($Index -lt 0 -or $Index -ge $script:CachedMessages.Count) {
        Write-Error "Index out of range. Valid range: 0-$($script:CachedMessages.Count - 1)"
        return
    }
    
    $msg = $script:CachedMessages[$Index]
    
    # Plain text output for piping
    if ($Plain) {
        $output = @()
        $senderEmail = Get-SmtpAddress -MailItem $msg
        $output += "From: $($msg.SenderName) <$senderEmail>"
        $output += "To: $($msg.To)"
        $output += "Subject: $($msg.Subject)"
        $output += "Received: $($msg.ReceivedTime)"
        
        $meaningfulAttCount = Get-MeaningfulAttachmentCount -Attachments $msg.Attachments
        if ($meaningfulAttCount -gt 0) {
            $output += "Attachments: $meaningfulAttCount"
            foreach ($att in $msg.Attachments) {
                if (-not (Test-SmallImage -Attachment $att)) {
                    $output += "  - $($att.FileName) ($([math]::Round($att.Size / 1KB, 2)) KB)"
                }
            }
        }
        
        $output += ""

        $body = Clean-EmailBody -Body $msg.Body -KeepLinks $KeepLinks

        # Apply condensing if requested
        if ($Condense) {
            # Remove all blank lines while preserving list structure
            $body = ($body -split "`n" | Where-Object { $_.Trim() -ne "" }) -join "`n"
        }

        $threadParts = Split-EmailThread -Body $body

        if ($ShowHistory) {
            $output += $threadParts.Latest
            if ($threadParts.HasHistory) {
                $output += ""
                $output += "----- Previous Messages -----"
                $output += $threadParts.History
            }
        } else {
            $output += $threadParts.Latest
            if ($threadParts.HasHistory) {
                $output += ""
                $output += "[History hidden. Re-run with -ShowHistory to view the full thread.]"
            }
        }

        Write-Output ($output -join "`n")
        
        # Handle mark as read and attachments silently
        if ($MarkAsRead -and $msg.UnRead) {
            $msg.UnRead = $false
            $msg.Save()
        }
        
        if ($SaveAttachments) {
            Save-OutlookMessageAttachments -Message $msg -AttachmentPath $AttachmentPath -AttachmentFilter $AttachmentFilter `
                -ExcludeFilter $ExcludeFilter -IncludeSmallImages:$IncludeSmallImages -SmallImageThreshold $SmallImageThreshold -Silent
        }
        
        return
    }

    # Display message details
    Write-Host ""
    Write-Host ("=" * 80) -ForegroundColor Cyan
    $senderEmail = Get-SmtpAddress -MailItem $msg
    Write-Host "From: " -NoNewline -ForegroundColor Gray
    Write-Host "$($msg.SenderName) <$senderEmail>" -ForegroundColor White
    Write-Host "To: " -NoNewline -ForegroundColor Gray
    Write-Host $msg.To -ForegroundColor White
    Write-Host "Subject: " -NoNewline -ForegroundColor Gray
    Write-Host $msg.Subject -ForegroundColor White
    Write-Host "Received: " -NoNewline -ForegroundColor Gray
    Write-Host $msg.ReceivedTime -ForegroundColor White
    
    $meaningfulAttCount = Get-MeaningfulAttachmentCount -Attachments $msg.Attachments
    if ($meaningfulAttCount -gt 0) {
        Write-Host "Attachments: " -NoNewline -ForegroundColor Gray
        Write-Host $meaningfulAttCount -ForegroundColor Magenta
        foreach ($att in $msg.Attachments) {
            if (-not (Test-SmallImage -Attachment $att)) {
                Write-Host "  - $($att.FileName) ($([math]::Round($att.Size / 1KB, 2)) KB)" -ForegroundColor Magenta
            }
        }
    }
    
    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host ""
    
    # Display body
    if ($msg.BodyFormat -eq 2) { # HTML
        # Strip HTML tags for terminal display
        $body = $msg.Body
    } else {
        $body = $msg.Body
    }

    $body = Clean-EmailBody -Body $body -KeepLinks $KeepLinks

    # Apply condensing if requested
    if ($Condense) {
        # Remove all blank lines while preserving list structure
        $body = ($body -split "`n" | Where-Object { $_.Trim() -ne "" }) -join "`n"
    }

    $threadParts = Split-EmailThread -Body $body

    Write-Host $threadParts.Latest
    if ($threadParts.HasHistory) {
        if ($ShowHistory) {
            Write-Host ""
            Write-Host ("-" * 32) -ForegroundColor DarkGray
            Write-Host "Previous Messages" -ForegroundColor DarkGray
            Write-Host ("-" * 32) -ForegroundColor DarkGray
            Write-Host $threadParts.History -ForegroundColor DarkGray
        } else {
            Write-Host ""
            Write-Host "[History hidden. Toggle thread view to display previous messages.]" -ForegroundColor DarkGray
        }
    }

    Write-Host "`n" + ("=" * 80) -ForegroundColor Cyan
    
    # Mark as read if requested
    if ($MarkAsRead -and $msg.UnRead) {
        $msg.UnRead = $false
        $msg.Save()
        Write-Host "Marked as read." -ForegroundColor Green
    }
    
    # Save attachments if requested
    if ($SaveAttachments) {
        Save-OutlookMessageAttachments -Message $msg -AttachmentPath $AttachmentPath -AttachmentFilter $AttachmentFilter `
            -ExcludeFilter $ExcludeFilter -IncludeSmallImages:$IncludeSmallImages -SmallImageThreshold $SmallImageThreshold
    }
}

function Send-OutlookReply {
    <#
    .SYNOPSIS
    Reply to an email message

    .PARAMETER Index
    The index number from Get-OutlookInbox

    .PARAMETER Body
    The reply message body. If not provided, opens the editor specified in $env:EDITOR, or 'micro' if not set

    .PARAMETER Direct
    Reply only to the sender instead of all recipients (default is reply-all)

    .PARAMETER Attachments
    Path(s) to file(s) to attach. Can be a single path or array of paths
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$Index,

        [string]$Body,

        [switch]$Direct,

        [string[]]$Attachments
    )

    if ($script:CachedMessages.Count -eq 0) {
        Write-Host "No cached messages. Run Get-OutlookInbox first." -ForegroundColor Yellow
        return
    }

    if ($Index -lt 0 -or $Index -ge $script:CachedMessages.Count) {
        Write-Error "Index out of range. Valid range: 0-$($script:CachedMessages.Count - 1)"
        return
    }

    $msg = $script:CachedMessages[$Index]

    # Create reply (default to reply-all unless -Direct is specified)
    if ($Direct) {
        $reply = $msg.Reply()
    } else {
        $reply = $msg.ReplyAll()
    }
    
    # Get reply body
    if ([string]::IsNullOrWhiteSpace($Body)) {
        # Create temp file for editing
        $tempFile = [System.IO.Path]::GetTempFileName()
        $tempFile = [System.IO.Path]::ChangeExtension($tempFile, ".txt")

        # Add helpful header
        $header = "# Enter your reply below this line. Save and close to send.`n"
        $header += "# To: $($msg.SenderName)`n"
        $header += "# Re: $($msg.Subject)`n"
        $header += "#`n"
        $header += "# Original message:`n"
        $header += "# " + ("-" * 70) + "`n"

        # Add original message body, commented out
        $cleanBody = Clean-EmailBody -Body $msg.Body
        $threadParts = Split-EmailThread -Body $cleanBody
        $originalText = $threadParts.Latest

        # Limit to first 30 lines to keep editor manageable
        $bodyLines = $originalText -split "`n"
        $lineCount = [Math]::Min($bodyLines.Count, 30)
        for ($i = 0; $i -lt $lineCount; $i++) {
            $header += "# $($bodyLines[$i])`n"
        }
        if ($bodyLines.Count -gt 30) {
            $header += "# ... (message truncated)`n"
        }

        $header += "# " + ("-" * 70) + "`n"
        $header += "`n"

        Set-Content -Path $tempFile -Value $header

        # Open in editor
        Write-Host "Opening editor... Save and close when done." -ForegroundColor Cyan
        if ($env:EDITOR) { & $env:EDITOR $tempFile } else { micro $tempFile }

        # Read the content
        $Body = Get-Content -Path $tempFile -Raw

        # Remove header lines
        $Body = $Body -replace "(?m)^#.*`n", ""

        Remove-Item $tempFile
    }
    
    if ([string]::IsNullOrWhiteSpace($Body)) {
        Write-Host "Reply cancelled (empty message)." -ForegroundColor Yellow
        return
    }
    
    # Set reply body (prepend to existing reply which includes original message)
    $reply.Body = $Body + "`n`n" + $reply.Body

    # Add attachments
    if ($Attachments) {
        foreach ($attachment in $Attachments) {
            if (Test-Path $attachment) {
                $reply.Attachments.Add((Resolve-Path $attachment).Path) | Out-Null
                Write-Host "Attached: $attachment" -ForegroundColor Green
            } else {
                Write-Warning "Attachment not found: $attachment"
            }
        }
    }

    # Confirm before sending
    if ($Direct) {
        Write-Host "`nDirect reply preview:" -ForegroundColor Cyan
    } else {
        Write-Host "`nReply-all preview:" -ForegroundColor Cyan
    }
    Write-Host "To: $($reply.To)" -ForegroundColor White
    if (-not $Direct -and $reply.CC) {
        Write-Host "CC: $($reply.CC)" -ForegroundColor White
    }
    Write-Host "Subject: $($reply.Subject)" -ForegroundColor White
    if ($reply.Attachments.Count -gt 0) {
        Write-Host "Attachments: $($reply.Attachments.Count) file(s)" -ForegroundColor Magenta
    }
    Write-Host "`nBody:" -ForegroundColor White
    Write-Host $Body -ForegroundColor Gray
    
    $confirm = Read-Host "`nSend this reply? (y/N)"
    
    if ($confirm -eq 'y' -or $confirm -eq 'Y') {
        $reply.Send()
        Write-Host "Reply sent!" -ForegroundColor Green

        # Mark original as read
        if ($msg.UnRead) {
            $msg.UnRead = $false
            $msg.Save()
        }

        # Check if sender is in contacts, offer to add if not
        $senderEmail = Get-SmtpAddress -MailItem $msg
        $senderInContacts = Test-OutlookContactExists -Email $senderEmail
        if (-not $senderInContacts) {
            $addContactChoice = Read-Host "`nAdd $($msg.SenderName) to contacts? (y/N)"
            if ($addContactChoice -eq 'y' -or $addContactChoice -eq 'Y') {
                # Parse sender name into first/last
                $firstName = ""
                $lastName = ""
                if ($msg.SenderName -and $msg.SenderName -ne $senderEmail) {
                    $nameParts = $msg.SenderName -split " ", 2
                    $firstName = $nameParts[0]
                    if ($nameParts.Count -gt 1) {
                        $lastName = $nameParts[1]
                    }
                }
                Add-OutlookContact -Email $senderEmail -FirstName $firstName -LastName $lastName
            }
        }

        # Prompt to archive
        $archiveChoice = Read-Host "`nArchive this email? (y/N)"
        if ($archiveChoice -eq 'y' -or $archiveChoice -eq 'Y') {
            try {
                # Connect to Outlook
                $outlook = New-Object -ComObject Outlook.Application
                $namespace = $outlook.GetNamespace("MAPI")

                # Get Archive folder
                $archiveFolderId = 109
                try {
                    $archiveFolderId = [int][Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderArchive
                } catch {
                    # Keep numeric fallback
                }

                try {
                    $archiveFolder = $namespace.GetDefaultFolder($archiveFolderId)
                } catch {
                    # Fallback: search for Archive folder by name
                    $inbox = $namespace.GetDefaultFolder(6)
                    $archiveFolder = $inbox.Parent.Folders | Where-Object { $_.Name -eq "Archive" } | Select-Object -First 1

                    if (!$archiveFolder) {
                        Write-Host "Archive folder not found. Message not archived." -ForegroundColor Yellow
                        return
                    }
                }

                # Move to archive
                $msg.Move($archiveFolder) | Out-Null
                Write-Host "Message archived." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to archive message: $_" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "Reply cancelled." -ForegroundColor Yellow
    }
}

function Get-OutlookCategories {
    <#
    .SYNOPSIS
    Lists all available Outlook categories/tags
    #>
    [CmdletBinding()]
    param()

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $categories = $namespace.Categories

        Write-Host "`nAvailable Outlook Categories:" -ForegroundColor Cyan
        Write-Host ("=" * 50) -ForegroundColor Cyan

        # Get color map for consistent mapping
        $colorMap = Get-CategoryColorMap

        foreach ($cat in $categories) {
            Write-Host "- " -NoNewline

            # Display category name in its associated console color
            $consoleColor = if ($colorMap.ContainsKey($cat.Name)) { $colorMap[$cat.Name] } else { "White" }
            Write-Host $cat.Name -ForegroundColor $consoleColor -NoNewline

            # Show Outlook color name
            $colorName = switch ($cat.Color) {
                0 { "None" }
                1 { "Red" }
                2 { "Orange" }
                3 { "Peach" }
                4 { "Yellow" }
                5 { "Green" }
                6 { "Teal" }
                7 { "Olive" }
                8 { "Blue" }
                9 { "Purple" }
                10 { "Maroon" }
                11 { "Steel" }
                12 { "DarkSteel" }
                13 { "Gray" }
                14 { "DarkGray" }
                15 { "Black" }
                16 { "DarkRed" }
                17 { "DarkOrange" }
                18 { "DarkPeach" }
                19 { "DarkYellow" }
                20 { "DarkGreen" }
                21 { "DarkTeal" }
                22 { "DarkOlive" }
                23 { "DarkBlue" }
                24 { "DarkPurple" }
                25 { "DarkMaroon" }
                default { "Unknown" }
            }

            Write-Host " ($colorName)" -ForegroundColor DarkGray
        }

        Write-Host "`nUse -Category parameter with Get-OutlookInbox to filter by category" -ForegroundColor Green
        Write-Host "Example: Get-OutlookInbox -Category 'Red Category' -All`n" -ForegroundColor Gray

    }
    catch {
        Write-Error "Failed to retrieve categories: $_"
    }
}

function Move-OutlookMessageToArchive {
    <#
    .SYNOPSIS
    Archive email messages by moving them to the Archive folder
    
    .PARAMETER Index
    The index number(s) from Get-OutlookInbox. Can be a single index or comma-separated list (e.g., 0,2,5)
    
    .PARAMETER MarkAsRead
    Mark the message(s) as read before archiving
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int[]]$Index,
        
        [switch]$MarkAsRead
    )
    
    if ($script:CachedMessages.Count -eq 0) {
        Write-Host "No cached messages. Run Get-OutlookInbox first." -ForegroundColor Yellow
        return
    }
    
    try {
        # Connect to Outlook
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Get Archive folder using proper default folder id (109 = Archive)
        # If that fails, try to find it by name
        $archiveFolderId = 109
        try {
            $archiveFolderId = [int][Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderArchive
        } catch {
            # Keep numeric fallback for older Outlook versions lacking the enum type
        }

        try {
            $archiveFolder = $namespace.GetDefaultFolder($archiveFolderId)
        } catch {
            # Fallback: search for Archive folder by name
            $inbox = $namespace.GetDefaultFolder(6)
            $archiveFolder = $inbox.Parent.Folders | Where-Object { $_.Name -eq "Archive" } | Select-Object -First 1
            
            if (!$archiveFolder) {
                Write-Error "Archive folder not found. Please ensure you have an Archive folder in Outlook."
                return
            }
        }
        
        $archivedCount = 0
        $errors = @()
        
        foreach ($i in $Index) {
            if ($i -lt 0 -or $i -ge $script:CachedMessages.Count) {
                $errors += "Index $i is out of range (valid: 0-$($script:CachedMessages.Count - 1))"
                continue
            }
            
            $msg = $script:CachedMessages[$i]
            
            # Mark as read if requested
            if ($MarkAsRead -and $msg.UnRead) {
                $msg.UnRead = $false
                $msg.Save()
            }
            
            # Move to archive folder
            $msg.Move($archiveFolder) | Out-Null
            
            Write-Host "[$i] Archived: " -NoNewline -ForegroundColor Green
            Write-Host $msg.Subject -ForegroundColor White
            
            $archivedCount++
        }
        
        # Show summary
        Write-Host "`nArchived $archivedCount message(s)" -ForegroundColor Cyan
        
        if ($errors.Count -gt 0) {
            Write-Host "`nErrors:" -ForegroundColor Red
            $errors | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        }
        
        # Suggest refreshing the inbox
        if ($archivedCount -gt 0) {
            Write-Host "Run Get-OutlookInbox to refresh the message list" -ForegroundColor Gray
        }
        
    }
    catch {
        Write-Error "Failed to archive message(s): $_"
    }
}

function Send-ForwardOutlookEmail {
    <#
    .SYNOPSIS
    Forward an email message to one or more recipients

    .PARAMETER Index
    The index number from Get-OutlookInbox

    .PARAMETER To
    Recipient email address(es). Separate multiple recipients with semicolons.
    If not provided, will prompt for recipients.

    .PARAMETER Body
    Optional message to add before the forwarded content. If not provided, opens the editor specified in $env:EDITOR, or 'micro' if not set

    .PARAMETER CC
    CC recipient(s). Separate multiple recipients with semicolons

    .PARAMETER BCC
    BCC recipient(s). Separate multiple recipients with semicolons
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$Index,

        [string]$To,

        [string]$Body,

        [string]$CC,

        [string]$BCC
    )

    if ($script:CachedMessages.Count -eq 0) {
        Write-Host "No cached messages. Run Get-OutlookInbox first." -ForegroundColor Yellow
        return
    }

    if ($Index -lt 0 -or $Index -ge $script:CachedMessages.Count) {
        Write-Error "Index out of range. Valid range: 0-$($script:CachedMessages.Count - 1)"
        return
    }

    $msg = $script:CachedMessages[$Index]

    # Create forward
    $forward = $msg.Forward()

    # Get recipients if not provided
    if ([string]::IsNullOrWhiteSpace($To)) {
        Write-Host "`nForwarding: " -NoNewline -ForegroundColor Cyan
        Write-Host $msg.Subject -ForegroundColor White
        $To = Read-Host "To (email addresses separated by semicolons)"

        if ([string]::IsNullOrWhiteSpace($To)) {
            Write-Host "Forward cancelled (no recipients specified)." -ForegroundColor Yellow
            return
        }
    }

    # Set recipients
    $forward.To = $To
    if ($CC) { $forward.CC = $CC }
    if ($BCC) { $forward.BCC = $BCC }

    # Get forward message body
    if ([string]::IsNullOrWhiteSpace($Body)) {
        # Create temp file for editing
        $tempFile = [System.IO.Path]::GetTempFileName()
        $tempFile = [System.IO.Path]::ChangeExtension($tempFile, ".txt")

        # Add helpful header
        $header = "# Enter your message below this line (optional). Save and close to send.`n"
        $header += "# Forwarding: $($msg.Subject)`n"
        $header += "# To: $To`n"
        if ($CC) { $header += "# CC: $CC`n" }
        if ($BCC) { $header += "# BCC: $BCC`n" }
        $header += "`n"
        Set-Content -Path $tempFile -Value $header

        # Open in editor
        Write-Host "Opening editor... Save and close when done." -ForegroundColor Cyan
        if ($env:EDITOR) { & $env:EDITOR $tempFile } else { micro $tempFile }

        # Read the content
        $Body = Get-Content -Path $tempFile -Raw

        # Remove header lines
        $Body = $Body -replace "(?m)^#.*`n", ""

        Remove-Item $tempFile
    }

    # Set forward body (prepend message to existing forward content if provided)
    if (-not [string]::IsNullOrWhiteSpace($Body)) {
        $forward.Body = $Body.Trim() + "`n`n" + $forward.Body
    }

    # Confirm before sending
    Write-Host "`nForward preview:" -ForegroundColor Cyan
    Write-Host "To: $($forward.To)" -ForegroundColor White
    if ($forward.CC) {
        Write-Host "CC: $($forward.CC)" -ForegroundColor White
    }
    if ($forward.BCC) {
        Write-Host "BCC: $($forward.BCC)" -ForegroundColor White
    }
    Write-Host "Subject: $($forward.Subject)" -ForegroundColor White

    if (-not [string]::IsNullOrWhiteSpace($Body)) {
        Write-Host "`nYour message:" -ForegroundColor White
        Write-Host $Body.Trim() -ForegroundColor Gray
    }

    $confirm = Read-Host "`nSend this forward? (y/N)"

    if ($confirm -eq 'y' -or $confirm -eq 'Y') {
        $forward.Send()
        Write-Host "Forward sent!" -ForegroundColor Green

        # Mark original as read
        if ($msg.UnRead) {
            $msg.UnRead = $false
            $msg.Save()
        }
    } else {
        Write-Host "Forward cancelled." -ForegroundColor Yellow
    }
}

function Send-OutlookEmail {
    <#
    .SYNOPSIS
    Send a new email message

    .PARAMETER To
    Recipient email address(es). Separate multiple recipients with semicolons

    .PARAMETER Subject
    Email subject line

    .PARAMETER Body
    Email body text. If not provided, opens the editor specified in $env:EDITOR, or 'micro' if not set

    .PARAMETER CC
    CC recipient(s). Separate multiple recipients with semicolons

    .PARAMETER BCC
    BCC recipient(s). Separate multiple recipients with semicolons

    .PARAMETER Attachments
    Path(s) to file(s) to attach. Can be a single path or array of paths

    .PARAMETER Importance
    Set importance level: High, Normal, or Low
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$To,

        [Parameter(Mandatory=$true)]
        [string]$Subject,

        [string]$Body,

        [string]$CC,

        [string]$BCC,

        [string[]]$Attachments,

        [ValidateSet("High", "Normal", "Low")]
        [string]$Importance = "Normal"
    )

    # Apply configuration defaults if parameters not explicitly provided
    if (-not $PSBoundParameters.ContainsKey('Importance')) {
        $configImportance = Get-ConfigValue -Section "SendEmail" -Key "Importance" -Default "Normal"
        # Validate the config value is one of the allowed values
        if ($configImportance -in @("High", "Normal", "Low")) {
            $Importance = $configImportance
        }
    }

    try {
        # Connect to Outlook
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0) # 0 = Mail item
        
        # Set recipients
        $mail.To = $To
        if ($CC) { $mail.CC = $CC }
        if ($BCC) { $mail.BCC = $BCC }
        
        # Set subject
        $mail.Subject = $Subject
        
        # Get body content
        if ([string]::IsNullOrWhiteSpace($Body)) {
            # Create temp file for editing
            $tempFile = [System.IO.Path]::GetTempFileName()
            $tempFile = [System.IO.Path]::ChangeExtension($tempFile, ".txt")
            
            # Add helpful header
            $header = "# Enter your email below this line. Save and close to send.`n"
            $header += "# To: $To`n"
            $header += "# Subject: $Subject`n"
            if ($CC) { $header += "# CC: $CC`n" }
            if ($BCC) { $header += "# BCC: $BCC`n" }
            $header += "`n"
            Set-Content -Path $tempFile -Value $header
            
            # Open in editor
            Write-Host "Opening editor... Save and close when done." -ForegroundColor Cyan
            if ($env:EDITOR) { & $env:EDITOR $tempFile } else { micro $tempFile }
            
            # Read the content
            $Body = Get-Content -Path $tempFile -Raw
            
            # Remove header lines
            $Body = $Body -replace "(?m)^#.*`n", ""
            
            Remove-Item $tempFile
        }
        
        if ([string]::IsNullOrWhiteSpace($Body)) {
            Write-Host "Email cancelled (empty message)." -ForegroundColor Yellow
            return
        }
        
        # Set body
        $mail.Body = $Body
        
        # Set importance
        $mail.Importance = switch ($Importance) {
            "Low" { 0 }
            "Normal" { 1 }
            "High" { 2 }
        }
        
        # Add attachments
        if ($Attachments) {
            foreach ($attachment in $Attachments) {
                if (Test-Path $attachment) {
                    $mail.Attachments.Add((Resolve-Path $attachment).Path) | Out-Null
                    Write-Host "Attached: $attachment" -ForegroundColor Green
                } else {
                    Write-Warning "Attachment not found: $attachment"
                }
            }
        }
        
        # Show preview
        Write-Host "`nEmail preview:" -ForegroundColor Cyan
        Write-Host "To: $($mail.To)" -ForegroundColor White
        if ($mail.CC) { Write-Host "CC: $($mail.CC)" -ForegroundColor White }
        if ($mail.BCC) { Write-Host "BCC: $($mail.BCC)" -ForegroundColor White }
        Write-Host "Subject: $($mail.Subject)" -ForegroundColor White
        if ($mail.Attachments.Count -gt 0) {
            Write-Host "Attachments: $($mail.Attachments.Count) file(s)" -ForegroundColor Magenta
        }
        if ($Importance -ne "Normal") {
            Write-Host "Importance: $Importance" -ForegroundColor Yellow
        }
        Write-Host "`nBody:" -ForegroundColor White
        Write-Host $Body -ForegroundColor Gray
        
        # Confirm before sending
        $confirm = Read-Host "`nSend this email? (y/N)"
        
        if ($confirm -eq 'y' -or $confirm -eq 'Y') {
            $mail.Send()
            Write-Host "Email sent!" -ForegroundColor Green
        } else {
            Write-Host "Email cancelled." -ForegroundColor Yellow
        }
        
    }
    catch {
        Write-Error "Failed to send email: $_"
    }
}


function Start-OutlookInteractive {
    <#
    .SYNOPSIS
    Interactive mode for reading and responding to emails

    .PARAMETER DaysBack
    Only show emails from the last N days (default: 7)

    .PARAMETER PreviewLength
    Number of characters to show in message previews (default: 600)

    .PARAMETER Before
    Combine with -DaysBack to focus on a window of relative days (e.g., 30 to 7 days ago)

    .PARAMETER StartDate
    Inclusive lower bound for ReceivedTime (mutually exclusive with -DaysBack/-Before)

    .PARAMETER EndDate
    Inclusive upper bound for ReceivedTime (mutually exclusive with -DaysBack/-Before)
    #>
    [CmdletBinding()]
    param(
        [Alias("days")]
        [int]$DaysBack = 7,
        [Alias("pre")]
        [int]$PreviewLength = 600,
        [Alias("read")]
        [switch]$IncludeRead,
        [switch]$KeepLinks,
        [switch]$Oldest,
        [switch]$Cached,
        [switch]$All,
        [switch]$Replied,
        [switch]$Flagged,
        [string]$Category,
        [string]$From,
        [switch]$HasAttachments,
        [ValidateSet("High","Normal","Low")]
        [string]$Importance,
        [string]$Subject,
        [Alias("n")]
        [int]$Limit = 0,
        [int]$Offset = 0,
        [int]$Before = 0,
        [datetime]$StartDate,
        [datetime]$EndDate,
        [switch]$Silent
    )

    # Apply configuration defaults if parameters not explicitly provided
    if (-not $PSBoundParameters.ContainsKey('PreviewLength')) {
        $configPreviewLength = Get-ConfigValue -Section "Interactive" -Key "PreviewLength" -Default 600
        $PreviewLength = $configPreviewLength
    }

    if (-not $PSBoundParameters.ContainsKey('DaysBack')) {
        $configDaysBack = Get-ConfigValue -Section "Interactive" -Key "DaysBack" -Default 7
        $DaysBack = $configDaysBack
    }

    if (-not $PSBoundParameters.ContainsKey('IncludeRead')) {
        $configIncludeRead = Get-ConfigValue -Section "Interactive" -Key "IncludeRead" -Default $false
        if ($configIncludeRead -eq $true) {
            $IncludeRead = $true
        }
    }

    if (-not $PSBoundParameters.ContainsKey('Limit')) {
        $configLimit = Get-ConfigValue -Section "Interactive" -Key "Limit" -Default 0
        if ($configLimit -gt 0) {
            $Limit = $configLimit
        }
    }

    if (-not $PSBoundParameters.ContainsKey('KeepLinks')) {
        $configKeepLinks = Get-ConfigValue -Section "Interactive" -Key "KeepLinks" -Default $false
        if ($configKeepLinks -eq $true) {
            $KeepLinks = $true
        }
    }

    if (-not $PSBoundParameters.ContainsKey('Oldest')) {
        $configOldest = Get-ConfigValue -Section "Interactive" -Key "Oldest" -Default $false
        if ($configOldest -eq $true) {
            $Oldest = $true
        }
    }

    Write-Host "`n=============================================================================" -ForegroundColor Cyan
    Write-Host "===    Outlook Interactive Mode                                           ===" -ForegroundColor Cyan
    Write-Host "=============================================================================`n" -ForegroundColor Cyan
    
    $filterParamNames = @(
        'IncludeRead','DaysBack','All','Replied','Flagged','Category',
        'From','HasAttachments','Importance','Subject','Limit','Offset','Oldest','KeepLinks',
        'Before','StartDate','EndDate'
    )
    $userProvidedFilters = $filterParamNames | Where-Object { $PSBoundParameters.ContainsKey($_) }

    function Get-InteractiveFilterSnapshot {
        param()
        $snapshot = @{}
        $absoluteDates = $PSBoundParameters.ContainsKey('StartDate') -or $PSBoundParameters.ContainsKey('EndDate')
        if ($absoluteDates) {
            if ($PSBoundParameters.ContainsKey('StartDate')) { $snapshot['StartDate'] = $StartDate }
            if ($PSBoundParameters.ContainsKey('EndDate')) { $snapshot['EndDate'] = $EndDate }
        } else {
            $snapshot['DaysBack'] = $DaysBack
            if ($Before -gt 0) { $snapshot['Before'] = $Before }
        }
        if ($IncludeRead)    { $snapshot['IncludeRead']   = $true }
        if ($All)            { $snapshot['All']           = $true }
        if ($Replied)        { $snapshot['Replied']       = $true }
        if ($Flagged)        { $snapshot['Flagged']       = $true }
        if ($Category)       { $snapshot['Category']      = $Category }
        if ($From)           { $snapshot['From']          = $From }
        if ($HasAttachments) { $snapshot['HasAttachments']= $true }
        if ($Importance)     { $snapshot['Importance']    = $Importance }
        if ($Subject)        { $snapshot['Subject']       = $Subject }
        if ($Limit -gt 0)    { $snapshot['Limit']         = $Limit }
        if ($Offset -gt 0)   { $snapshot['Offset']        = $Offset }
        if ($Oldest)         { $snapshot['Oldest']        = $true }
        if ($KeepLinks)      { $snapshot['KeepLinks']     = $true }
        return $snapshot
    }

    if ($Cached -and $userProvidedFilters.Count -gt 0) {
        Write-Host "Cached mode ignores new filters. Remove -Cached to re-query with filters." -ForegroundColor Yellow
        return
    }

    $sessionInboxParams = @{}
    if ($userProvidedFilters.Count -gt 0) {
        $sessionInboxParams = Get-InteractiveFilterSnapshot
    } elseif ($script:LastInboxParameters.Count -gt 0) {
        foreach ($key in $script:LastInboxParameters.Keys) {
            $sessionInboxParams[$key] = $script:LastInboxParameters[$key]
        }
    } else {
        $sessionInboxParams = Get-InteractiveFilterSnapshot
    }

    $cacheAvailable = $script:CachedMessages.Count -gt 0
    $useCacheOnly = $Cached -and $cacheAvailable
    $shouldLoadFresh = -not $useCacheOnly

    if ($Cached -and -not $cacheAvailable) {
        Write-Host "No cached messages available. Loading with the current filters instead." -ForegroundColor Yellow
        $shouldLoadFresh = $true
        $useCacheOnly = $false
    }

    if (-not $useCacheOnly -and $cacheAvailable -and $userProvidedFilters.Count -eq 0) {
        Write-Host "Using $($script:CachedMessages.Count) cached message(s). Pass filters or omit cache to refresh." -ForegroundColor Gray
        $shouldLoadFresh = $false
    }

    if ($shouldLoadFresh) {
        $initialParams = @{}
        foreach ($key in $sessionInboxParams.Keys) {
            $initialParams[$key] = $sessionInboxParams[$key]
        }
        $initialParams['Compact'] = $true
        if ($Silent) {
            $initialParams['Silent'] = $true
        } elseif ($initialParams.ContainsKey('Silent')) {
            $initialParams.Remove('Silent')
        }
        Get-OutlookInbox @initialParams
    }

    if ($script:CachedMessages.Count -eq 0) {
        Write-Host "No messages match the current criteria. Exiting." -ForegroundColor Yellow
        return
    }

    $refreshAction = {
        param([switch]$SilentRefresh)
        $params = @{}
        foreach ($key in $sessionInboxParams.Keys) {
            $params[$key] = $sessionInboxParams[$key]
        }
        $params['Compact'] = $true
        if ($SilentRefresh) {
            $params['Silent'] = $true
        } else {
            $params.Remove('Silent')
        }
        Get-OutlookInbox @params
    }
    
    $currentIndex = 0
    $showFullThread = $false
    
    while ($true) {
        $msg = $script:CachedMessages[$currentIndex]
        
        Write-Host "`n=======================================================================" -ForegroundColor Cyan
        Write-Host "Message $($currentIndex + 1) of $($script:CachedMessages.Count)" -ForegroundColor Cyan
        Write-Host "=======================================================================`n" -ForegroundColor Cyan

        # Display message preview
        $senderEmail = Get-SmtpAddress -MailItem $msg
        Write-Host "From: " -NoNewline -ForegroundColor Gray
        Write-Host "$($msg.SenderName) <$senderEmail>" -ForegroundColor White
        Write-Host "Subject: " -NoNewline -ForegroundColor Gray
        Write-Host $msg.Subject -ForegroundColor White
        Write-Host "Received: " -NoNewline -ForegroundColor Gray
        Write-Host $msg.ReceivedTime -ForegroundColor White

        $meaningfulAttCount = Get-MeaningfulAttachmentCount -Attachments $msg.Attachments
        if ($meaningfulAttCount -gt 0) {
            Write-Host "Attachments: " -NoNewline -ForegroundColor Gray
            Write-Host "$meaningfulAttCount file(s)" -ForegroundColor Magenta
        }

        # Show flag status
        if ($msg.FlagStatus -eq 2) {
            Write-Host "Flag: " -NoNewline -ForegroundColor Gray
            Write-Host "FLAGGED" -NoNewline -ForegroundColor Red
            if ($msg.FlagRequest) {
                Write-Host " ($($msg.FlagRequest))" -NoNewline -ForegroundColor DarkGray
            }
            if ($msg.FlagDueBy) {
                Write-Host " - Due: $($msg.FlagDueBy)" -NoNewline -ForegroundColor Yellow
            }
            Write-Host ""
        }

        # Show categories
        if ($msg.Categories) {
            Write-Host "Categories: " -NoNewline -ForegroundColor Gray
            Write-ColoredCategories -Categories ([string]$msg.Categories) -NoNewline
            Write-Host ""
        }

        # Show importance if high
        if ($msg.Importance -eq 2) {
            Write-Host "Importance: " -NoNewline -ForegroundColor Gray
            Write-Host "HIGH" -ForegroundColor Red
        }

        Write-Host "`nPreview:" -ForegroundColor Gray
        $cleanBody = Clean-EmailBody -Body $msg.Body -KeepLinks $KeepLinks
        $previewSource = Split-EmailThread -Body $cleanBody
        $preview = $previewSource.Latest -replace "`r`n", " " -replace "`n", " "
        if ($preview.Length -gt $PreviewLength) {
            $preview = $preview.Substring(0, $PreviewLength) + "..."
        }
        Write-Host $preview -ForegroundColor DarkGray
        
        # Show menu
        Write-Host "`n----------------------------------------" -ForegroundColor DarkCyan
        $threadLabel = if ($showFullThread) { "[T]hread: full" } else { "[T]hread: latest" }
        $menuRows = @(
            @(
                @{ Label = "[V]iew message"; Color = "Green" },
                @{ Label = "[R]eply"; Color = "Green" },
                @{ Label = "[F]orward"; Color = "Green" }
            ),
            @(
                @{ Label = "[M]ark read"; Color = "Green" },
                @{ Label = "[A]rchive"; Color = "Red" },
                @{ Label = "[S]ave Attachments"; Color = "Magenta" }
            ),
            @(
                @{ Label = "[G] Flag"; Color = "Red" },
                @{ Label = "[C]ategory"; Color = "Blue" },
                @{ Label = "[N]ext"; Color = "Yellow" }
            ),
            @(
                @{ Label = "[P]revious"; Color = "Yellow" },
                @{ Label = "[J]ump"; Color = "Yellow" },
                @{ Label = $threadLabel; Color = "Yellow" }
            ),
            @(
                @{ Label = "[B] Add to contacts"; Color = "Cyan" },
                @{ Label = "[Q]uit"; Color = "Gray" },
                @{ Label = ""; Color = "Gray" }
            )
        )

        foreach ($row in $menuRows) {
            foreach ($item in $row) {
                $text = $item.Label.PadRight(18)
                Write-Host $text -NoNewline -ForegroundColor $item.Color
            }
            Write-Host ""
        }
        Write-Host "----------------------------------------" -ForegroundColor DarkCyan

        $choice = Read-Host "`nYour choice"

        switch ($choice.ToLower()) {
            'v' {
                # View full message
                Write-Host ""
                Read-OutlookMessage -Index $currentIndex -KeepLinks:$KeepLinks -ShowHistory:$showFullThread

                Write-Host "`nPress any key to return to menu..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            'r' {
                # Reply
                Send-OutlookReply -Index $currentIndex

                # Always refresh after replying since the message is marked as read
                & $refreshAction -SilentRefresh

                # Stay at current index - the next message will now be at this position
                # If we replied to the last message, adjust to the new last message
                if ($currentIndex -ge $script:CachedMessages.Count) {
                    $currentIndex = [Math]::Max(0, $script:CachedMessages.Count - 1)
                }

                if ($script:CachedMessages.Count -eq 0) {
                    Write-Host "`nNo more messages! Exiting." -ForegroundColor Green
                    return
                }
            }
            'f' {
                # Forward
                Send-ForwardOutlookEmail -Index $currentIndex

                if (-not $useCacheOnly) {
                    & $refreshAction -SilentRefresh
                }

                # Adjust index if needed
                if ($currentIndex -ge $script:CachedMessages.Count) {
                    $currentIndex = [Math]::Max(0, $script:CachedMessages.Count - 1)
                }

                if ($script:CachedMessages.Count -eq 0) {
                    Write-Host "`nNo more messages! Exiting." -ForegroundColor Green
                    return
                }
            }
            'm' {
                # Mark as read
                if ($msg.UnRead) {
                    $msg.UnRead = $false
                    $msg.Save()
                    Write-Host "`nMarked as read." -ForegroundColor Green

                    if (-not $useCacheOnly) {
                        & $refreshAction -SilentRefresh
                    }

                    # Adjust index if needed
                    if ($currentIndex -ge $script:CachedMessages.Count) {
                        $currentIndex = [Math]::Max(0, $script:CachedMessages.Count - 1)
                    }

                    if ($script:CachedMessages.Count -eq 0) {
                        Write-Host "`nNo more messages! Exiting." -ForegroundColor Green
                        return
                    }
                } else {
                    Write-Host "`nAlready marked as read." -ForegroundColor Yellow
                }
            }
            'a' {
                # Archive
                Write-Host ""
                Move-OutlookMessageToArchive -Index $currentIndex

                # Always refresh after archiving since the message is no longer in the inbox
                & $refreshAction -SilentRefresh

                # Stay at current index - the next message will now be at this position
                # If we archived the last message, adjust to the new last message
                if ($currentIndex -ge $script:CachedMessages.Count) {
                    $currentIndex = [Math]::Max(0, $script:CachedMessages.Count - 1)
                }

                if ($script:CachedMessages.Count -eq 0) {
                    Write-Host "`nNo more messages! Exiting." -ForegroundColor Green
                    return
                }
            }
            's' {
                # Save attachments
                $meaningfulAttCount = Get-MeaningfulAttachmentCount -Attachments $msg.Attachments
                if ($meaningfulAttCount -eq 0) {
                    Write-Host "`nNo attachments to save." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                } else {
                    Write-Host ""
                    $targetPath = Read-Host "Directory (press Enter for current folder)"
                    if ([string]::IsNullOrWhiteSpace($targetPath)) {
                        $targetPath = "."
                    }
                    Save-OutlookMessageAttachments -Message $msg -AttachmentPath $targetPath
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
            }
            'g' {
                # Flag message
                Write-Host "`nDo you want to add a comment to the flag? (optional)" -ForegroundColor Gray
                $flagComment = Read-Host "Comment (press Enter to skip)"

                Write-Host "Do you want to set a due date? (optional)" -ForegroundColor Gray
                $dueDateInput = Read-Host "Due date (e.g., '2025-12-31' or press Enter to skip)"

                $flagParams = @{
                    Index = @($currentIndex)
                }

                if (![string]::IsNullOrWhiteSpace($flagComment)) {
                    $flagParams['Comment'] = $flagComment
                }

                if (![string]::IsNullOrWhiteSpace($dueDateInput)) {
                    try {
                        $dueDate = [DateTime]::Parse($dueDateInput)
                        $flagParams['DueBy'] = $dueDate
                    } catch {
                        Write-Host "Invalid date format. Continuing without due date." -ForegroundColor Yellow
                    }
                }

                Add-OutlookFlag @flagParams
                Write-Host "`nMessage flagged." -ForegroundColor Green

                # Refresh to show updated flag status
                $msg = $script:CachedMessages[$currentIndex]
            }
            'c' {
                # Add category/tag
                Write-Host ""
                $selectedCategory = Select-OutlookCategory

                if ($null -ne $selectedCategory) {
                    Add-OutlookCategory -Index @($currentIndex) -Category $selectedCategory
                    Write-Host "`nCategory '$selectedCategory' added to message." -ForegroundColor Green

                    # Refresh to show updated category
                    $msg = $script:CachedMessages[$currentIndex]
                } else {
                    Write-Host "`nCategory assignment cancelled." -ForegroundColor Yellow
                }
            }
            'n' {
                # Next message
                if ($currentIndex -lt $script:CachedMessages.Count - 1) {
                    $currentIndex++
                } else {
                    Write-Host "`nAlready at last message." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }
            'p' {
                # Previous message
                if ($currentIndex -gt 0) {
                    $currentIndex--
                } else {
                    Write-Host "`nAlready at first message." -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }
            'j' {
                # Jump to message
                $jumpTo = Read-Host "Enter message number (1-$($script:CachedMessages.Count))"
                try {
                    $jumpIndex = [int]$jumpTo - 1
                    if ($jumpIndex -ge 0 -and $jumpIndex -lt $script:CachedMessages.Count) {
                        $currentIndex = $jumpIndex
                    } else {
                        Write-Host "Invalid message number." -ForegroundColor Red
                        Start-Sleep -Seconds 1
                    }
                } catch {
                    Write-Host "Invalid input." -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
            }
            't' {
                $showFullThread = -not $showFullThread
                if ($showFullThread) {
                    Write-Host "`nThread view set to show full history." -ForegroundColor Gray
                } else {
                    Write-Host "`nThread view set to latest message only." -ForegroundColor Gray
                }
                Start-Sleep -Milliseconds 800
            }
            'b' {
                # Add sender to contacts
                $senderEmail = Get-SmtpAddress -MailItem $msg
                $senderInContacts = Test-OutlookContactExists -Email $senderEmail

                if ($senderInContacts) {
                    Write-Host "`n$($msg.SenderName) <$senderEmail> is already in your contacts." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                } else {
                    Write-Host ""
                    # Parse sender name into first/last
                    $firstName = ""
                    $lastName = ""
                    if ($msg.SenderName -and $msg.SenderName -ne $senderEmail) {
                        $nameParts = $msg.SenderName -split " ", 2
                        $firstName = $nameParts[0]
                        if ($nameParts.Count -gt 1) {
                            $lastName = $nameParts[1]
                        }
                    }
                    Add-OutlookContact -Email $senderEmail -FirstName $firstName -LastName $lastName

                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
            }
            'q' {
                Write-Host "`nExiting interactive mode. Goodbye!" -ForegroundColor Green
                return
            }
            default {
                Write-Host "`nInvalid choice. Try again." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    }
}

function Add-OutlookFlag {
    <#
    .SYNOPSIS
    Adds a follow-up flag to one or more cached Outlook messages by index.

    .PARAMETER Index
    The index number(s) from Get-OutlookInbox. Can be a single index or array.

    .PARAMETER Comment
    Optional flag text/comment (e.g., "Follow up", "Review", etc.).

    .PARAMETER DueBy
    Optional due date/time for the flag (DateTime). If provided, a reminder is also set for this time.

    .PARAMETER Reminder
    If specified without DueBy, a reminder will be set for 24 hours from now.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [int[]]$Index,

        [string]$Comment,

        [Nullable[datetime]]$DueBy,

        [switch]$Reminder
    )

    if ($script:CachedMessages.Count -eq 0) {
        Write-Host "No cached messages. Run Get-OutlookInbox first." -ForegroundColor Yellow
        return
    }

    $errors = @()
    $updated = 0

    foreach ($i in $Index) {
        if ($i -lt 0 -or $i -ge $script:CachedMessages.Count) {
            $errors += "Index $i is out of range (valid: 0-$($script:CachedMessages.Count - 1))"
            continue
        }

        try {
            $msg = $script:CachedMessages[$i]

            # Set as flagged
            $msg.FlagStatus = 2  # 2 = Flagged

            if ($Comment) {
                $msg.FlagRequest = $Comment
            }

            if ($DueBy) {
                $msg.FlagDueBy = [datetime]$DueBy
                $msg.ReminderSet = $true
                $msg.ReminderTime = [datetime]$DueBy
            } elseif ($Reminder) {
                # Set a default reminder if user asked for one but didn't provide DueBy
                $defaultReminder = (Get-Date).AddDays(1)
                $msg.ReminderSet = $true
                $msg.ReminderTime = $defaultReminder
            }

            $msg.Save()

            Write-Host "[$i] Flagged: " -NoNewline -ForegroundColor Green
            Write-Host $msg.Subject -ForegroundColor White
            $updated++
        }
        catch {
            $errors += "Failed to flag message at index ${i}: $_"
        }
    }

    Write-Host "`nFlagged $updated message(s)" -ForegroundColor Cyan
    if ($errors.Count -gt 0) {
        Write-Host "`nErrors:" -ForegroundColor Red
        $errors | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    }
}

function Add-OutlookCategory {
    <#
    .SYNOPSIS
    Adds one or more categories to one or more cached Outlook messages by index.

    .PARAMETER Index
    The index number(s) from Get-OutlookInbox. Can be a single index or array.

    .PARAMETER Category
    One or more category names to add to the message(s). Categories that do not exist
    will be created automatically with no color.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [int[]]$Index,

        [Parameter(Mandatory=$true)]
        [string[]]$Category
    )

    if ($script:CachedMessages.Count -eq 0) {
        Write-Host "No cached messages. Run Get-OutlookInbox first." -ForegroundColor Yellow
        return
    }

    try {
        # Ensure categories exist in Outlook master list
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $masterCategories = $namespace.Categories

        $newCategoriesAdded = $false
        foreach ($cat in $Category) {
            $name = $cat.Trim()
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $exists = $false
            foreach ($mc in $masterCategories) {
                if ($mc.Name -eq $name) { $exists = $true; break }
            }
            if (-not $exists) {
                # Color 0 = None
                $null = $masterCategories.Add($name, 0)
                $newCategoriesAdded = $true
            }
        }

        # Refresh color cache if new categories were added
        if ($newCategoriesAdded) {
            $script:CategoryColorMap = $null
        }

        $errors = @()
        $updated = 0

        foreach ($i in $Index) {
            if ($i -lt 0 -or $i -ge $script:CachedMessages.Count) {
                $errors += "Index $i is out of range (valid: 0-$($script:CachedMessages.Count - 1))"
                continue
            }

            try {
                $msg = $script:CachedMessages[$i]

                $existing = @()
                if (-not [string]::IsNullOrWhiteSpace($msg.Categories)) {
                    $existing = $msg.Categories -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                }

                $combined = @($existing + ($Category | ForEach-Object { $_.Trim() } | Where-Object { $_ })) | Select-Object -Unique
                $msg.Categories = ($combined -join ';')
                $msg.Save()

                Write-Host "[$i] Categorized: " -NoNewline -ForegroundColor Green
                Write-Host "$($msg.Subject)  ->  " -NoNewline -ForegroundColor White
                Write-ColoredCategories -Categories ([string]$msg.Categories) -Prefix "[" -Suffix "]" -NoNewline
                Write-Host ""
                $updated++
            }
            catch {
                $errors += "Failed to categorize message at index ${i}: $_"
            }
        }

        Write-Host "`nUpdated categories on $updated message(s)" -ForegroundColor Cyan
        if ($errors.Count -gt 0) {
            Write-Host "`nErrors:" -ForegroundColor Red
            $errors | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        }
    }
    catch {
        Write-Error "Failed to add categories: $_"
    }
}

function Select-OutlookCategory {
    <#
    .SYNOPSIS
        Displays a numbered list of existing Outlook categories for user selection.

    .DESCRIPTION
        Queries all existing Outlook categories, displays them as a numbered list,
        and prompts the user to select one. Returns the selected category name.

    .EXAMPLE
        $category = Select-OutlookCategory
        if ($category) { Add-OutlookCategory -MessageIndices 0 -Category $category }

    .OUTPUTS
        String - The selected category name, or $null if cancelled
    #>

    try {
        # Get Outlook application and categories
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $categories = $namespace.Categories

        if ($categories.Count -eq 0) {
            Write-Host "No categories found in Outlook." -ForegroundColor Yellow
            Write-Host "You can create a new category by typing a name, or press Enter to cancel." -ForegroundColor Gray
            $newCategory = Read-Host "Category name"
            return if ([string]::IsNullOrWhiteSpace($newCategory)) { $null } else { $newCategory.Trim() }
        }

        # Display categories
        Write-Host "`nAvailable Categories:" -ForegroundColor Cyan
        Write-Host ("=" * 50) -ForegroundColor Gray

        # Get color map for consistent coloring
        $colorMap = Get-CategoryColorMap

        $categoryList = @()
        $index = 1
        foreach ($cat in $categories) {
            $categoryList += $cat.Name
            $consoleColor = if ($colorMap.ContainsKey($cat.Name)) { $colorMap[$cat.Name] } else { "White" }
            Write-Host ("{0,3}. " -f $index) -NoNewline -ForegroundColor Gray
            Write-Host $cat.Name -ForegroundColor $consoleColor
            $index++
        }

        Write-Host ("=" * 50) -ForegroundColor Gray
        Write-Host "Enter number to select, type new category name, or press Enter to cancel" -ForegroundColor Gray

        # Get user input
        $input = Read-Host "Selection"

        if ([string]::IsNullOrWhiteSpace($input)) {
            return $null
        }

        # Check if input is a number
        $selectedIndex = 0
        if ([int]::TryParse($input, [ref]$selectedIndex)) {
            if ($selectedIndex -ge 1 -and $selectedIndex -le $categoryList.Count) {
                return $categoryList[$selectedIndex - 1]
            } else {
                Write-Host "Invalid selection number." -ForegroundColor Red
                return $null
            }
        }

        # Otherwise treat as new category name
        return $input.Trim()
    }
    catch {
        Write-Error "Failed to select category: $_"
        return $null
    }
    finally {
        if ($namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null }
        if ($outlook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null }
    }
}

function Search-OutlookArchive {
    <#
    .SYNOPSIS
    Searches archived emails with required filters to prevent returning too many results.

    .DESCRIPTION
    Searches the Outlook Archive folder with mandatory filtering. At least one of
    -From, -Subject, or -Category must be specified to prevent accidentally
    returning thousands of emails.

    .PARAMETER From
    Filter by sender name or email (partial match). REQUIRED if Subject and Category not provided.

    .PARAMETER Subject
    Filter by subject line (partial match). REQUIRED if From and Category not provided.

    .PARAMETER Category
    Filter by category/tag name (partial match). REQUIRED if From and Subject not provided.

    .PARAMETER DaysBack
    Only search emails from the last N days (default: 90, max: 365).

    .PARAMETER StartDate
    Earliest ReceivedTime to include (inclusive). Cannot be combined with -DaysBack.

    .PARAMETER EndDate
    Latest ReceivedTime to include (inclusive). Cannot be combined with -DaysBack.

    .PARAMETER Limit
    Maximum number of messages to return (default: 50, max: 200).

    .PARAMETER IncludeRead
    Include read messages (default: true for archive search).

    .PARAMETER HasAttachments
    Show only messages with attachments.

    .PARAMETER Importance
    Filter by importance level: High, Normal, or Low.

    .PARAMETER Compact
    Use compact 2-line display format.

    .PARAMETER Plain
    Output plain text without colors (suitable for piping).

    .PARAMETER Oldest
    Sort by oldest first instead of newest.

    .EXAMPLE
    Search-OutlookArchive -From "john" -DaysBack 30
    Searches archived emails from "john" in the last 30 days.

    .EXAMPLE
    Search-OutlookArchive -Subject "quarterly report" -StartDate "2024-01-01" -EndDate "2024-03-31"
    Searches archived emails with "quarterly report" in subject from Q1 2024.
    #>
    [CmdletBinding()]
    param(
        [string]$From,
        [string]$Subject,
        [string]$Category,

        [ValidateRange(1, 365)]
        [int]$DaysBack = 90,

        [datetime]$StartDate,
        [datetime]$EndDate,

        [ValidateRange(1, 200)]
        [int]$Limit = 50,

        [switch]$HasAttachments,

        [ValidateSet("High", "Normal", "Low")]
        [string]$Importance,

        [switch]$Compact,
        [switch]$Plain,
        [switch]$Oldest,
        [switch]$KeepLinks
    )

    # Apply configuration defaults if parameters not explicitly provided
    if (-not $PSBoundParameters.ContainsKey('DaysBack')) {
        $configDaysBack = Get-ConfigValue -Section "Archive" -Key "DaysBack" -Default 90
        $DaysBack = $configDaysBack
    }

    if (-not $PSBoundParameters.ContainsKey('Limit')) {
        $configLimit = Get-ConfigValue -Section "Archive" -Key "Limit" -Default 50
        $Limit = $configLimit
    }

    if (-not $PSBoundParameters.ContainsKey('Compact')) {
        $configCompact = Get-ConfigValue -Section "Archive" -Key "Compact" -Default $false
        if ($configCompact -eq $true) {
            $Compact = $true
        }
    }

    # Validate that at least one search filter is provided
    if ([string]::IsNullOrWhiteSpace($From) -and [string]::IsNullOrWhiteSpace($Subject) -and [string]::IsNullOrWhiteSpace($Category)) {
        Write-Host "Error: At least one search filter is required." -ForegroundColor Red
        Write-Host "Please specify -From, -Subject, or -Category to search archived emails." -ForegroundColor Yellow
        Write-Host "`nExamples:" -ForegroundColor Cyan
        Write-Host "  Search-OutlookArchive -From 'john.doe'" -ForegroundColor Gray
        Write-Host "  Search-OutlookArchive -Subject 'project update'" -ForegroundColor Gray
        Write-Host "  Search-OutlookArchive -Category 'Important' -DaysBack 60" -ForegroundColor Gray
        return
    }

    try {
        $usingAbsoluteDates = $PSBoundParameters.ContainsKey("StartDate") -or $PSBoundParameters.ContainsKey("EndDate")

        if ($usingAbsoluteDates -and $PSBoundParameters.ContainsKey("DaysBack") -and $DaysBack -ne 90) {
            throw "StartDate/EndDate cannot be combined with DaysBack."
        }

        if ($StartDate -and $EndDate -and $StartDate -gt $EndDate) {
            throw "StartDate must be earlier than or equal to EndDate."
        }

        # Calculate date bounds
        $lowerBound = $null
        $upperBound = $null

        if ($usingAbsoluteDates) {
            if ($StartDate) {
                $lowerBound = [datetime]$StartDate
            }
            if ($EndDate) {
                $upperBound = [datetime]$EndDate
                if ($upperBound.TimeOfDay -eq [TimeSpan]::Zero) {
                    $upperBound = $upperBound.AddDays(1).AddTicks(-1)
                }
            }
        } else {
            $referenceNow = Get-Date
            $lowerBound = $referenceNow.AddDays(-$DaysBack)
        }

        # Connect to Outlook
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")

        # Get Archive folder (folder ID 109)
        $archiveFolderId = 109
        try {
            $archiveFolderId = [int][Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderArchive
        } catch {
            # Keep numeric fallback
        }

        $archiveFolder = $null
        try {
            $archiveFolder = $namespace.GetDefaultFolder($archiveFolderId)
        } catch {
            # Fallback: search for Archive folder by name
            $inbox = $namespace.GetDefaultFolder(6)
            $archiveFolder = $inbox.Parent.Folders | Where-Object { $_.Name -eq "Archive" } | Select-Object -First 1

            if (!$archiveFolder) {
                Write-Host "Archive folder not found." -ForegroundColor Red
                return
            }
        }

        # Get messages
        $messages = $archiveFolder.Items
        $messages.Sort("[ReceivedTime]", -not $Oldest)

        # Apply filters
        $filteredMessages = @()
        $scannedCount = 0
        $maxScan = 5000  # Limit how many messages we scan to prevent hanging

        foreach ($msg in $messages) {
            $scannedCount++
            if ($scannedCount -gt $maxScan) {
                Write-Host "Warning: Reached scan limit ($maxScan messages). Consider narrowing your search criteria." -ForegroundColor Yellow
                break
            }

            # Skip if not a mail item
            if ($msg.Class -ne 43) { continue }

            # Filter by date
            if ($lowerBound -and $msg.ReceivedTime -lt $lowerBound) {
                if (-not $Oldest) {
                    # If sorting newest first and we hit old messages, we can stop
                    continue
                }
                continue
            }
            if ($upperBound -and $msg.ReceivedTime -gt $upperBound) { continue }

            # Filter by sender
            if ($From) {
                $senderEmail = Get-SmtpAddress -MailItem $msg
                $senderMatch = $msg.SenderName -like "*$From*" -or $senderEmail -like "*$From*"
                if (!$senderMatch) { continue }
            }

            # Filter by subject
            if ($Subject -and $msg.Subject -notlike "*$Subject*") { continue }

            # Filter by category
            if ($Category -and $msg.Categories -notlike "*$Category*") { continue }

            # Filter by attachments
            if ($HasAttachments -and (Get-MeaningfulAttachmentCount -Attachments $msg.Attachments) -eq 0) { continue }

            # Filter by importance
            if ($Importance) {
                $importanceValue = switch ($Importance) {
                    "Low" { 0 }
                    "Normal" { 1 }
                    "High" { 2 }
                }
                if ($msg.Importance -ne $importanceValue) { continue }
            }

            $filteredMessages += $msg

            # Stop if we've reached the limit
            if ($filteredMessages.Count -ge $Limit) {
                break
            }
        }

        # Cache results
        $script:CachedArchiveMessages = $filteredMessages

        # Display results
        if ($filteredMessages.Count -eq 0) {
            if (!$Plain) {
                Write-Host "`nNo archived messages found matching criteria." -ForegroundColor Yellow
                Write-Host "Scanned $scannedCount message(s) in Archive folder." -ForegroundColor Gray
            }
            return
        }

        # Plain text output
        if ($Plain) {
            $output = @()
            for ($i = 0; $i -lt $filteredMessages.Count; $i++) {
                $msg = $filteredMessages[$i]
                $senderEmail = Get-SmtpAddress -MailItem $msg
                $meaningfulAttCount = Get-MeaningfulAttachmentCount -Attachments $msg.Attachments
                $attachments = if ($meaningfulAttCount -gt 0) { "ATTACHMENTS:$meaningfulAttCount" } else { "" }
                $tags = @($attachments) | Where-Object { $_ -ne "" }
                $tagString = if ($tags.Count -gt 0) { "[" + ($tags -join ",") + "]" } else { "" }
                $output += "[$i] $tagString $($msg.SenderName) <$senderEmail> | $($msg.Subject) | $($msg.ReceivedTime)"
            }
            Write-Output ($output -join "`n")
            return
        }

        # Formatted display
        Write-Host "`nFound $($filteredMessages.Count) archived message(s) (scanned $scannedCount)`n" -ForegroundColor Cyan

        for ($i = 0; $i -lt $filteredMessages.Count; $i++) {
            $msg = $filteredMessages[$i]

            $cleanBody = Clean-EmailBody -Body $msg.Body -KeepLinks $KeepLinks
            $preview = $cleanBody -replace "`r`n", " " -replace "`n", " "
            $previewLength = if ($Compact) { 100 } else { 80 }
            if ($preview.Length -gt $previewLength) {
                $preview = $preview.Substring(0, $previewLength) + "..."
            }

            $meaningfulAttCount = Get-MeaningfulAttachmentCount -Attachments $msg.Attachments
            $attachmentIndicator = if ($meaningfulAttCount -gt 0) { " [ATT:$meaningfulAttCount]" } else { "" }
            $importanceIndicator = if ($msg.Importance -eq 2) { " [!!]" } else { "" }

            $senderEmail = Get-SmtpAddress -MailItem $msg

            if ($Compact) {
                Write-Host "[$i] " -NoNewline -ForegroundColor Green
                Write-Host "[ARCHIVE]" -NoNewline -ForegroundColor DarkCyan
                Write-Host "$attachmentIndicator$importanceIndicator" -NoNewline -ForegroundColor Magenta
                if ($msg.Categories) {
                    Write-ColoredCategories -Categories ([string]$msg.Categories) -Prefix " [" -Suffix "]" -NoNewline
                }
                Write-Host " $($msg.SenderName) - " -NoNewline -ForegroundColor Cyan
                Write-Host $msg.Subject -ForegroundColor White
                Write-Host "    $preview" -ForegroundColor DarkGray
            } else {
                Write-Host "[$i] " -NoNewline -ForegroundColor Green
                Write-Host "[ARCHIVE] " -NoNewline -ForegroundColor DarkCyan
                Write-Host "$($msg.SenderName) <$senderEmail>" -NoNewline -ForegroundColor Cyan
                Write-Host "$attachmentIndicator$importanceIndicator" -NoNewline -ForegroundColor Magenta
                if ($msg.Categories) {
                    Write-ColoredCategories -Categories ([string]$msg.Categories) -Prefix " [" -Suffix "]" -NoNewline
                }
                Write-Host ""
                Write-Host "    Subject: " -NoNewline -ForegroundColor Gray
                Write-Host $msg.Subject -ForegroundColor White
                Write-Host "    Received: " -NoNewline -ForegroundColor Gray
                Write-Host $msg.ReceivedTime -ForegroundColor White
                Write-Host "    Preview: " -NoNewline -ForegroundColor Gray
                Write-Host $preview -ForegroundColor DarkGray
                Write-Host ""
            }
        }

        Write-Host "Use 'Read-OutlookMessage -Index <n> -Archive' to read an archived message`n" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to search archive: $_"
        Write-Host "Make sure Outlook is installed and you're logged in." -ForegroundColor Yellow
    }
}

function Get-CachedArchiveMessages {
    <#
    .SYNOPSIS
    Returns the currently cached messages from the last Search-OutlookArchive call.

    .DESCRIPTION
    Helper function to access cached archive messages for debugging or external use.
    #>
    return $script:CachedArchiveMessages
}

function Format-CalendarItem {
    <#
    .SYNOPSIS
    Formats a calendar appointment item for display.

    .DESCRIPTION
    Helper function to consistently format calendar items across different views.

    .PARAMETER Appointment
    The Outlook appointment COM object.

    .PARAMETER Compact
    Use compact single-line format.

    .PARAMETER IncludeDate
    Include the date in the output (useful for weekly views).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.__ComObject]$Appointment,
        [switch]$Compact,
        [switch]$IncludeDate
    )

    $startTime = $Appointment.Start
    $endTime = $Appointment.End
    $duration = $endTime - $startTime

    # Determine busy status color
    $statusColor = switch ($Appointment.BusyStatus) {
        0 { "Green" }      # Free
        1 { "Cyan" }       # Tentative
        2 { "Red" }        # Busy
        3 { "DarkRed" }    # Out of Office
        4 { "Gray" }       # Working Elsewhere
        default { "White" }
    }

    $statusText = switch ($Appointment.BusyStatus) {
        0 { "Free" }
        1 { "Tentative" }
        2 { "Busy" }
        3 { "OOO" }
        4 { "Remote" }
        default { "" }
    }

    # Format time display
    if ($Appointment.AllDayEvent) {
        $timeDisplay = "All Day"
    } else {
        $timeDisplay = "$($startTime.ToString('h:mm tt')) - $($endTime.ToString('h:mm tt'))"
    }

    # Format duration
    $durationDisplay = ""
    if (-not $Appointment.AllDayEvent) {
        if ($duration.TotalHours -ge 1) {
            $hours = [Math]::Floor($duration.TotalHours)
            $mins = $duration.Minutes
            if ($mins -gt 0) {
                $durationDisplay = "${hours}h ${mins}m"
            } else {
                $durationDisplay = "${hours}h"
            }
        } else {
            $durationDisplay = "$($duration.TotalMinutes)m"
        }
    }

    if ($Compact) {
        # Compact single-line format
        $datePrefix = ""
        if ($IncludeDate) {
            $datePrefix = $startTime.ToString("ddd MM/dd") + " "
        }
        Write-Host "  " -NoNewline
        Write-Host $datePrefix -NoNewline -ForegroundColor DarkGray
        Write-Host $timeDisplay.PadRight(20) -NoNewline -ForegroundColor $statusColor
        Write-Host $Appointment.Subject -NoNewline -ForegroundColor White
        if ($Appointment.Location) {
            Write-Host " @ $($Appointment.Location)" -NoNewline -ForegroundColor DarkGray
        }
        Write-Host ""
    } else {
        # Detailed format
        if ($IncludeDate) {
            Write-Host "  Date: " -NoNewline -ForegroundColor Gray
            Write-Host $startTime.ToString("dddd, MMMM dd, yyyy") -ForegroundColor White
        }
        Write-Host "  Time: " -NoNewline -ForegroundColor Gray
        Write-Host $timeDisplay -NoNewline -ForegroundColor $statusColor
        if ($durationDisplay) {
            Write-Host " ($durationDisplay)" -NoNewline -ForegroundColor DarkGray
        }
        Write-Host ""

        Write-Host "  Subject: " -NoNewline -ForegroundColor Gray
        Write-Host $Appointment.Subject -ForegroundColor White

        if ($Appointment.Location) {
            Write-Host "  Location: " -NoNewline -ForegroundColor Gray
            Write-Host $Appointment.Location -ForegroundColor Cyan
        }

        if ($Appointment.Organizer -and $Appointment.Organizer -ne $Appointment.Subject) {
            Write-Host "  Organizer: " -NoNewline -ForegroundColor Gray
            Write-Host $Appointment.Organizer -ForegroundColor DarkCyan
        }

        # Show required attendees count if it's a meeting
        if ($Appointment.MeetingStatus -ne 0 -and $Appointment.RequiredAttendees) {
            $attendeeCount = ($Appointment.RequiredAttendees -split ";").Count
            Write-Host "  Attendees: " -NoNewline -ForegroundColor Gray
            Write-Host "$attendeeCount required" -ForegroundColor DarkGray
        }

        Write-Host "  Status: " -NoNewline -ForegroundColor Gray
        Write-Host $statusText -ForegroundColor $statusColor

        Write-Host ""
    }
}

function Get-OutlookCalendarDay {
    <#
    .SYNOPSIS
    Shows calendar appointments for a specific day.

    .DESCRIPTION
    Displays all appointments for today or a specified date.

    .PARAMETER Date
    The date to show appointments for. Defaults to today.

    .PARAMETER Compact
    Use compact single-line display format.

    .PARAMETER Plain
    Output plain text without colors (suitable for piping).

    .EXAMPLE
    Get-OutlookCalendarDay
    Shows today's appointments.

    .EXAMPLE
    Get-OutlookCalendarDay -Date "2025-12-25"
    Shows appointments for December 25, 2025.
    #>
    [CmdletBinding()]
    param(
        [datetime]$Date = (Get-Date),
        [switch]$Compact,
        [switch]$Plain
    )

    # Apply configuration defaults if parameters not explicitly provided
    if (-not $PSBoundParameters.ContainsKey('Compact')) {
        $configCompact = Get-ConfigValue -Section "Calendar" -Key "Compact" -Default $false
        if ($configCompact -eq $true) {
            $Compact = $true
        }
    }

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $calendar = $namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

        # Set date range for the day
        $startOfDay = $Date.Date
        $endOfDay = $startOfDay.AddDays(1).AddTicks(-1)

        # Get appointments using restriction
        $filter = "[Start] >= '$($startOfDay.ToString("g"))' AND [Start] < '$($endOfDay.AddTicks(1).ToString("g"))'"
        $appointments = $calendar.Items
        $appointments.Sort("[Start]")
        $appointments.IncludeRecurrences = $true
        $filteredAppts = $appointments.Restrict($filter)

        # Collect appointments into array
        $apptList = @()
        $currentTime = Get-Date
        foreach ($appt in $filteredAppts) {
            if ($appt.Start -ge $startOfDay -and $appt.Start -lt $endOfDay.AddTicks(1)) {
                # Filter out events that have already ended
                if ($appt.AllDayEvent) {
                    # For all-day events, check if the date is today or future
                    if ($appt.Start.Date -ge $currentTime.Date) {
                        $apptList += $appt
                    }
                } else {
                    # For regular events, check if EndTime hasn't passed
                    if ($appt.End -gt $currentTime) {
                        $apptList += $appt
                    }
                }
            }
        }

        if ($Plain) {
            $output = @()
            $output += "Calendar for $($Date.ToString('dddd, MMMM dd, yyyy'))"
            $output += "=" * 50

            if ($apptList.Count -eq 0) {
                $output += "No appointments scheduled."
            } else {
                foreach ($appt in $apptList) {
                    $timeStr = if ($appt.AllDayEvent) { "All Day" } else { "$($appt.Start.ToString('h:mm tt')) - $($appt.End.ToString('h:mm tt'))" }
                    $locStr = if ($appt.Location) { " @ $($appt.Location)" } else { "" }
                    $output += "$timeStr | $($appt.Subject)$locStr"
                }
            }
            Write-Output ($output -join "`n")
            return
        }

        # Formatted output
        Write-Host "`n" -NoNewline
        Write-Host ("=" * 60) -ForegroundColor Cyan
        Write-Host "  Calendar for $($Date.ToString('dddd, MMMM dd, yyyy'))" -ForegroundColor Cyan
        Write-Host ("=" * 60) -ForegroundColor Cyan
        Write-Host ""

        if ($apptList.Count -eq 0) {
            Write-Host "  No appointments scheduled for this day." -ForegroundColor Gray
            Write-Host ""
        } else {
            foreach ($appt in $apptList) {
                Format-CalendarItem -Appointment $appt -Compact:$Compact
            }
        }
    }
    catch {
        Write-Error "Failed to access calendar: $_"
        Write-Host "Make sure Outlook is installed and you're logged in." -ForegroundColor Yellow
    }
}

function Get-OutlookCalendarWeek {
    <#
    .SYNOPSIS
    Shows calendar appointments for a week.

    .DESCRIPTION
    Displays all appointments for the current week or a week starting from a specified date.

    .PARAMETER StartDate
    The start date of the week to display. Defaults to today.

    .PARAMETER Days
    Number of days to show (default: 7).

    .PARAMETER Compact
    Use compact display format.

    .PARAMETER Plain
    Output plain text without colors (suitable for piping).

    .PARAMETER WorkWeek
    Show only Monday through Friday (5 days).

    .EXAMPLE
    Get-OutlookCalendarWeek
    Shows appointments for the next 7 days.

    .EXAMPLE
    Get-OutlookCalendarWeek -WorkWeek
    Shows appointments for the work week (Mon-Fri).

    .EXAMPLE
    Get-OutlookCalendarWeek -StartDate "2025-12-01" -Days 14
    Shows appointments for 2 weeks starting December 1st.
    #>
    [CmdletBinding()]
    param(
        [datetime]$StartDate = (Get-Date),
        [ValidateRange(1, 31)]
        [int]$Days = 7,
        [switch]$Compact,
        [switch]$Plain,
        [switch]$WorkWeek
    )

    # Apply configuration defaults if parameters not explicitly provided
    if (-not $PSBoundParameters.ContainsKey('Days')) {
        $configDays = Get-ConfigValue -Section "Calendar" -Key "Days" -Default 7
        $Days = $configDays
    }

    if (-not $PSBoundParameters.ContainsKey('Compact')) {
        $configCompact = Get-ConfigValue -Section "Calendar" -Key "Compact" -Default $false
        if ($configCompact -eq $true) {
            $Compact = $true
        }
    }

    if (-not $PSBoundParameters.ContainsKey('WorkWeek')) {
        $configWorkWeek = Get-ConfigValue -Section "Calendar" -Key "WorkWeek" -Default $false
        if ($configWorkWeek -eq $true) {
            $WorkWeek = $true
        }
    }

    try {
        if ($WorkWeek) {
            # Adjust to start of work week (Monday)
            while ($StartDate.DayOfWeek -ne [DayOfWeek]::Monday) {
                $StartDate = $StartDate.AddDays(-1)
            }
            $Days = 5
        }

        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $calendar = $namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

        $startOfRange = $StartDate.Date
        $endOfRange = $startOfRange.AddDays($Days)

        # Get appointments
        $filter = "[Start] >= '$($startOfRange.ToString("g"))' AND [Start] < '$($endOfRange.ToString("g"))'"
        $appointments = $calendar.Items
        $appointments.Sort("[Start]")
        $appointments.IncludeRecurrences = $true
        $filteredAppts = $appointments.Restrict($filter)

        # Group by day
        $apptsByDay = @{}
        for ($d = 0; $d -lt $Days; $d++) {
            $dayDate = $startOfRange.AddDays($d).Date
            $apptsByDay[$dayDate] = @()
        }

        $currentTime = Get-Date
        foreach ($appt in $filteredAppts) {
            # Filter out events that have already ended
            $includeEvent = $false
            if ($appt.AllDayEvent) {
                # For all-day events, check if the date is today or future
                $includeEvent = ($appt.Start.Date -ge $currentTime.Date)
            } else {
                # For regular events, check if EndTime hasn't passed
                $includeEvent = ($appt.End -gt $currentTime)
            }

            if ($includeEvent) {
                $apptDate = $appt.Start.Date
                if ($apptsByDay.ContainsKey($apptDate)) {
                    $apptsByDay[$apptDate] += $appt
                }
            }
        }

        if ($Plain) {
            $output = @()
            $output += "Calendar Week: $($startOfRange.ToString('MMM dd')) - $($endOfRange.AddDays(-1).ToString('MMM dd, yyyy'))"
            $output += "=" * 60

            for ($d = 0; $d -lt $Days; $d++) {
                $dayDate = $startOfRange.AddDays($d).Date
                $dayAppts = $apptsByDay[$dayDate]
                $output += ""
                $output += $dayDate.ToString("dddd, MMMM dd")
                $output += "-" * 30

                if ($dayAppts.Count -eq 0) {
                    $output += "  No appointments"
                } else {
                    foreach ($appt in $dayAppts) {
                        $timeStr = if ($appt.AllDayEvent) { "All Day" } else { "$($appt.Start.ToString('h:mm tt')) - $($appt.End.ToString('h:mm tt'))" }
                        $locStr = if ($appt.Location) { " @ $($appt.Location)" } else { "" }
                        $output += "  $timeStr | $($appt.Subject)$locStr"
                    }
                }
            }
            Write-Output ($output -join "`n")
            return
        }

        # Formatted output
        Write-Host "`n" -NoNewline
        Write-Host ("=" * 70) -ForegroundColor Cyan
        Write-Host "  Calendar: $($startOfRange.ToString('MMMM dd')) - $($endOfRange.AddDays(-1).ToString('MMMM dd, yyyy'))" -ForegroundColor Cyan
        Write-Host ("=" * 70) -ForegroundColor Cyan

        for ($d = 0; $d -lt $Days; $d++) {
            $dayDate = $startOfRange.AddDays($d).Date
            $dayAppts = $apptsByDay[$dayDate]
            $isToday = $dayDate -eq (Get-Date).Date
            $isWeekend = $dayDate.DayOfWeek -eq [DayOfWeek]::Saturday -or $dayDate.DayOfWeek -eq [DayOfWeek]::Sunday

            Write-Host ""
            $dayHeader = $dayDate.ToString("dddd, MMMM dd")
            if ($isToday) {
                Write-Host "  $dayHeader (TODAY)" -ForegroundColor Yellow
            } elseif ($isWeekend) {
                Write-Host "  $dayHeader" -ForegroundColor DarkGray
            } else {
                Write-Host "  $dayHeader" -ForegroundColor White
            }
            Write-Host ("  " + "-" * 40) -ForegroundColor DarkGray

            if ($dayAppts.Count -eq 0) {
                Write-Host "    No appointments" -ForegroundColor DarkGray
            } else {
                foreach ($appt in $dayAppts) {
                    Format-CalendarItem -Appointment $appt -Compact:$true
                }
            }
        }
        Write-Host ""
    }
    catch {
        Write-Error "Failed to access calendar: $_"
        Write-Host "Make sure Outlook is installed and you're logged in." -ForegroundColor Yellow
    }
}

function New-OutlookCalendarEvent {
    <#
    .SYNOPSIS
    Creates a new calendar event/appointment.

    .DESCRIPTION
    Creates a new Outlook calendar event with the specified details.

    .PARAMETER Subject
    The subject/title of the event. Required.

    .PARAMETER Start
    Start date and time. Can be a DateTime or string (e.g., "2025-12-25 14:00").

    .PARAMETER End
    End date and time. If not specified, defaults to 1 hour after start.

    .PARAMETER Duration
    Duration in minutes. Alternative to specifying End time.

    .PARAMETER Location
    Location of the event.

    .PARAMETER Body
    Description/body of the event.

    .PARAMETER AllDay
    Create an all-day event.

    .PARAMETER Reminder
    Reminder time in minutes before the event (default: 15).

    .PARAMETER NoReminder
    Disable reminder for this event.

    .PARAMETER BusyStatus
    Show As status: Free, Tentative, Busy, OutOfOffice, WorkingElsewhere.

    .PARAMETER Attendees
    Email addresses of required attendees (comma-separated or array).

    .PARAMETER OptionalAttendees
    Email addresses of optional attendees (comma-separated or array).

    .PARAMETER SendInvites
    Send meeting invitations to attendees (requires Attendees parameter).

    .EXAMPLE
    New-OutlookCalendarEvent -Subject "Team Meeting" -Start "2025-12-01 10:00" -Duration 60 -Location "Conference Room A"

    .EXAMPLE
    New-OutlookCalendarEvent -Subject "Company Holiday" -Start "2025-12-25" -AllDay

    .EXAMPLE
    New-OutlookCalendarEvent -Subject "Project Review" -Start "2025-12-15 14:00" -End "2025-12-15 15:30" -Attendees "john@example.com,jane@example.com" -SendInvites
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Subject,

        [Parameter(Mandatory=$true)]
        [datetime]$Start,

        [datetime]$End,

        [int]$Duration,

        [string]$Location,

        [string]$Body,

        [switch]$AllDay,

        [int]$Reminder = 15,

        [switch]$NoReminder,

        [ValidateSet("Free", "Tentative", "Busy", "OutOfOffice", "WorkingElsewhere")]
        [string]$BusyStatus = "Busy",

        [string[]]$Attendees,

        [string[]]$OptionalAttendees,

        [switch]$SendInvites
    )

    try {
        $outlook = New-Object -ComObject Outlook.Application

        # Create appointment
        $appt = $outlook.CreateItem(1)  # 1 = olAppointmentItem

        $appt.Subject = $Subject

        if ($AllDay) {
            $appt.AllDayEvent = $true
            $appt.Start = $Start.Date
            $appt.End = $Start.Date.AddDays(1)
        } else {
            $appt.Start = $Start

            if ($End) {
                $appt.End = $End
            } elseif ($Duration) {
                $appt.End = $Start.AddMinutes($Duration)
            } else {
                # Default to 1 hour
                $appt.End = $Start.AddHours(1)
            }
        }

        if ($Location) {
            $appt.Location = $Location
        }

        if ($Body) {
            $appt.Body = $Body
        }

        # Set busy status
        $appt.BusyStatus = switch ($BusyStatus) {
            "Free" { 0 }
            "Tentative" { 1 }
            "Busy" { 2 }
            "OutOfOffice" { 3 }
            "WorkingElsewhere" { 4 }
        }

        # Set reminder
        if ($NoReminder) {
            $appt.ReminderSet = $false
        } else {
            $appt.ReminderSet = $true
            $appt.ReminderMinutesBeforeStart = $Reminder
        }

        # Add attendees if specified (converts to meeting request)
        if ($Attendees -or $OptionalAttendees) {
            if ($Attendees) {
                foreach ($email in $Attendees) {
                    $recipient = $appt.Recipients.Add($email.Trim())
                    $recipient.Type = 1  # Required
                }
            }
            if ($OptionalAttendees) {
                foreach ($email in $OptionalAttendees) {
                    $recipient = $appt.Recipients.Add($email.Trim())
                    $recipient.Type = 2  # Optional
                }
            }
            $appt.MeetingStatus = 1  # olMeeting
        }

        # Show preview
        Write-Host "`nNew Calendar Event Preview:" -ForegroundColor Cyan
        Write-Host ("=" * 50) -ForegroundColor Cyan
        Write-Host "Subject: " -NoNewline -ForegroundColor Gray
        Write-Host $Subject -ForegroundColor White

        if ($AllDay) {
            Write-Host "Date: " -NoNewline -ForegroundColor Gray
            Write-Host "$($Start.ToString('dddd, MMMM dd, yyyy')) (All Day)" -ForegroundColor White
        } else {
            Write-Host "Start: " -NoNewline -ForegroundColor Gray
            Write-Host $appt.Start.ToString("dddd, MMMM dd, yyyy HH:mm") -ForegroundColor White
            Write-Host "End: " -NoNewline -ForegroundColor Gray
            Write-Host $appt.End.ToString("dddd, MMMM dd, yyyy HH:mm") -ForegroundColor White
        }

        if ($Location) {
            Write-Host "Location: " -NoNewline -ForegroundColor Gray
            Write-Host $Location -ForegroundColor Cyan
        }

        Write-Host "Show As: " -NoNewline -ForegroundColor Gray
        Write-Host $BusyStatus -ForegroundColor Yellow

        if ($Attendees) {
            Write-Host "Attendees: " -NoNewline -ForegroundColor Gray
            Write-Host ($Attendees -join ", ") -ForegroundColor White
        }
        if ($OptionalAttendees) {
            Write-Host "Optional: " -NoNewline -ForegroundColor Gray
            Write-Host ($OptionalAttendees -join ", ") -ForegroundColor DarkGray
        }

        Write-Host ("=" * 50) -ForegroundColor Cyan

        # Confirm
        $confirm = Read-Host "`nCreate this event? (y/N)"

        if ($confirm -eq 'y' -or $confirm -eq 'Y') {
            if ($SendInvites -and ($Attendees -or $OptionalAttendees)) {
                $appt.Send()
                Write-Host "Meeting created and invitations sent!" -ForegroundColor Green
            } else {
                $appt.Save()
                Write-Host "Calendar event created!" -ForegroundColor Green
            }
        } else {
            Write-Host "Event creation cancelled." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Failed to create calendar event: $_"
    }
}

function Add-OutlookContact {
    <#
    .SYNOPSIS
    Creates a new Outlook contact with prompts for additional details.

    .DESCRIPTION
    Creates a new contact in Outlook. If only email is provided, prompts for
    additional details like name, phone, company, and notes.

    .PARAMETER Email
    The email address for the contact. Required.

    .PARAMETER FirstName
    First name of the contact.

    .PARAMETER LastName
    Last name of the contact.

    .PARAMETER FullName
    Full name (used if FirstName/LastName not provided separately).

    .PARAMETER Phone
    Phone number.

    .PARAMETER Company
    Company/organization name.

    .PARAMETER JobTitle
    Job title/position.

    .PARAMETER Notes
    Additional notes about the contact.

    .PARAMETER NoPrompt
    Skip prompts and create contact with provided information only.

    .EXAMPLE
    Add-OutlookContact -Email "john.doe@example.com"
    Creates a new contact, prompting for additional details.

    .EXAMPLE
    Add-OutlookContact -Email "jane@example.com" -FirstName "Jane" -LastName "Smith" -Company "Acme Corp" -NoPrompt
    Creates a contact without prompting for additional details.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Email,

        [string]$FirstName,

        [string]$LastName,

        [string]$FullName,

        [string]$Phone,

        [string]$Company,

        [string]$JobTitle,

        [string]$Notes,

        [switch]$NoPrompt
    )

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")

        # Check if contact already exists
        $contacts = $namespace.GetDefaultFolder(10)  # 10 = olFolderContacts
        $existingContact = $contacts.Items | Where-Object {
            $_.Email1Address -eq $Email -or
            $_.Email2Address -eq $Email -or
            $_.Email3Address -eq $Email
        } | Select-Object -First 1

        if ($existingContact) {
            Write-Host "Contact already exists: $($existingContact.FullName) <$Email>" -ForegroundColor Yellow
            return $false
        }

        # Prompt for additional details if not using NoPrompt
        if (-not $NoPrompt) {
            Write-Host "`nAdding new contact: $Email" -ForegroundColor Cyan
            Write-Host ("=" * 50) -ForegroundColor Cyan
            Write-Host "Enter details (press Enter to skip):`n" -ForegroundColor Gray

            if ([string]::IsNullOrWhiteSpace($FirstName)) {
                $FirstName = Read-Host "First Name"
            }

            if ([string]::IsNullOrWhiteSpace($LastName)) {
                $LastName = Read-Host "Last Name"
            }

            if ([string]::IsNullOrWhiteSpace($Phone)) {
                $Phone = Read-Host "Phone"
            }

            if ([string]::IsNullOrWhiteSpace($Company)) {
                $Company = Read-Host "Company"
            }

            if ([string]::IsNullOrWhiteSpace($JobTitle)) {
                $JobTitle = Read-Host "Job Title"
            }

            if ([string]::IsNullOrWhiteSpace($Notes)) {
                $Notes = Read-Host "Notes"
            }
        }

        # Create contact
        $contact = $outlook.CreateItem(2)  # 2 = olContactItem

        $contact.Email1Address = $Email

        if ($FirstName) { $contact.FirstName = $FirstName }
        if ($LastName) { $contact.LastName = $LastName }
        if ($FullName -and -not $FirstName -and -not $LastName) {
            $contact.FullName = $FullName
        }
        if ($Phone) { $contact.BusinessTelephoneNumber = $Phone }
        if ($Company) { $contact.CompanyName = $Company }
        if ($JobTitle) { $contact.JobTitle = $JobTitle }
        if ($Notes) { $contact.Body = $Notes }

        # Generate file as if needed
        if (-not $contact.FileAs) {
            if ($LastName -and $FirstName) {
                $contact.FileAs = "$LastName, $FirstName"
            } elseif ($FullName) {
                $contact.FileAs = $FullName
            } else {
                $contact.FileAs = $Email
            }
        }

        # Show preview
        Write-Host "`nContact Preview:" -ForegroundColor Cyan
        Write-Host ("=" * 40) -ForegroundColor Gray
        if ($contact.FullName) {
            Write-Host "Name: $($contact.FullName)" -ForegroundColor White
        }
        Write-Host "Email: $Email" -ForegroundColor White
        if ($Phone) { Write-Host "Phone: $Phone" -ForegroundColor White }
        if ($Company) { Write-Host "Company: $Company" -ForegroundColor White }
        if ($JobTitle) { Write-Host "Title: $JobTitle" -ForegroundColor White }
        if ($Notes) { Write-Host "Notes: $Notes" -ForegroundColor DarkGray }
        Write-Host ("=" * 40) -ForegroundColor Gray

        # Confirm unless NoPrompt
        $shouldSave = $true
        if (-not $NoPrompt) {
            $confirm = Read-Host "`nSave this contact? (Y/n)"
            $shouldSave = ($confirm -eq '' -or $confirm -eq 'y' -or $confirm -eq 'Y')
        }

        if ($shouldSave) {
            $contact.Save()
            Write-Host "Contact saved!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "Contact not saved." -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        Write-Error "Failed to create contact: $_"
        return $false
    }
}

function Get-OutlookContacts {
    <#
    .SYNOPSIS
    Lists or searches Outlook contacts.

    .DESCRIPTION
    Displays contacts from the Outlook contacts folder. Can search by name or email.

    .PARAMETER Search
    Search term to filter contacts (partial match on name or email).

    .PARAMETER Limit
    Maximum number of contacts to display (default: 50).

    .PARAMETER Plain
    Output plain text without colors.

    .EXAMPLE
    Get-OutlookContacts
    Lists up to 50 contacts.

    .EXAMPLE
    Get-OutlookContacts -Search "john"
    Searches for contacts with "john" in name or email.
    #>
    [CmdletBinding()]
    param(
        [string]$Search,

        [int]$Limit = 50,

        [switch]$Plain
    )

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $contacts = $namespace.GetDefaultFolder(10)  # 10 = olFolderContacts

        $items = $contacts.Items
        $items.Sort("[FileAs]")

        $filteredContacts = @()
        $count = 0

        foreach ($contact in $items) {
            # Skip if not a contact item
            if ($contact.Class -ne 40) { continue }  # 40 = olContact

            # Apply search filter
            if ($Search) {
                $match = $contact.FullName -like "*$Search*" -or
                         $contact.Email1Address -like "*$Search*" -or
                         $contact.Email2Address -like "*$Search*" -or
                         $contact.CompanyName -like "*$Search*"
                if (-not $match) { continue }
            }

            $filteredContacts += $contact
            $count++

            if ($count -ge $Limit) { break }
        }

        if ($filteredContacts.Count -eq 0) {
            if ($Search) {
                Write-Host "No contacts found matching '$Search'." -ForegroundColor Yellow
            } else {
                Write-Host "No contacts found." -ForegroundColor Yellow
            }
            return
        }

        if ($Plain) {
            $output = @()
            $output += "Contacts ($($filteredContacts.Count) found)"
            $output += "=" * 50
            foreach ($c in $filteredContacts) {
                $email = if ($c.Email1Address) { $c.Email1Address } else { "(no email)" }
                $company = if ($c.CompanyName) { " @ $($c.CompanyName)" } else { "" }
                $output += "$($c.FullName) <$email>$company"
            }
            Write-Output ($output -join "`n")
            return
        }

        # Formatted output
        Write-Host "`nContacts ($($filteredContacts.Count) found)`n" -ForegroundColor Cyan

        $index = 0
        foreach ($c in $filteredContacts) {
            Write-Host "[$index] " -NoNewline -ForegroundColor Green
            Write-Host $c.FullName -NoNewline -ForegroundColor White

            if ($c.Email1Address) {
                Write-Host " <$($c.Email1Address)>" -NoNewline -ForegroundColor Cyan
            }

            if ($c.CompanyName) {
                Write-Host " @ $($c.CompanyName)" -NoNewline -ForegroundColor DarkGray
            }

            Write-Host ""

            if ($c.BusinessTelephoneNumber) {
                Write-Host "    Phone: $($c.BusinessTelephoneNumber)" -ForegroundColor DarkGray
            }

            $index++
        }
        Write-Host ""
    }
    catch {
        Write-Error "Failed to retrieve contacts: $_"
    }
}

function Test-OutlookContactExists {
    <#
    .SYNOPSIS
    Checks if a contact with the given email exists.

    .DESCRIPTION
    Internal helper function to check if a contact already exists.

    .PARAMETER Email
    The email address to check.

    .RETURNS
    Boolean - True if contact exists, False otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Email
    )

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $contacts = $namespace.GetDefaultFolder(10)

        foreach ($contact in $contacts.Items) {
            if ($contact.Class -ne 40) { continue }
            if ($contact.Email1Address -eq $Email -or
                $contact.Email2Address -eq $Email -or
                $contact.Email3Address -eq $Email) {
                return $true
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

function Sync-OutlookContactsFromReplies {
    <#
    .SYNOPSIS
    Syncs contacts from sent email replies.

    .DESCRIPTION
    Scans Sent Items folder for unique recipients and offers to add them as contacts
    if they don't already exist.

    .PARAMETER DaysBack
    Number of days to scan in Sent Items (default: 30).

    .PARAMETER WhatIf
    Preview mode - shows what contacts would be added without actually adding them.

    .PARAMETER Interactive
    Prompt to confirm each contact before adding.

    .PARAMETER AutoAdd
    Automatically add all new contacts without prompting.

    .EXAMPLE
    Sync-OutlookContactsFromReplies -DaysBack 14 -WhatIf
    Preview what contacts would be added from the last 14 days of sent emails.

    .EXAMPLE
    Sync-OutlookContactsFromReplies -Interactive
    Interactively add contacts from the last 30 days of sent emails.
    #>
    [CmdletBinding()]
    param(
        [int]$DaysBack = 30,

        [switch]$WhatIf,

        [switch]$Interactive,

        [switch]$AutoAdd
    )

    # Apply configuration defaults if parameters not explicitly provided
    if (-not $PSBoundParameters.ContainsKey('DaysBack')) {
        $configDaysBack = Get-ConfigValue -Section "Contacts" -Key "DaysBack" -Default 30
        $DaysBack = $configDaysBack
    }

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")

        # Get Sent Items folder (folder ID 5)
        $sentItems = $namespace.GetDefaultFolder(5)

        $cutoffDate = (Get-Date).AddDays(-$DaysBack)

        Write-Host "`nScanning Sent Items from the last $DaysBack days..." -ForegroundColor Cyan

        # Collect unique recipients
        $recipients = @{}
        $scannedCount = 0

        $items = $sentItems.Items
        $items.Sort("[SentOn]", $true)  # Newest first

        foreach ($msg in $items) {
            if ($msg.Class -ne 43) { continue }  # Mail items only
            if ($msg.SentOn -lt $cutoffDate) { break }  # Stop when we hit old messages

            $scannedCount++

            # Get all recipients from To field
            foreach ($recip in $msg.Recipients) {
                $email = $null

                # Try to get SMTP address
                if ($recip.AddressEntry.Type -eq "SMTP") {
                    $email = $recip.Address
                } else {
                    try {
                        $exchUser = $recip.AddressEntry.GetExchangeUser()
                        if ($exchUser) {
                            $email = $exchUser.PrimarySmtpAddress
                        }
                    } catch { }
                }

                if ($email -and -not $recipients.ContainsKey($email.ToLower())) {
                    $recipients[$email.ToLower()] = @{
                        Email = $email
                        Name = $recip.Name
                    }
                }
            }
        }

        Write-Host "Scanned $scannedCount sent message(s)" -ForegroundColor Gray
        Write-Host "Found $($recipients.Count) unique recipient(s)" -ForegroundColor Gray

        if ($recipients.Count -eq 0) {
            Write-Host "No recipients found to process." -ForegroundColor Yellow
            return
        }

        # Check which recipients are not already contacts
        $newContacts = @()
        $existingCount = 0

        Write-Host "`nChecking against existing contacts..." -ForegroundColor Gray

        foreach ($recipInfo in $recipients.Values) {
            $exists = Test-OutlookContactExists -Email $recipInfo.Email
            if (-not $exists) {
                $newContacts += $recipInfo
            } else {
                $existingCount++
            }
        }

        Write-Host "$existingCount recipient(s) already in contacts" -ForegroundColor DarkGray
        Write-Host "$($newContacts.Count) new contact(s) to add`n" -ForegroundColor Cyan

        if ($newContacts.Count -eq 0) {
            Write-Host "All recipients are already in your contacts!" -ForegroundColor Green
            return
        }

        # Process new contacts
        if ($WhatIf) {
            Write-Host "Would add the following contacts:" -ForegroundColor Yellow
            foreach ($c in $newContacts) {
                Write-Host "  - $($c.Name) <$($c.Email)>" -ForegroundColor White
            }
            return
        }

        $addedCount = 0
        $skippedCount = 0

        foreach ($c in $newContacts) {
            Write-Host "`n$($c.Name) <$($c.Email)>" -ForegroundColor White

            $shouldAdd = $true

            if ($Interactive -and -not $AutoAdd) {
                $choice = Read-Host "Add to contacts? (y/N/q to quit)"
                if ($choice -eq 'q' -or $choice -eq 'Q') {
                    Write-Host "Sync cancelled." -ForegroundColor Yellow
                    break
                }
                $shouldAdd = ($choice -eq 'y' -or $choice -eq 'Y')
            }

            if ($shouldAdd) {
                # Parse name into first/last if possible
                $firstName = ""
                $lastName = ""
                if ($c.Name -and $c.Name -ne $c.Email) {
                    $nameParts = $c.Name -split " ", 2
                    $firstName = $nameParts[0]
                    if ($nameParts.Count -gt 1) {
                        $lastName = $nameParts[1]
                    }
                }

                $added = Add-OutlookContact -Email $c.Email -FirstName $firstName -LastName $lastName -NoPrompt
                if ($added) {
                    $addedCount++
                }
            } else {
                $skippedCount++
            }
        }

        Write-Host "`n" + ("=" * 40) -ForegroundColor Cyan
        Write-Host "Sync complete: $addedCount added, $skippedCount skipped" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to sync contacts: $_"
    }
}

# Convenient aliases
New-Alias -Name flag-email -Value Add-OutlookFlag
New-Alias -Name tag-email -Value Add-OutlookCategory
New-Alias -Name refresh-categories -Value Reset-CategoryColorCache

# Export new functions and aliases
Export-ModuleMember -Function Add-OutlookFlag, Add-OutlookCategory, Reset-CategoryColorCache
Export-ModuleMember -Alias flag-email, tag-email, refresh-categories

# Create aliases for convenience
New-Alias -Name inbox -Value Get-OutlookInbox
New-Alias -Name send-email -Value Send-OutlookEmail
New-Alias -Name read-email -Value Read-OutlookMessage
New-Alias -Name reply-email -Value Send-OutlookReply
New-Alias -Name forward-email -Value Send-ForwardOutlookEmail
New-Alias -Name iinbox -Value Start-OutlookInteractive
New-Alias -Name email-categories -Value Get-OutlookCategories
New-Alias -Name archive-email -Value Move-OutlookMessageToArchive

# New aliases for archive search
New-Alias -Name search-archive -Value Search-OutlookArchive

# New aliases for calendar functions
New-Alias -Name calendar-today -Value Get-OutlookCalendarDay
New-Alias -Name calendar-day -Value Get-OutlookCalendarDay
New-Alias -Name calendar-week -Value Get-OutlookCalendarWeek
New-Alias -Name new-event -Value New-OutlookCalendarEvent

# New aliases for contact functions
New-Alias -Name add-contact -Value Add-OutlookContact
New-Alias -Name contacts -Value Get-OutlookContacts
New-Alias -Name sync-contacts -Value Sync-OutlookContactsFromReplies

# Export functions and aliases
Export-ModuleMember -Function Get-OutlookInbox, Read-OutlookMessage, Send-OutlookReply, Send-ForwardOutlookEmail, Start-OutlookInteractive, Get-OutlookCategories, Send-OutlookEmail, Move-OutlookMessageToArchive, Get-CachedMessages, Write-ColoredCategories, Get-CachedCategoryColorMap
Export-ModuleMember -Alias inbox, read-email, reply-email, forward-email, iinbox, email-categories, archive-email, send-email, flag-email, tag-email, refresh-categories

# Export archive search functions
Export-ModuleMember -Function Search-OutlookArchive, Get-CachedArchiveMessages
Export-ModuleMember -Alias search-archive

# Export calendar functions
Export-ModuleMember -Function Get-OutlookCalendarDay, Get-OutlookCalendarWeek, New-OutlookCalendarEvent
Export-ModuleMember -Alias calendar-today, calendar-day, calendar-week, new-event

# Export contact functions
Export-ModuleMember -Function Add-OutlookContact, Get-OutlookContacts, Sync-OutlookContactsFromReplies
Export-ModuleMember -Alias add-contact, contacts, sync-contacts

# Export configuration management functions
Export-ModuleMember -Function Get-EmailCliConfig, Get-ConfigValue, Set-ConfigValue, Reset-EmailCliConfig

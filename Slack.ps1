Add-Type -AssemblyName "System.Net.Http"

# To Do 
# reminders
# comments (inc delete)
# set User Profile (fix)
# Search pagination
# Access Logs
# OAuth authentication
# upload file text not binary

# Enumerations
# Channels
enum ChannelType {
    public
    private
    im
    mpim
}

enum ChannelShared {
    not_shared
    channel_shared
    org_shared
    global_shared
    external_shared
}

# Search
enum SearchSort {
    score
    timestamp
}

enum SearchSortDirection {
    ascending
    descending
}

# Files
enum SearchFileType {
    all      # All files
    spaces   # Posts
    snippets # Snippets
    images   # Image files
    gdocs    # Google docs
    zips     # Zip files
    pdfs     # PDF files
}

enum FileType {
    auto	#	Auto Detect Type
    text	#	Plain Text
    ai	#	Illustrator File
    apk	#	APK
    applescript	#	AppleScript
    binary	#	Binary
    bmp	#	Bitmap
    boxnote	#	BoxNote
    c	#	C
    csharp	#	C#
    cpp	#	C++
    css	#	CSS
    csv	#	CSV
    clojure	#	Clojure
    coffeescript	#	CoffeeScript
    cfm	#	ColdFusion
    d	#	D
    dart	#	Dart
    diff	#	Diff
    doc	#	Word Document
    docx	#	Word document
    dockerfile	#	Docker
    dotx	#	Word template
    email	#	Email
    eps	#	EPS
    epub	#	EPUB
    erlang	#	Erlang
    fla	#	Flash FLA
    flv	#	Flash video
    fsharp	#	F#
    fortran	#	Fortran
    gdoc	#	GDocs Document
    gdraw	#	GDocs Drawing
    gif	#	GIF
    go	#	Go
    gpres	#	GDocs Presentation
    groovy	#	Groovy
    gsheet	#	GDocs Spreadsheet
    gzip	#	GZip
    html	#	HTML
    handlebars	#	Handlebars
    haskell	#	Haskell
    haxe	#	Haxe
    indd	#	InDesign Document
    java	#	Java
    javascript	#	JavaScript/JSON
    jpg	#	JPEG
    keynote	#	Keynote Document
    kotlin	#	Kotlin
    latex	#	LaTeX/sTeX
    lisp	#	Lisp
    lua	#	Lua
    m4a	#	MPEG 4 audio
    markdown	#	Markdown (raw)
    matlab	#	MATLAB
    mhtml	#	MHTML
    mkv	#	Matroska video
    mov	#	QuickTime video
    mp3	#	mp4
    mp4	#	MPEG 4 video
    mpg	#	MPEG video
    mumps	#	MUMPS
    numbers	#	Numbers Document
    nzb	#	NZB
    objc	#	Objective-C
    ocaml	#	OCaml
    odg	#	OpenDocument Drawing
    odi	#	OpenDocument Image
    odp	#	OpenDocument Presentation
    odd	#	OpenDocument Spreadsheet
    odt	#	OpenDocument Text
    ogg	#	Ogg Vorbis
    ogv	#	Ogg video
    pages	#	Pages Document
    pascal	#	Pascal
    pdf	#	PDF
    perl	#	Perl
    php	#	PHP
    pig	#	Pig
    png	#	PNG
    post	#	Slack Post
    powershell	#	PowerShell
    ppt	#	PowerPoint presentation
    pptx	#	PowerPoint presentation
    psd	#	Photoshop Document
    puppet	#	Puppet
    python	#	Python
    qtz	#	Quartz Composer Composition
    r	#	R
    rtf	#	Rich Text File
    ruby	#	Ruby
    rust	#	Rust
    sql	#	SQL
    sass	#	Sass
    scala	#	Scala
    scheme	#	Scheme
    sketch	#	Sketch File
    shell	#	Shell
    smalltalk	#	Smalltalk
    svg	#	SVG
    swf	#	Flash SWF
    swift	#	Swift
    tar	#	Tarball
    tiff	#	TIFF
    tsv	#	TSV
    vb	#	VB.NET
    vbscript	#	VBScript
    vcard	#	vCard
    velocity	#	Velocity
    verilog	#	Verilog
    wav	#	Waveform audio
    webm	#	WebM
    wmv	#	Windows Media Video
    xls	#	Excel spreadsheet
    xlsx	#	Excel spreadsheet
    xlsb	#	Excel Spreadsheet (Binary, Macro Enabled)
    xlsm	#	Excel Spreadsheet (Macro Enabled)
    xltx	#	Excel template
    xml	#	XML
    yaml	#	YAML
    zip	#	Zip
}

enum ContentType {
    application_json
    application_xwwwformurlencoded
    multipart_formdata
}

enum Presence {
    auto
    away
}

# API

# Class definition
Class Slack {
    [string] $Token
    [string] $BaseUri = "https://slack.com/api"
    [int] $Limit = 1000000
    hidden [int] $AddUsersBatch = 30
    hidden [int] $SearchPageSize = 100

    # Utility methods
    [bool] Test () {
        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "api.test")
        if ($result.ok) {
            return $true
        } else {
            return $false
        }
    }

    [string] ConvertToUnixTime ([datetime] $date) {
        return (get-date ($date) -uformat %s).tostring()
    }

    # Team
    [object] GetTeamInfo () {
        $return = $this.GetTeamInfo($null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetTeamInfo ([string] $TeamID) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($TeamID)) { $body.Add("team",$TeamID) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "team.info", $body)
        if ($result.ok) {
            $return = $result.team
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetTeamProfile () {
        $return = $null

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "team.profile.get", $null)
        if ($result.ok) {
            $return = $result.profile
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] AddTeamUser ([string] $EmailAddress) {
        # restricted to adding 30 at a time
        $return = $null
        $body = new-object hashtable
        $body.Add("email", $EmailAddress)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "users.admin.invite", $body)
        if ($result.ok) {
            $return = $result.channel
        }

        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    # Apps - This method does not work - unknown method returned
    [object] GetApps () {
        $return = $null

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "bots.list")
        if ($result.ok) {
            $return = $result.bots
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }
    

    # Channels - does not deal with private channels - use Conversations instead
    [object] GetChannel ([string] $Name) {
        $return = $this.GetChannel($name, @([ChannelType]::public,[ChannelType]::private), $null, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetChannel ([string] $Name, [ChannelType[]]$Types, [ChannelShared[]] $IncludeSharedTypes, [switch] $IncludeArchived = $false) {
        $channels = $this.GetChannels($Types, $IncludeSharedTypes, $IncludeArchived)

        $return = $channels | where {$_.Name -like $Name}
        if ($return.count -eq 1) { $return = $return[0] }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetChannelInfo ([string] $ChannelID) {
        $return = $null
        $body = new-object hashtable
        $body.Add("channel",$ChannelID)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "channels.info", $body)
        if ($result.ok) {
            $return = $result.channel
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetChannels () {
        $return = $this.GetChannels(@([ChannelType]::public,[ChannelType]::private), $null, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetChannels ([ChannelType[]] $ChannelType) {
        $return = $this.GetChannels($ChannelType, [channelshared]::not_shared, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetChannels ([ChannelType[]] $ChannelType, [ChannelShared[]] $SharedTypes) {
        $return = $this.GetChannels($ChannelType, $SharedTypes, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetChannels ([ChannelType[]] $ChannelType, [ChannelShared[]] $SharedTypes, [switch] $Archived = $false) {
        $body = new-object hashtable
        $body.Add("exclude_archived", (!$Archived))
        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "channels.list", $body, $cursor)
            $return += $result.channels | where { `
                ($SharedTypes -eq $null -or `
                    ($_.is_shared -eq $true -and  $SharedTypes.contains([ChannelShared]::channel_shared)) -or `
                    ($_.is_org_shared -eq $true -and  $SharedTypes.contains([ChannelShared]::org_shared)) -or `
                    ($_.is_global_shared -eq $true -and  $SharedTypes.contains([ChannelShared]::global_shared)) -or `
                    ($_.is_ext_shared -eq $true -and  $SharedTypes.contains([ChannelShared]::external_shared)) -or `
                    ($SharedTypes.Contains([ChannelShared]::not_shared) -and ($_.is_shared -eq $false -or $_.is_shared -eq $null) -and ($_.is_org_shared -eq $false -or $_.is_org_shared -eq $null) -and ($_.is_global_shared -eq $false -or $_.is_global_shared -eq $null) -and ($_.is_ext_shared -eq $false -or $_.is_ext_shared -eq $null))) -and `
                ($ChannelType -eq $null -or `
                    ($_.is_private -eq $true -and  $ChannelType.contains([ChannelType]::private)) -or `
                    (($_.is_private -eq $false -or $_.is_private -eq $null)  -and  $ChannelType.contains([ChannelType]::public)) -or `
                    ($_.is_im -eq $true -and  $ChannelType.contains([ChannelType]::im)) -or `
                    ($_.is_mpim -eq $true -and  $ChannelType.contains([ChannelType]::mpim))) }
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    # Groups
    [object] GetGroup ([string] $Name) {
        $return = $this.GetGroup($name, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetGroup ([string] $Name, [switch] $IncludeArchived = $false) {
        $Groups = $this.GetGroups($IncludeArchived)

        $return = $Groups | where {$_.Name -like $Name}
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetGroupInfo ([string] $ChannelID) {
        $return = $null
        $body = new-object hashtable
        $body.Add("channel",$ChannelID)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "groups.info", $body)
        if ($result.ok) {
            $return = $result.group
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetGroups () {
        $return = $this.GetGroups($null, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetGroups ([ChannelType[]] $ChannelType) {
        $return = $this.GetGroups($ChannelType, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetGroups ([ChannelType[]] $ChannelType, [switch] $Archived = $false) {
        $body = new-object hashtable
        $body.Add("exclude_archived", (!$Archived))
        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "groups.list", $body, $cursor)
            $return += $result.Groups | where {
                ($ChannelType -eq $null -or `
                    ($_.is_mpim -eq $false -and  $ChannelType.contains([ChannelType]::private)) -or `
                    ($_.is_mpim -eq $true -and  $ChannelType.contains([ChannelType]::mpim)))
            }
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    # User Groups
    [object] GetUserGroups () {
        $return = $this.GetUserGroups($false, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUserGroups ([bool] $includeDisabled = $false, [bool] $includeUsers = $false) {
        $body = new-object hashtable
        $body.Add("include_count", $true)
        $body.Add("include_disabled", $includeDisabled)
        $body.Add("include_users", $includeUsers)
        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "usergroups.list", $body, $cursor)
            $return += $result.UserGroups
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] DisableUserGroup ([string] $UserGroupID = $false) {
        $body = new-object hashtable
        $body.Add("usergroup", $UserGroupID)
        $body.Add("include_count", $true)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "usergroups.disable", $body)
        $return = $result.Usergroup
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] EnableUserGroup ([string] $UserGroupID = $false) {
        $body = new-object hashtable
        $body.Add("usergroup", $UserGroupID)
        $body.Add("include_count", $true)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "usergroups.enable", $body)
        $return = $result.Usergroup
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] CreateUserGroup ([string] $Name) {
        $return = $this.CreateUserGroup($Name, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] CreateUserGroup ([string] $Name, [string] $Description) {
        $return = $this.CreateUserGroup($Name, $Description, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] CreateUserGroup ([string] $Name, [string] $Description, [string] $Channels) {
        $return = $this.CreateUserGroup($Name, $Description, $Channels, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] CreateUserGroup ([string] $Name, [string] $Description, [string] $ChannelIDs, [string] $Handle) {
        $return = $null
        $body = new-object hashtable

        $body.Add("name",$name)
        if (![string]::IsNullOrEmpty($ChannelIDs)) { $body.Add("channels",$ChannelIDs) }
        if (![string]::IsNullOrEmpty($Description)) { $body.Add("description",$description) }
        if (![string]::IsNullOrEmpty($Handle)) { $body.Add("handle",$Handle) }
        $body.Add("include_count",$true)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "usergroups.create", $body)
        $return = $result.usergroup
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] UpdateUserGroup ([string]$UserGroupID, [string] $Name, [string] $Description, [string] $ChannelIDs, [string] $Handle) {
        $return = $null
        $body = new-object hashtable

        $body.Add("usergroup",$UserGroupID)
        if (![string]::IsNullOrEmpty($Name)) { $body.Add("name",$Name) }
        if (![string]::IsNullOrEmpty($ChannelIDs)) { $body.Add("channels",$ChannelIDs) }
        if (![string]::IsNullOrEmpty($Description)) { $body.Add("description",$description) }
        if (![string]::IsNullOrEmpty($Handle)) { $body.Add("handle",$Handle) }
        $body.Add("include_count",$true)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "usergroups.create", $body)
        $return = $result.usergroup
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUserGroupUsers ([string] $UserGroupID) {
        $return = $this.GetUserGroupUsers($UserGroupID, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUserGroupUsers ([string] $UserGroupID, [bool] $includeDisabled = $false) {
        $body = new-object hashtable
        $body.Add("usergroup", $UserGroupID)
        $body.Add("include_disabled", $includeDisabled)
        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "usergroups.users.list", $body, $cursor)
            $return += $result.Users
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] UpdateUserGroupUsers ([string]$UserGroupID, [string] $UserIDs) {
        $return = $null
        $body = new-object hashtable

        $body.Add("usergroup",$UserGroupID)
        $body.Add("users",$UserIDs)
        $body.Add("include_count",$true)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "usergroups.users.update", $body)
        $return = $result.usergroup
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }


    # IM
    [object] GetIM ([string] $userId) {
        $IMs = $this.GetIMs()

        $return = $IMs | where {$_.user -eq $userId}
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetIMs () {
        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "im.list", $cursor)
            $return += $result.IMs
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    # Conversations
    [object] GetConversation ([string] $Name) {
        $return = $this.GetConversation($name, @([ChannelType]::public,[ChannelType]::private), $null, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetConversation ([string] $Name, [ChannelType[]]$Types, [ChannelShared[]] $IncludeSharedTypes, [switch] $IncludeArchived = $false) {
        $channels = $this.GetConversations($Types, $IncludeSharedTypes, $IncludeArchived)

        $return = $channels | where {$_.Name -like $Name -or $_.id -eq $name}
        if ($return.count -eq 1) { $return = $return[0] }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetConversations () {
        $return = $this.GetConversations(@([ChannelType]::public,[ChannelType]::private), [channelshared]::channel_shared, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetConversations ([ChannelType[]] $ChannelType) {
        $return = $this.GetConversations($ChannelType, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetConversations ([ChannelType[]] $ChannelType, [ChannelShared[]] $SharedTypes) {
        $return = $this.GetConversations($ChannelType, $SharedTypes, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetConversations ([ChannelType[]] $ChannelType, [ChannelShared[]] $SharedTypes, [switch] $Archived = $false) {
        $return = $null
        $body = new-object hashtable

        if ($ChannelType -eq $null) {
            # All all types
            $types = "private_channel,public_channel,im,mpim"
        } else {
            $types = ""
            $ChannelType | foreach {
                switch ($_) {
                    ([ChannelType]::private) { if (![string]::IsNullOrEmpty($types)) { $types += "," }; $types += "private_channel" }
                    ([ChannelType]::public) { if (![string]::IsNullOrEmpty($types)) { $types += "," }; $types += "public_channel" }
                    ([ChannelType]::im) { if (![string]::IsNullOrEmpty($types)) { $types += "," }; $types += "im" }
                    ([ChannelType]::mpim) { if (![string]::IsNullOrEmpty($types)) { $types += "," }; $types += "mpim" }
                }
            }
        }
        if (![string]::IsNullOrEmpty($types)) {
            $body.Add("types",$types)
        }

        if (!$Archived) { $body.Add("exclude_archived", $true) }

        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "conversations.list", $body, $cursor)
            $return += $result.channels | where {
                ($SharedTypes -eq $null -or `
                    ($_.is_shared -eq $true -and  $SharedTypes.contains([ChannelShared]::channel_shared)) -or `
                    ($_.is_org_shared -eq $true -and  $SharedTypes.contains([ChannelShared]::org_shared)) -or `
                    ($_.is_global_shared -eq $true -and  $SharedTypes.contains([ChannelShared]::global_shared)) -or `
                    ($_.is_ext_shared -eq $true -and  $SharedTypes.contains([ChannelShared]::external_shared)) -or `
                    ($SharedTypes.Contains([ChannelShared]::not_shared) -and ($_.is_shared -eq $false -or $_.is_shared -eq $null) -and ($_.is_org_shared -eq $false -or $_.is_org_shared -eq $null) -and ($_.is_global_shared -eq $false -or $_.is_global_shared -eq $null) -and ($_.is_ext_shared -eq $false -or $_.is_ext_shared -eq $null))) -and `
                ($ChannelType -eq $null -or `
                    ($_.is_private -eq $true -and  $ChannelType.contains([ChannelType]::private)) -or `
                    (($_.is_private -eq $false -or $_.is_private -eq $null)  -and  $ChannelType.contains([ChannelType]::public)) -or `
                    ($_.is_im -eq $true -and  $ChannelType.contains([ChannelType]::im)) -or `
                    ($_.is_mpim -eq $true -and  $ChannelType.contains([ChannelType]::mpim)))
            }
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetConversationInfo ([string] $ChannelID) {
        $return = $this.GetConversationInfo($ChannelID, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetConversationInfo ([string] $ChannelID, [switch] $IncludeMemberCount = $false) {
        $return = $null
        $body = new-object hashtable
        $body.Add("channel",$ChannelID)
        $body.Add("include_num_members", $IncludeMemberCount)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "conversations.info", $body)
        if ($result.ok) {
            $return = $result.channel
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] AddUsers ([string] $ChannelID, [string] $Users) {
        $return = $this.AddUsers($ChannelID, ($Users -split ','))
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] AddUsers ([string] $ChannelID, [array] $Users) {
        # restricted to adding 30 at a time
        $return = $null
        $body = new-object hashtable
        $body.Add("channel",$ChannelID)
        $body.Add("users", "")

        for ($i = 1; $i -le $users.count; $I+=$this.AddUsersBatch) {
            $body["users"] = $users[($i-1)..($i+($this.AddUsersBatch-2))] -join ','
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "conversations.invite", $body)
            if ($result.ok) {
                $return = $result.channel
            }
        }

        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [bool] RemoveUsers ([string] $ChannelID, [string] $UserIDs) {
        $return = $this.AddUsers($ChannelID, ($UserIDs -split ','))
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [bool] RemoveUsers ([string] $ChannelID, [array] $UserIDs) {
        $return = $null
        $result = $true

        foreach ($userID in $userIDs) {
            $result = $result -and $this.RemoveUser($ChannelID, $userID)
        }
        $return = $result

        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [bool] RemoveUser ([string] $ChannelID, [string] $UserID) {
        $return = $null
        $body = new-object hashtable
        $body.Add("channel",$ChannelID)
        $body.Add("user", $userID)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "conversations.kick", $body)
        $return = $result.ok

        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }


    # Users
    [object] GetUser ([string] $Name) {
        $users = $this.GetUsers()

        $return = $users | where {$_.email -like $Name -or $_.realname -like $name -or $_.id -eq $Name -or $_.name -like $Name -or $_.displayname -like $Name}
        if ($return.count -eq 1) { $return = $return[0] }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUserInfo ([string] $UserID) {
        $return = $null
        $body = new-object hashtable
        $body.Add("user",$UserID)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "users.info", $body)
        if ($result.ok) {
            $return = $result.user
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUserProfile () {
        $return = $this.GetUserProfile($null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUserProfile ([string] $UserID) {
        $return = $null
        $body = new-object hashtable
        $body.Add("include_labels",$true)
        if (![string]::IsNullOrEmpty($UserID)) {$body.Add("user",$UserID)}

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "users.profile.get", $body)
        if ($result.ok) {
            $return = $result.profile
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SetUserProfile ([hashtable] $Fields) {
        $return = $this.SetUserProfile($null, $fields)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SetUserProfile ([string] $UserID, [hashtable] $Fields) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($UserID)) { $body.Add("user",$UserID) }
        $body.Add("profile", (ConvertTo-Json $Fields))

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "users.profile.set", $body)
        if ($result.ok) {
            $return = $result.profile
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    # Presence
    [object] GetUserPresence () {
        $return = $this.GetUserPresence($null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUserPresence ([string] $UserID) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($UserID)) { $body.Add("user",$UserID) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "users.getPresence", $body)
        if ($result.ok) {
            $return = $result
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SetUserPresence ([Presence] $presence) {
        $return = $false
        $body = new-object hashtable
        $body.Add("presence",$presence)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "users.setPresence", $body)
        if ($result.ok) {
            $return = $true
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUserPreferences () {
        $return = $null

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "users.prefs.get")
        if ($result.ok) {
            $return = $result.prefs
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SetUserPreferences ([hashtable] $Fields) {
        $return = $null
        $body = new-object hashtable
        $body.Add("prefs",$Fields)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "users.prefs.set", $body)
        if ($result.ok) {
            $return = $result.prefs
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUsers () {
        $return = $this.GetUsers($false, $false)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetUsers ([bool] $IncludeDeleted=$false, [bool] $IncludeRestricted=$false) {
        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "users.list", $cursor)
            $return += $result.members | where { $_.is_bot -eq $false -and ($_.deleted -eq $false -or $IncludeDeleted) -and ($_.is_restricted -eq $false -or $IncludeRestricted) -and ($_.is_ultra_restricted -eq $false -or $IncludeRestricted) }
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)

        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetChannelMembers ([string] $channelID) {
        $return = $null
        $body = new-object hashtable

        $body.Add("channel",$channelID)

        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "conversations.members", $body, $cursor)
            $return += $result.members
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SetUserPhoto ([string]$FilePath, [filetype] $fileContentType) {
        $return = $this.SetUserPhoto((Get-Content -Path $FilePath -Raw -Encoding Byte), $fileContentType, 0, 0, 0)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }
    
    [object] SetUserPhoto ([Byte[]]$ImageContent, [filetype] $fileContentType) {
        $return = $this.SetUserPhoto($ImageContent, $fileContentType, 0, 0, 0)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SetUserPhoto ([Byte[]]$ImageContent, [filetype] $fileContentType, [int]$CropWidth, [int]$CropX, [int]$CropY) {
        $return = $null
        $body = new-object hashtable
        if ($CropWidth -gt 0) { $body.Add("crop_w", $CropWidth) }
        if ($CropX -gt 0) { $body.Add("crop_x", $CropX) }
        if ($CropY -gt 0) { $body.Add("crop_y", $CropY) }
        $body.add("image", $ImageContent)
        $body.add("filename", "profile.$fileContentType")

        $result = $this.CallBinaryAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "users.setPhoto", $body)
        if ($result.ok) {
            $return = $result.profile
        }
        if ($return) { return $return } else { return $false } # Due to method chaining / null collection bug
    }

    [bool] DeleteUserPhoto () {
        $return = $null

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "users.deletePhoto")
        $return = $result.ok
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }


    # Search
    [object] Search ([string] $query) {
        $return = $this.Search($query, $this.SearchPageSize, 1, $false, [searchsort]::score, [searchsortdirection]::descending)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] Search ([string] $query, $Page) {
        $return = $this.Search($query, $this.SearchPageSize, $Page, $false, [searchsort]::score, [searchsortdirection]::descending)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] Search ([string] $query, [int] $Count, [int] $Page, [bool] $Highlight=$false, [searchsort] $Sort = [searchsort]::score, [searchsortdirection] $SortDirection = [searchsortdirection]::descending) {
        $return = $null
        $body = new-object hashtable
        $body.Add("query",$query)
        if (![string]::IsNullOrEmpty($Count)) { $body.Add("count",$Count) }
        if (![string]::IsNullOrEmpty($page)) { $body.Add("page",$page) }
        if (![string]::IsNullOrEmpty($Highlight)) { $body.Add("highlight",$Highlight) }
        switch ($sort) {
            ([SearchSort]::score) { $body.Add("sort","score") }
            ([SearchSort]::timestamp) { $body.Add("sort","timestamp") }
        }
        switch ($SortDirection) {
            ([SearchSortDirection]::ascending) { $body.Add("sort_dir","asc") }
            ([SearchSortDirection]::descending) { $body.Add("sort_dir","desc") }
        }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "search.all", $body)
        if ($result.ok) {
            $return = @{ "files"=$result.files;"messages"=$result.messages;"posts"=$result.posts }
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SearchMessages ([string] $query) {
        $return = $this.SearchMessages($query, $this.SearchPageSize, 1, $false, [searchsort]::score, [searchsortdirection]::descending)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SearchMessages ([string] $query, $Page) {
        $return = $this.SearchMessages($query, $this.SearchPageSize, $Page, $false, [searchsort]::score, [searchsortdirection]::descending)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SearchMessages ([string] $query, [int] $Count, [int] $Page, [bool] $Highlight=$false, [searchsort] $Sort = [searchsort]::score, [searchsortdirection] $SortDirection = [searchsortdirection]::descending) {
        $return = $null
        $body = new-object hashtable
        $body.Add("query",$query)
        if (![string]::IsNullOrEmpty($Count)) { $body.Add("count",$Count) }
        if (![string]::IsNullOrEmpty($page)) { $body.Add("page",$page) }
        if (![string]::IsNullOrEmpty($Highlight)) { $body.Add("highlight",$Highlight) }
        switch ($sort) {
            ([SearchSort]::score) { $body.Add("sort","score") }
            ([SearchSort]::timestamp) { $body.Add("sort","timestamp") }
        }
        switch ($SortDirection) {
            ([SearchSortDirection]::ascending) { $body.Add("sort_dir","asc") }
            ([SearchSortDirection]::descending) { $body.Add("sort_dir","desc") }
        }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "search.messages", $body)
        if ($result.ok) {
            $return = @{ "messages"=$result.messages;"users"=$result.users;"teams"=$result.teams;"bots"=$result.bots }
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SearchFiles ([string] $query) {
        $return = $this.SearchFiles($query, $this.SearchPageSize, 1, $false, [searchsort]::score, [searchsortdirection]::descending)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SearchFiles ([string] $query, $Page) {
        $return = $this.SearchFiles($query, $this.SearchPageSize, $Page, $false, [searchsort]::score, [searchsortdirection]::descending)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] SearchFiles ([string] $query, [int] $Count, [int] $Page, [bool] $Highlight=$false, [searchsort] $Sort = [searchsort]::score, [searchsortdirection] $SortDirection = [searchsortdirection]::descending) {
        $return = $null
        $body = new-object hashtable
        $body.Add("query",$query)
        if (![string]::IsNullOrEmpty($Count)) { $body.Add("count",$Count) }
        if (![string]::IsNullOrEmpty($page)) { $body.Add("page",$page) }
        if (![string]::IsNullOrEmpty($Highlight)) { $body.Add("highlight",$Highlight) }
        switch ($sort) {
            ([SearchSort]::score) { $body.Add("sort","score") }
            ([SearchSort]::timestamp) { $body.Add("sort","timestamp") }
        }
        switch ($SortDirection) {
            ([SearchSortDirection]::ascending) { $body.Add("sort_dir","asc") }
            ([SearchSortDirection]::descending) { $body.Add("sort_dir","desc") }
        }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "search.files", $body)
        if ($result.ok) {
            $return = $result.Files
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    # Messages
    [object] GetMessage ([string] $ChannelId, [string] $Timestamp) {
        $return = $this.GetMessages($ChannelId, $Timestamp, $timestamp, $true, 1)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetMessages ([string] $ChannelId) {
        $return = $this.GetMessages($ChannelId, $null, $null, $true, 0)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetMessages ([string] $ChannelId, [string] $FromTimestamp) {
        $return = $this.GetMessages($ChannelId, $fromTimestamp, $null, $true, 0)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetMessages ([string] $ChannelId, [string] $FromTimestamp, [string] $toTimestamp) {
        $return = $this.GetMessages($ChannelId, $FromTimestamp, $toTimestamp, $true, 0)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetMessages ([string] $ChannelId, [string] $FromTimestamp, [string] $toTimestamp, [bool] $DatesInclusive = $true, [int] $limit) {
        $return = $null
        $body = new-object hashtable

        $body.Add("channel",$channelId)
        if (![string]::IsNullOrEmpty($FromTimestamp)) { $body.Add("oldest", $FromTimestamp) }
        if (![string]::IsNullOrEmpty($toTimestamp)) { $body.Add("latest", $toTimestamp) }
        $body.Add("inclusive",$DatesInclusive)
        if ($limit -and $limit -gt 0) { $body.Add("limit", $limit) }

        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "conversations.history", $body, $cursor)
            $return += $result.messages
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] Getreplies ([string] $ChannelId, [string] $timestamp) {
        $return = $this.Getreplies($ChannelId, $timestamp, $null, $null, $true, 0)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] Getreplies ([string] $channelId, [string] $timestamp, [string] $fromTimestamp, [string] $toTimestamp, [bool] $DatesInclusive = $true, [int] $limit) {
        $return = $null
        $body = new-object hashtable

        $body.Add("channel",$channelId)
        $body.Add("ts",$timestamp)
        if (![string]::IsNullOrEmpty($FromTimestamp)) { $body.Add("oldest", $FromTimestamp) }
        if (![string]::IsNullOrEmpty($toTimestamp)) { $body.Add("latest", $toTimestamp) }
        $body.Add("inclusive",$DatesInclusive)
        if ($limit -and $limit -gt 0) { $body.Add("limit", $limit) } else { $body.add("limit", 1000) }

        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "conversations.replies", $body, $cursor)
            $return += $result.messages
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetEmojis () {
        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "emoji.list", $cursor)
            $return += $result.emoji
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] PostMessage ([string] $ChannelID, [string] $Text) {
        $return = $this.PostMessage($channelID, $text, $null, $null, $null, $true, $null, $null, $null, $true)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] PostMessage ([string] $ChannelID, [string] $Text, [string] $Attachments, [string] $Blocks, [string] $ThreadID, [bool] $AsUser = $false, [string] $AppName, [string] $IconEmoji, [string] $IconUrl, [bool] $linkNames = $true) {
        $return = $null
        $body = new-object hashtable

        $body.Add("channel",$channelID)
        if (![string]::IsNullOrEmpty($text)) { $body.Add("text",$text) }
        if ($attachments) { $body.Add("attachments",$Attachments) }
        if ($Blocks) { $body.Add("blocks",$Blocks) }
        if (![string]::IsNullOrEmpty($ThreadID)) { $body.Add("thread_ts",$ThreadID) }
        $body.Add("as_user",$AsUser)
        if (![string]::IsNullOrEmpty($AppName)) { $body.Add("username",$AppName) }
        if (![string]::IsNullOrEmpty($IconEmoji)) { $body.Add("icon_emoji",$IconEmoji) }
        if (![string]::IsNullOrEmpty($IconUrl)) { $body.Add("icon_url",$IconUrl) }
        $body.Add("link_names",$linkNames)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "chat.postMessage", $body)
        $return = $result.message
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] PostMessageEphemeral ([string] $ChannelID, [string] $UserID, [string] $Text) {
        $return = $this.PostMessageEphemeral($channelID, $UserID, $text, $null, $null, $null, $true, $true)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] PostMessageEphemeral ([string] $ChannelID, [string] $UserID, [string] $Text, [string] $Attachments, [string] $Blocks, [string] $ThreadID, [bool] $AsUser = $false, [bool] $linkNames = $true) {
        $return = $null
        $body = new-object hashtable

        $body.Add("channel",$channelID)
        $body.Add("user",$UserID)
        if (![string]::IsNullOrEmpty($text)) { $body.Add("text",$text) }
        if ($attachments) { $body.Add("attachments",$Attachments) }
        if ($Blocks) { $body.Add("blocks",$Blocks) }
        if (![string]::IsNullOrEmpty($ThreadID)) { $body.Add("thread_ts",$ThreadID) }
        $body.Add("as_user",$AsUser)
        $body.Add("link_names",$linkNames)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "chat.postEphemeral", $body)
        $return = $result.message
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] DeleteMessage ([string] $ChannelID, [string] $MessageID) {
        $return = $this.DeleteMessage($channelID, $MessageID, $true)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [bool] DeleteMessage ([string] $ChannelID, [string] $MessageID, [bool] $AsUser = $false) {
        $return = $null
        $body = new-object hashtable

        $body.Add("channel",$channelID)
        $body.Add("ts",$MessageID)
        $body.Add("as_user",$AsUser)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "chat.delete", $body)
        $return = $result.ok
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetPins ([string] $ChannelID) {
        $return = $null
        $body = new-object hashtable
        $body.Add("channel",$ChannelID)

        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "pins.list", $body, $cursor)
            $return += $result.items
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] AddPin ([string] $ChannelID, [string] $FileID, [string] $FileCommentID, [string] $MessageTimestamp) {
        $return = $null
        $body = new-object hashtable
        $body.Add("channel",$ChannelID)
        if (![string]::IsNullOrEmpty($FileID)) { $body.Add("file",$FileID) }
        if (![string]::IsNullOrEmpty($FileCommentID)) { $body.Add("file_comment",$FileCommentID) }
        if (![string]::IsNullOrEmpty($MessageTimestamp)) { $body.Add("timestamp",$MessageTimestamp) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "pins.add", $body)
        if ($result.ok) {
            $return = $true
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] RemovePin ([string] $ChannelID, [string] $FileID, [string] $FileCommentID, [string] $MessageTimestamp) {
        $return = $null
        $body = new-object hashtable
        $body.Add("channel",$ChannelID)
        if (![string]::IsNullOrEmpty($FileID)) { $body.Add("file",$FileID) }
        if (![string]::IsNullOrEmpty($FileCommentID)) { $body.Add("file_comment",$FileCommentID) }
        if (![string]::IsNullOrEmpty($MessageTimestamp)) { $body.Add("timestamp",$MessageTimestamp) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "pins.remove", $body)
        if ($result.ok) {
            $return = $true
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetStars () {
        $return = $this.GetStars($this.SearchPageSize, 1, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetStars ([int] $Count, [int] $Page, [string] $Cursor) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($Count)) { $body.Add("count",$Count) }
        if (![string]::IsNullOrEmpty($page)) { $body.Add("page",$page) }
        if (![string]::IsNullOrEmpty($Cursor)) { $body.Add("cursor",$Cursor) }

        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "stars.list", $body, $cursor)
            $return += $result.items
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] AddStar ([string] $ChannelID, [string] $FileID, [string] $FileCommentID, [string] $MessageTimestamp) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($ChannelID)) { $body.Add("channel",$ChannelID) }
        if (![string]::IsNullOrEmpty($FileID)) { $body.Add("file",$FileID) }
        if (![string]::IsNullOrEmpty($FileCommentID)) { $body.Add("file_comment",$FileCommentID) }
        if (![string]::IsNullOrEmpty($MessageTimestamp)) { $body.Add("timestamp",$MessageTimestamp) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "stars.add", $body)
        if ($result.ok) {
            $return = $true
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] RemoveStar ([string] $ChannelID, [string] $FileID, [string] $FileCommentID, [string] $MessageTimestamp) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($ChannelID)) { $body.Add("channel",$ChannelID) }
        if (![string]::IsNullOrEmpty($FileID)) { $body.Add("file",$FileID) }
        if (![string]::IsNullOrEmpty($FileCommentID)) { $body.Add("file_comment",$FileCommentID) }
        if (![string]::IsNullOrEmpty($MessageTimestamp)) { $body.Add("timestamp",$MessageTimestamp) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "stars.remove", $body)
        if ($result.ok) {
            $return = $true
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetMessageReactions () {
        $return = $this.GetMessageReactions($null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetMessageReactions ([string] $ChannelID) {
        $return = $this.GetMessageReactions($ChannelID, $null, $null, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetMessageReactions ([string] $ChannelID, [string] $FileID, [string] $FileCommentID, [string] $MessageTimestamp) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($ChannelID)) { $body.Add("channel",$ChannelID) }
        if (![string]::IsNullOrEmpty($FileID)) { $body.Add("file",$FileID) }
        if (![string]::IsNullOrEmpty($FileCommentID)) { $body.Add("file_comment",$FileCommentID) }
        if (![string]::IsNullOrEmpty($MessageTimestamp)) { $body.Add("timestamp",$MessageTimestamp) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "reactions.get", $body)
        if ($result.ok) {
            $return = $result.file
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetReactions () {
        $return = $this.GetReactions($this.SearchPageSize, 1, $null, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetReactions ([int] $Count, [int] $Page, [string] $Cursor, [string] $UserID) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($Count)) { $body.Add("count",$Count) }
        if (![string]::IsNullOrEmpty($page)) { $body.Add("page",$page) }
        if (![string]::IsNullOrEmpty($Cursor)) { $body.Add("cursor",$Cursor) }

        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "reactions.list", $body, $cursor)
            $return += $result.items
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] AddReaction ([string] $ChannelID, [string] $IconEmoji, [string] $FileID, [string] $FileCommentID, [string] $MessageTimestamp) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($ChannelID)) { $body.Add("channel",$ChannelID) }
        $body.Add("name",$IconEmoji)
        if (![string]::IsNullOrEmpty($FileID)) { $body.Add("file",$FileID) }
        if (![string]::IsNullOrEmpty($FileCommentID)) { $body.Add("file_comment",$FileCommentID) }
        if (![string]::IsNullOrEmpty($MessageTimestamp)) { $body.Add("timestamp",$MessageTimestamp) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "reactions.add", $body)
        if ($result.ok) {
            $return = $true
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] RemoveReaction ([string] $ChannelID, [string] $IconEmoji, [string] $FileID, [string] $FileCommentID, [string] $MessageTimestamp) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($ChannelID)) { $body.Add("channel",$ChannelID) }
        $body.Add("name",$IconEmoji)
        if (![string]::IsNullOrEmpty($FileID)) { $body.Add("file",$FileID) }
        if (![string]::IsNullOrEmpty($FileCommentID)) { $body.Add("file_comment",$FileCommentID) }
        if (![string]::IsNullOrEmpty($MessageTimestamp)) { $body.Add("timestamp",$MessageTimestamp) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "reactions.remove", $body)
        if ($result.ok) {
            $return = $true
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }


    # Commands
    [object] PostCommand ([string] $ChannelID, [string] $Command) {
        $return = $this.PostCommand($channelID, $Command, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] PostCommand ([string] $ChannelID, [string] $Command, [string] $Parameters) {
        $return = $null
        $body = new-object hashtable

        $body.Add("channel",$channelID)
        $body.Add("command",$command)
        if (![string]::IsNullOrEmpty($parameters)) { $body.Add("text",$parameters) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "chat.command", $body)
        $return = $result.response
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetCommands () {
        $return = $null

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "commands.list")
        $return = $result.commands
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }


    # Files
    [object] GetFiles ([SearchFileType[]] $FileType) {
        $return = $this.GetFiles($null, $FileType)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetFiles ([string] $ChannelID) {
        $return = $this.GetFiles($ChannelID, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetFiles ([string] $ChannelID, [SearchFileType[]] $FileType) {
        $return = $this.GetFiles($ChannelID, $FileType, $null)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetFiles ([string] $ChannelID, [SearchFileType[]] $FileType, [string] $UserID) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($ChannelID)) { $body.Add("channel",$ChannelID) }
        if (![string]::IsNullOrEmpty($UserID)) { $body.Add("user",$UserID) }
        if ($FileType) { $body.Add("types",$FileType -join ',') }

        $return = @()
        $cursor = $null

        do {
            $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "files.list", $body, $cursor)
            $return += $result.files
            if ($result.response_metadata) { $cursor = $result.response_metadata.next_cursor } else { $cursor = $null }
        } while ($cursor)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] GetFileInfo ([string] $FileID) {
        $return = $null
        $body = new-object hashtable
        if (![string]::IsNullOrEmpty($FileID)) { $body.Add("file",$FileID) }

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Get, "files.info", $body)
        if ($result.ok) {
            $return = $result.file
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] UploadFile ([string[]]$ChannelIDs, [string]$FilePath, [FileType]$FileType, [string]$FileName, [string]$Title, [string]$Comment, [string] $ThreadID) {
        $return = $this.UploadFile($ChannelIDs, (Get-Content -Path $FilePath -Raw -Encoding Byte), $filetype, $FileName, $Title, $Comment, $ThreadID)
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] UploadFile ([string[]]$ChannelIDs, [Byte[]]$Content, [FileType]$FileType, [string]$FileName, [string]$Title, [string]$Comment, [string] $ThreadID) {

        $return = $null
        $body = new-object hashtable
        if ($ChannelIDs) { $body.Add("channels", $ChannelIDs -join ',') }
        if (![string]::IsNullOrEmpty($FileType)) { $body.Add("filetype", $FileType) }
        if (![string]::IsNullOrEmpty($FileName)) { $body.Add("filename", $FileName) }
        if (![string]::IsNullOrEmpty($Title)) { $body.Add("title", $Title) }
        if (![string]::IsNullOrEmpty($Comment)) { $body.Add("initial_comment", $Comment) }
        $body.Add("file", $Content)
        if (![string]::IsNullOrEmpty($ThreadID)) { $body.Add("thread_ts",$ThreadID) }

        $result = $this.CallBinaryAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "files.upload", $body)
        if ($result.ok) {
            $return = $result.file
        }
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [bool] DeleteFile ([string] $FileID) {
        $return = $null
        $body = new-object hashtable

        $body.Add("file",$FileID)

        $result = $this.CallAPI([Microsoft.PowerShell.Commands.WebRequestMethod]::Post, "files.delete", $body)
        $return = $result.ok
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }


    # internal Function to call the API
    [object] CallAPI ([Microsoft.PowerShell.Commands.WebRequestMethod] $method, [string] $function) {
        $return = $this.CallAPI($method, $function, (New-Object hashtable), $null) # Need to pass in a new object otherwise this crashes PowerShell
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] CallAPI ([Microsoft.PowerShell.Commands.WebRequestMethod] $method, [string] $function, [string] $cursor = $null) {
        $return = $this.CallAPI($method, $function, (New-Object hashtable), $cursor) # Need to pass in a new object otherwise this crashes PowerShell
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] CallAPI ([Microsoft.PowerShell.Commands.WebRequestMethod] $method, [string] $function, [hashtable] $body) {
        $return = $this.CallAPI($method, $function, $body, $null) # Need to pass in a new object otherwise this crashes PowerShell
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] CallAPI ([Microsoft.PowerShell.Commands.WebRequestMethod] $method, [string] $function, [hashtable] $body, [string] $cursor = $null) {
        $optionalParams = @{}

        # Add the token to the body
        if (!$body) { $body = New-Object hashtable }
        if (!$body.Contains("token")) {
            $body.Add("token", $this.Token)
        }
        if (!$body.Contains("limit")) {
            $body.Add("limit", $this.Limit)
        }
        if ($cursor) {
            if ($body.Contains("cursor")) {
                $body["cursor"] = $cursor
            } else {
                $body.Add("cursor", $cursor)
            }
        } else {
            $body.Remove("cursor")
        }
            
        # Add some functionality to handle the 429 too many requests error
        $retry = $false
        $response = $null # Set in the do..while loop
        $result = $null # Set in the do..while loop
        $statusCode = $null # Set in the Invoke-WebRequest

        # If it is runing Core, add the additional parameters
        if ($global:PSVersionTable.PSEdition -eq "Core") {
            $optionalParams.add("SkipHttpErrorCheck", $null)
        }

        do {
            # Does not respect -ErrorAction parameter, try Ignore, or $ErrorActionPreference. PS Core will ignore errors, Windows PS will fall into Catch
            try {
                $retry = $false
                $response = Invoke-WebRequest -Method $method -Uri ($this.BaseUri + "/" + $function) -Body $body @optionalParams -ErrorAction Ignore
                $result = Convertfrom-json $response.content
                if ($result.ok -eq $false) {
                    # Will break here for PowerShell Core for response code 429
                    switch ($response.statusCode) {
                        200 {
                            # Throw te Slack error
                            throw ("Slack error: $($result.error)")
                            break
                        }
                        429 {
                            # Check for "too many requests" and find the delay we need to wait for
                            $retry = $true
                            if ($response.headers.'retry-after') {
                                $delay = [int]::Parse($response.headers.'retry-after'[0])
                            } else {
                                $delay = 5
                            }
                            Write-Debug "Too Many Requests, sleeping $($delay)s"
                            sleep $delay
                            break
                        }
                        default {
                            # Throw an exception
                            throw ("Slack API call failed with status code $($response.statusCode): $($response.statusDescription)")
                        }
                    }
                }
            }
            catch {
                # Will break here for Windows PowerShell for response code 429
                # Full response headers (values) are not available only the names
                $response = $_.Exception.response
                switch ($response.statusCode) {
                    429 {
                        # Check for "too many requests" and find the delay we need to wait for
                        $retry = $true
                        if ($response.headers.AllKeys -contains "retry-after") {
                            $delay = [int]::Parse($response.headers.Get("retry-after"))
                        } else {
                            $delay = 5
                        }
                        Write-Debug "Too Many Requests, sleeping $($delay)s"
                        sleep $delay
                        break
                    }
                    default {
                        # Check for a terminating error
                        if ($response.statusCode) {
                            throw ("Slack API call failed with status code $($response.statusCode): $($response.statusDescription)")
                        } else {
                            throw
                        }
                    }
                }
            }
        } while ($retry)

        # Look at what came back
        if (!$result.ok) {
            throw ("'$($result.error)': $($result.response_metadata.messages -join ', ')")
        } else {
            return $result
        }
    }

    [object] CallBinaryAPI ([Microsoft.PowerShell.Commands.WebRequestMethod] $method, [string] $function, [hashtable] $Body, [string] $filename, [byte[]] $fileContents) {
        $return = $this.CallBinaryAPI($method, $function, $Body, $filename, $fileContents, [FileType]::auto, "") # Need to pass in a new object otherwise this crashes PowerShell
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [object] CallBinaryAPI ([Microsoft.PowerShell.Commands.WebRequestMethod] $method, [string] $function, [hashtable] $Body, [byte[]] $fileContents, [filetype] $fileContentType, [string] $fileArgumentName = "file") {
        $return = $this.CallBinaryAPI($method, $function, $Body, "", $fileContents, $fileContentType, $fileArgumentName) # Need to pass in a new object otherwise this crashes PowerShell
        if ($return) { return $return } else { return $null } # Due to method chaining / null collection bug
    }

    [string] GenerateMultiPartStringSection ([string] $boundary, [string] $LF = "`r`n", [string] $name, [string] $value) {
        $return = 
            ("--$boundary$LF" +
            "Content-Disposition: form-data; name=`"$name`"$LF" +
            "Content-Type: multipart/form-data$LF$LF" +
            "$($value -join ",")$LF")
        return $return
    }

    [string] GenerateMultiPartBinarySection ([string] $boundary, [string] $LF = "`r`n", [string] $name, [byte[]] $value, [string] $filename) {
        # Encode the binary data
        $enc = [System.Text.Encoding]::GetEncoding("iso-8859-1")
        $fileEnc = $enc.GetString($value)
        $return = 
            ("--$boundary$LF" +
            "Content-Disposition: form-data; name=`"$name`"; filename=`"$fileName`"$LF" +
            "Content-Type: 'multipart/form-data'$LF$LF" +
            "$fileEnc$LF")
        return $return
    }

    [object] CallBinaryAPI([Microsoft.PowerShell.Commands.WebRequestMethod] $method, [string] $function, [hashtable] $body) {

        #Note: For binary data uploads we need to specify the filename along with the image data
        # Make sure the body is passed in as this will contain the data to upload
        if ($body) {
            try {
                $LF = "`r`n"
                $boundary = [System.Guid]::NewGuid().ToString()

                if (!$body.Contains("token")) {
                    $body.Add("token",$this.Token)
                }
                if (!$body.Contains("limit")) {
                    $body.Add("limit",$this.Limit)
                }
                # Find the filename to pass in to the binary section generator
                $filename = $body["filename"]

                # generate the payload
                $bodyLines = ""
                switch ($body.keys) {
                    'image'     {$bodyLines += $this.GenerateMultiPartBinarySection($boundary, $lf, $_, $body[$_], $filename)}
                    'file'     {$bodyLines += $this.GenerateMultiPartBinarySection($boundary, $lf, $_, $body[$_], $filename)}
                    default     {$bodyLines += $THIS.GenerateMultiPartStringSection($boundary, $lf, $_, $body[$_])}
                }
                $bodyLines += "--$boundary--$LF"
                
                $result = Invoke-RestMethod -method $method -Uri ($this.BaseUri + "/" + $function) -Body $bodyLines -ContentType "multipart/form-data; boundary=`"$boundary`""

                # Look at what came back
                if (!$result.ok) {
                    throw ("'$($result.error)': $($result.response_metadata.messages -join ', ')")
                } else {
                    return $result
                }
            }
            catch {
                throw $_.Exception.Message
            }
        } else {
            throw "Parameter 'Body' is mandatory"
        }
    }
}

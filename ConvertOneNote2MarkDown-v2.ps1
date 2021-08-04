[CmdletBinding()]
param ()

Function Validate-Dependencies {
    [CmdletBinding()]
    param ()

    # Validate assemblies
    if ( ($env:OS -imatch 'Windows') -and ! (Get-Item -Path $env:windir\assembly\GAC_MSIL\*onenote*) ) {
        "There are missing onenote assemblies. Please ensure the Desktop version of Onenote 2016 or above is installed." | Write-Warning
    }

    # Validate dependencies
    if (! (Get-Command -Name 'pandoc.exe') ) {
        throw "Could not locate pandoc.exe. Please ensure pandoc is installed."
    }
}

Function Get-DefaultConfiguration {
    [CmdletBinding()]
    param ()

    # The default configuration
    $config = [ordered]@{
        notesdestpath = @{
            description = @'
Specify folder path that will contain your resulting Notes structure - Default: c:\temp\notes
'@
            default = 'c:\temp\notes'
            validateOptions = 'directoryexists'
        }
        targetNotebook = @{
            description = @'
Specify a notebook name to convert
'': Convert all notebooks - Default
'mynotebook': Convert specific notebook named 'mynotebook'
'@
            default = ''
        }
        usedocx = @{
            description = @'
Whether to create new word docs or reuse existing ones
1: Always create new .docx files - Default
2: Use existing .docx files (90% faster)
'@
            default = 1
        }
        keepdocx = @{
            description = @'
Whether to discard word docs after conversion
1: Discard intermediate .docx files - Default
2: Keep .docx files
'@
            default = 1
        }
        prefixFolders = @{
            description = @'
Whether to use prefix vs subfolders
1: Create folders for subpages (e.g. Page\Subpage.md) - Default
2: Add prefixes for subpages (e.g. Page_Subpage.md)
'@
            default = 1
        }
        medialocation = @{
            description = @'
Whether to store media in single or multiple folders
1: Images stored in single 'media' folder at Notebook-level - Default
2: Separate 'media' folder for each folder in the hierarchy
'@
            default = 1
        }
        conversion = @{
            description = @'
Specify conversion type
1: markdown (Pandoc) - Default
2: commonmark (CommonMark Markdown)
3: gfm (GitHub-Flavored Markdown)
4: markdown_mmd (MultiMarkdown)
5: markdown_phpextra (PHP Markdown Extra)
6: markdown_strict (original unextended Markdown)
'@
            default = 1
        }
        headerTimestampEnabled = @{
            description = @'
Whether to include page timestamp and separator at top of document
1: Include - Default
2: Don't include
'@
            default = 1
        }
        keepspaces = @{
            description = @'
Whether to clear double spaces between bullets
1: Clear double spaces in bullets - Default
2: Keep double spaces
'@
            default = 1
        }
        keepescape = @{
            description = @'
Whether to clear escape symbols from md files
1: Clear '\' symbol escape character from files - Default
2: Keep '\' symbol escape
'@
            default = 1
        }
        keepPathSpaces = @{
            description = @'
Whether to replace spaces with dashes i.e. '-' in file and folder names
1: Replace spaces with dashes in file and folder names - Default
2: Keep spaces in file and folder names (1 space between words, removes preceding and trailing spaces)"
'@
            default = 1
        }
    }

    $config
}

Function New-ConfigurationFile {
    [CmdletBinding()]
    param ()

    # Generate a configuration file config.example.ps1
    @'
#
# Note: This config file is for those who are lazy to type in configuration everytime you run ./ConvertOneNote2MarkDown-v2.ps1
#
# Steps:
#   1) Rename this file to config.ps1. Ensure it is in the same folder as the ConvertOneNote2MarkDown-v2.ps1 script
#   2) Configure the options below to your liking
#   3) Run the main script: ./ConvertOneNote2MarkDown-v2.ps1. Sit back while the script starts converting immediately.
'@ | Out-File "$PSScriptRoot/config.example.ps1" -Encoding utf8

    $defaultConfig = Get-DefaultConfiguration
    foreach ($key in $defaultConfig.Keys) {
        # Add a '#' in front of each line of the option description
        $defaultConfig[$key]['description'].Trim() -replace "^|`n", "`n# " | Out-File "$PSScriptRoot/config.example.ps1" -Encoding utf8 -Append

        # Write the variable
        if ( $defaultConfig[$key]['default'] -is [string]) {
            "`$$key = '$( $defaultConfig[$key]['default'] )'" | Out-File "$PSScriptRoot/config.example.ps1" -Encoding utf8 -Append
        }else {
            "`$$key = $( $defaultConfig[$key]['default'] )" | Out-File "$PSScriptRoot/config.example.ps1" -Encoding utf8 -Append
        }
    }
}

Function Compile-Configuration {
    [CmdletBinding()]
    param ()

    # Get a default configuration
    $config = Get-DefaultConfiguration

    # Override configuration
    if (Test-Path $PSScriptRoot/config.ps1) {
        # Get override configuration from config file ./config.ps1
        & {
            . $PSScriptRoot/config.ps1
            foreach ($key in @($config.Keys)) {
                $config[$key]['value'] = Get-Variable -Name $key -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Value
                # Trim string
                if ($config[$key]['value'] -is [string]) {
                    $config[$key]['value'] = $config[$key]['value'].Trim()
                }
                # Remove trailing slash(es) for paths
                if ($key -match 'path' -and $config[$key]['value'] -match '[/\\]') {
                    $config[$key]['value'] = $config[$key]['value'].TrimEnd('/').TrimEnd('\')
                }
            }
        }
    }else {
        # Get override configuration from interactive prompts
        foreach ($key in $config.Keys) {
            "" | Write-Host -ForegroundColor Cyan
            $config[$key]['description'] | Write-Host -ForegroundColor Cyan
            # E.g. 'string', 'int'
            $typeName = [Microsoft.PowerShell.ToStringCodeMethods]::Type($config[$key]['default'].GetType())
            # Keep prompting until we get a answer of castable type
            do {
                # Cast the input as a type. E.g. Read-Host -Prompt 'Entry' -as [int]
                $config[$key]['value'] = Invoke-Expression -Command "(Read-Host -Prompt 'Entry') -as [$typeName]"
            }while ($null -eq $config[$key]['value'])
            # Fallback on default value if the input is empty string
            if ($config[$key]['value'] -is [string] -and $config[$key]['value'] -eq '') {
                $config[$key]['value'] = $config[$key]['default']
            }
            # Fallback on default value if the input is empty integer (0)
            if ($config[$key]['value'] -is [int] -and $config[$key]['value'] -eq 0) {
                $config[$key]['value'] = $config[$key]['default']
            }
        }
    }

    $config
}

Function Validate-Configuration {
    [CmdletBinding(DefaultParameterSetName='default')]
    param (
        [Parameter(ParameterSetName='default',Position=0)]
        [object]
        $Config
    ,
        [Parameter(ParameterSetName='pipeline',ValueFromPipeline)]
        [object]
        $InputObject
    )
    process {
        if ($InputObject) {
            $Config = $InputObject
        }
        if ($null -eq $Config) {
            throw "No input parameters specified."
        }

        # Validate a given configuration against a prototype configuration
        $defaultConfig = Get-DefaultConfiguration
        foreach ($key in $defaultConfig.Keys) {
            if (! $Config.Contains($key)) {
                throw "Missing configuration option '$key'"
            }
            if ($defaultConfig[$key]['default'].GetType().FullName -ne $Config[$key]['value'].GetType().FullName) {
                throw "Invalid configuration option '$key'. Expected a value of type $( $defaultConfig[$key]['default'].GetType().FullName ), but value was of type $( $config[$key]['value'].GetType().FullName )"
            }
            if ($defaultConfig[$key].Contains('validateOptions')) {
                if ($defaultConfig[$key]['validateOptions'] -contains 'directoryexists') {
                    if ( ! $config[$key]['value'] -or ! (Test-Path $config[$key]['value'] -PathType Container -ErrorAction SilentlyContinue) ) {
                        throw "Invalid configuration option '$key'. The directory '$( $config[$key]['value'] )' does not exist, or is a file"
                    }
                }
            }
        }

        # Warn of unknown configuration options
        foreach ($key in $config.Keys) {
            if (! $defaultConfig.Contains($key)) {
                "Unknown configuration option '$key'" | Write-Warning
            }
        }

        $Config
    }
}

Function Remove-InvalidFileNameChars {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]$Name,
        [switch]$KeepPathSpaces
    )

    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    $newName = $newName -replace "\[", "("
    $newName = $newName -replace "\]", ")"
    $newName =  if ($KeepPathSpaces) {
                    $newName -replace "\s", " "
                } else {
                    $newName -replace "\s", "-"
                }
    $newName = $newName.Substring(0, $(@{$true = 130; $false = $newName.length }[$newName.length -gt 150]))
    return $newName.Trim() # Remove boundary whitespaces
}

Function Remove-InvalidFileNameCharsInsertedFiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]$Name,
        [string]$Replacement = "",
        [string]$SpecialChars = "#$%^*[]'<>!@{};",
        [switch]$KeepPathSpaces
    )

    $rePattern = ($SpecialChars.ToCharArray() | ForEach-Object { [regex]::Escape($_) }) -join "|"

    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    $newName = $newName -replace $rePattern, ""
    $newName =  if ($KeepPathSpaces) {
                    $newName -replace "\s", " "
                } else {
                    $newName -replace "\s", "-"
                }
    return $newName.Trim() # Remove boundary whitespaces
}

Function New-OneNoteConnection {
    [CmdletBinding()]
    param ()

    # Create a OneNote connection. See: See: https://docs.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote
    if ($PSVersionTable.PSVersion.Major -le 5) {
        $OneNote = New-Object -ComObject OneNote.Application
    }else {
        # Works between powershell 6.0 and 7.0, but not >= 7.1
        Add-Type -Path $env:windir\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll # -PassThru
        $OneNote = [Microsoft.Office.Interop.OneNote.ApplicationClass]::new()
    }

    $OneNote
}

Function Remove-OneNoteConnection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    )

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNoteConnection) | Out-Null
}

Function Get-OneNoteHierarchy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $OneNoteConnection
    )

    # Open OneNote hierarchy
    [xml]$hierarchy = ""
    $OneNoteConnection.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$hierarchy)

    $hierarchy
}

Function Get-OneNotePageContent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $OneNoteConnection
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [int]
        $PageId
    )

    # Get page's xml content
    [xml]$pagexml = ""
    $OneNoteConnection.GetPageContent($page.ID, [ref]$pagexml, 7)

    $pagexml
}

Function Publish-OneNotePageToDocx {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $OneNoteConnection
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [int]
        $PageId
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Destination
    )

    $OneNoteConnection.Publish($PageId, $Destination, "pfWord", "")
}

Function New-SectionGroupConversionConfig {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $NotesDestination
    ,
        [Parameter(Mandatory)]
        [object]
        $Config
    ,
        [Parameter(Mandatory)]
        [array]
        $SectionGroups
    ,
        [Parameter(Mandatory)]
        [int]
        $LevelsFromRoot
    )

    $sectionGroupConversionConfig = [System.Collections.ArrayList]@()

    # Build an object representing the conversion of a Section Group (treat a Notebook as a Section Group, it is no different)
    foreach ($sectionGroup in $SectionGroups) {
        $cfg = [ordered]@{}

        if ($LevelsFromRoot -eq 0) {
            "`n$( '#' * ($LevelsFromRoot + 1) ) Building conversion configuration for $( $sectionGroup.name ) [Notebook]" | Write-Host -ForegroundColor DarkGreen
        }else {
            "`n$( '#' * ($LevelsFromRoot + 1) ) Building conversion configuration for $( $sectionGroup.name ) [Section Group]" | Write-Host -ForegroundColor DarkGray
        }

        # Build this Section Group
        $cfg = [ordered]@{}
        $cfg['object'] = $sectionGroup # Keep a reference to the SectionGroup object
        $cfg['kind'] = 'SectionGroup'
        $cfg['levelsFromRoot'] = $LevelsFromRoot
        $cfg['uri'] = $sectionGroup.path # E.g. https://d.docs.live.net/01234567890abcde/Skydrive Notebooks/mynotebook/mysectiongroup
        $cfg['fileName'] = $sectionGroup.name | Remove-InvalidFileNameChars -KeepPathSpaces:($config['keepPathSpaces']['value'] -eq 2)
        $cfg['notesDirectory'] = Join-Path $NotesDestination $cfg['fileName']
        $cfg['notesBaseDirectory'] = & {
            # E.g. 'c:\temp\notes\mynotebook\mysectiongroup'
            # E.g. levelsFromRoot: 1
            $split = $cfg['notesDirectory'].Split([IO.Path]::DirectorySeparatorChar)
            # E.g. 5
            $totalLevels = $split.Count
            # E.g. 0..(5-1-1) -> 'c:\temp\notes\mynotebook'
            $split[0..($totalLevels - $cfg['levelsFromRoot'] - 1)] -join [IO.Path]::DirectorySeparatorChar
        }
        $cfg['notesDocxDirectory'] = Join-Path $cfg['notesBaseDirectory'] 'docx'
        $cfg['directoriesToCreate'] = @()

        # Build this Section Group's sections
        $cfg['sections'] = [System.Collections.ArrayList]@()
        foreach ($section in $sectionGroup.Section) {
            "$( '#' * ($LevelsFromRoot + 2) ) Building conversion configuration for $( $section.name ) [Page]" | Write-Host -ForegroundColor DarkGray

            $sectionCfg = [ordered]@{}
            $sectionCfg['object'] = $section # Keep a reference to the Section object
            $sectionCfg['kind'] = 'Section'
            $sectionCfg['levelsFromRoot'] = $cfg['levelsFromRoot'] + 1
            $sectionCfg['uri'] = $section.path # E.g. https://d.docs.live.net/01234567890abcde/Skydrive Notebooks/mynotebook/mysectiongroup/mysection
            $sectionCfg['lastModifiedTime'] = [Datetime]::ParseExact($section.lastModifiedTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
            $sectionCfg['fileName'] = $section.name | Remove-InvalidFileNameChars -KeepPathSpaces:($config['keepPathSpaces']['value'] -eq 2)
            $sectionCfg['pages'] = [System.Collections.ArrayList]@()

            # Build Section's pages
            foreach ($page in $section.Page) {
                "$( '#' * ($LevelsFromRoot + 3) ) Building conversion configuration for $( $page.name ) [Page]" | Write-Host -ForegroundColor DarkGray

                $previousPage = if ($sectionCfg['pages'].Count -gt 0) { $sectionCfg['pages'][$sectionCfg['pages'].Count - 1] } else { $null }
                $pageCfg = [ordered]@{}
                $pageCfg['object'] = $page # Keep a reference to the Page object
                $pageCfg['kind'] = 'Page'
                $pageCfg['levelsFromRoot'] = $sectionCfg['levelsFromRoot']
                # There's no $page.path property, so we generate one
                $pageCfg['uri'] = "$( $sectionCfg['object'].path )/$( $page.name )" # E.g. https://d.docs.live.net/01234567890abcde/Skydrive Notebooks/mynotebook/mysectiongroup/mysection/mypage
                $pageCfg['dateTime'] = [Datetime]::ParseExact($page.dateTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                $pageCfg['lastModifiedTime'] = [Datetime]::ParseExact($page.lastModifiedTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                $pageCfg['pageLevel'] = $page.pageLevel -as [int]
                $pageFileNameDesired = $page.name | Remove-InvalidFileNameChars -KeepPathSpaces:($config['keepPathSpaces']['value'] -eq 2)
                $pageCfg['converter'] = switch ($config['conversion']['value']) {
                    1 { 'markdown' }
                    2 { 'commonmark' }
                    3 { 'gfm' }
                    4 { 'markdown_mmd' }
                    5 { 'markdown_phpextra' }
                    6 { 'markdown_strict' }
                    default { 'markdown' }
                }
                $pageCfg['prefixjoiner'] = if ($config['prefixFolders']['value'] -eq 2) { '_' } else { '\' }
                $pageCfg['pagePrefix'] = switch ($pageCfg['pageLevel']) {
                    # process for subpage prefixes
                    1 {
                        ''
                    }
                    2 {
                        if ($null -ne $previousPage -and $previousPage['pageLevel'] -eq 1) {
                            $previousPage['fileName']
                        }else {
                            ''
                        }
                    }
                    3 {
                        if ($null -ne $previousPage -and $previousPage['pageLevel'] -eq 2) {
                            "$( $previousPage['pagePrefix'] )$( $pageCfg['prefixjoiner'] )$( $previousPage['fileName'] )"
                        }
                        # level 3 under level 1, without a level 2
                        elseif ($null -ne $previousPage -and $previousPage['pageLevel'] -eq 1) {
                            "$( $previousPage['fileName'] )$( $pageCfg['prefixjoiner'] )"
                        }
                    }
                    default {
                        ''
                    }
                }
                $pageCfg['fullexportdirpath'] = if ($pageCfg['pagePrefix'] -and $config['prefixFolders']['value'] -eq 1) {
                    Join-Path ( Join-Path $cfg['notesDirectory'] $sectionCfg['fileName'] ) $pageCfg['pagePrefix']
                }else {
                    Join-Path $cfg['notesDirectory'] $sectionCfg['fileName']
                }
                $pageCfg['fullfilepathwithoutextension'] = & {
                    $fullfilepathwithoutextensionDesired = Join-Path $pageCfg['fullexportdirpath'] $pageFileNameDesired
                    # in case multiple pages with the same name exist in a section, postfix the filename
                    $recurrence = 0
                    foreach ($p in $sectionCfg['pages']) {
                        if ($p['fullfilepathwithoutextension'] -eq $fullfilepathwithoutextensionDesired) {
                            $recurrence++
                        }
                    }
                    if ($recurrence -gt 0) {
                        "$fullfilepathwithoutextensionDesired-$recurrence"
                    }else {
                        $fullfilepathwithoutextensionDesired
                    }
                }
                $pageCfg['fileName'] = Split-Path $pageCfg['fullfilepathwithoutextension'] -Leaf
                $pageCfg['levelsPrefix'] = if ($config['medialocation']['value'] -eq 2) {
                    ''
                }else {
                    "$( '../' * ($pageCfg['levelsFromRoot'] + $pageCfg['pageLevel'] - 1) )"
                }
                $pageCfg['mediaParentPath'] = & {
                    # Normalize markdown media paths to use front slashes, i.e. '/' and lowercased drive letter
                    $s = if ($config['medialocation']['value'] -eq 2) {
                        $pageCfg['fullexportdirpath'].Replace('\', '/')
                    }else {
                        $cfg['notesBaseDirectory'].Replace('\', '/')
                    }
                    $s.Substring(0, 1).tolower() + $s.Substring(1)
                }
                $pageCfg['mediaPath'] = "$( $pageCfg['mediaParentPath'] )/media" # Normalize markdown media paths to use front slashes
                $pageCfg['fullexportpath'] = Join-Path $cfg['notesDocxDirectory'] "$( $pageCfg['fileName'] ).docx"
                $pageCfg['insertedAttachments'] = @(
                    & {
                        $pagexml = Get-OneNotePageContent -OneNoteConnection $OneNoteConnection -PageId $page.$id

                        # Get any attachment(s) found in pages
                        $insertedFiles = $pagexml.Page.Outline.OEChildren.OE | Where-Object { $null -ne $_ -and (Get-Member -InputObject $_ -Name 'InsertedFile' -Membertype Properties) } | ForEach-Object { $_.InsertedFile }
                        foreach ($i in $insertedFiles) {
                            $attachmentCfg = [ordered]@{}
                            $attachmentCfg['object'] =  $i
                            $attachmentCfg['fileName'] =  $i.preferredName | Remove-InvalidFileNameCharsInsertedFiles -KeepPathSpaces:($config['keepPathSpaces']['value'] -eq 2)
                            $attachmentCfg['markdownFileName'] =  $attachmentCfg['fileName'].Replace("$", "\$").Replace("^", "\^").Replace("'", "\'")
                            $attachmentCfg['source'] =  $i.pathCache
                            $attachmentCfg['destination'] =  Join-Path $pageCfg['mediaPath'] $attachmentCfg['fileName']

                            $attachmentCfg
                        }
                    }
                )
                $pageCfg['mutations'] = @(
                    # Markdown mutations

                    foreach ($attachmentCfg in $pageCfg['insertedAttachments']) {
                        @{
                            description = 'Change MD file Object Name References'
                            replacements = @(
                                @{
                                    searchRegex = [regex]::Escape( $attachmentCfg['object'].preferredName )
                                    replacement = "[$( $attachmentCfg['markdownFileName'] )]($( $pageCfg['mediaPath'] )/$( $attachmentCfg['markdownFileName'] ))"
                                }
                            )
                        }
                    }
                    @{
                        description = 'Replace media (e.g. images, attachments) absolute paths with relative paths'
                        replacements = @(
                            # E.g. 'c:/temp/notes/mynotebook/media/image1.jpg' -> '../media/image1.jpg'
                            @{
                                searchRegex = [regex]::Escape("$( $pageCfg['mediaParentPath'] )/")
                                replacement = $pageCfg['levelsPrefix']
                            }
                        )
                    }
                    @{
                        description = 'Add heading'
                        replacements = @(
                            @{
                                searchRegex = '^[^\r\n]*'
                                replacement = & {
                                    $heading = "# $( $pageCfg['object'].name )"
                                    if ($config['headerTimestampEnabled']['value'] -eq 1) {
                                        $heading += $pageCfg['dateTime'].ToString("`n`nyyyy-MM-dd HH:mm:ss")
                                        $heading += "`r`n`r`n---`r`n"
                                    }
                                    $heading
                                }
                            }
                        )
                    }
                    if ($config['keepspaces']['value'] -eq 1 ) {
                        @{
                            description = 'Clear double spaces from bullets and nonbreaking spaces from blank lines'
                            replacements = @(
                                @{
                                    searchRegex = [regex]::Escape([char]0x00A0)
                                    replacement = ''
                                }
                                @{
                                    searchRegex = "`r?`n`r?`n- "
                                    replacement = "`r`n`- "
                                }
                            )
                        }
                    }
                    if ($config['keepescape']['value'] -eq 1) {
                        @{
                            description = 'Clear backslash escape symbols'
                            replacements = @(
                                @{
                                    searchRegex = [regex]::Escape('\')
                                    replacement = ''
                                }
                            )
                        }
                    }
                )

                $pageCfg['directoriesToCreate'] = @(
                    # The directories to be created
                    $cfg['notesDocxDirectory']
                    $cfg['notesDirectory']
                    $pageCfg['fullexportdirpath']
                    if ($pageCfg['pagePrefix'] -and $config['prefixFolders']['value'] -eq 2) {
                        Join-Path $pageCfg['fullexportdirpath'] $pageCfg['pagePrefix']
                    }
                    $pageCfg['mediaPath']
                )

                # Populate the pages array
                $sectionCfg['pages'].Add( $pageCfg ) > $null
            }

            # Populate the sections array
            $cfg['sections'].Add( $sectionCfg ) > $null
        }

        # Build this Section Group's Section Groups
        if ($sectiongroup.SectionGroup) {
            # $sectionGroupName = $sectionGroup.Name | Remove-InvalidFileNameChars -KeepPathSpaces:($config['keepPathSpaces']['value'] -eq 2)
            $cfg['sectionGroups'] = New-SectionGroupConversionConfig -OneNoteConnection $OneNote -NotesDestination $cfg['notesDirectory'] -Config $Config -SectionGroups $sectiongroup.SectionGroup -LevelsFromRoot ($LevelsFromRoot + 1)
        }

        # Populate the conversion config
        $sectionGroupConversionConfig.Add( $cfg ) > $null
    }

    # Return the final conversion config
    ,$sectionGroupConversionConfig # This syntax is needed to send an array down the pipeline without it being unwrapped. (It works by wrapping it in an array with a null sibling)
}

Function Convert-OneNoteSection {
    [CmdletBinding()]
    param (
        # Onenote connection object
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    ,
        # ConvertOneNote2MarkDown configuration object
        [Parameter(Mandatory)]
        [object]
        $Config
    ,
        # Conversion object
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $ConversionConfig
    )

    foreach ($pageCfg in $ConversionConfig['pages']) {
        try {
            "$( '#' * ($pageCfg['levelsFromRoot'] + $pageCfg['pageLevel']) ) $( $pageCfg['object'].name ) [$( $pageCfg['kind'] )]" | Write-Host
            "Uri: $( $pageCfg['uri'] )" | Write-Verbose

            # Create directories
            foreach ($d in $pageCfg['directoriesToCreate']) {
                try {
                    $item = New-Item -Path $d -ItemType Directory -Force -ErrorAction Stop
                    "Directory: $( $item.FullName )" | Write-Verbose
                }catch {
                    throw "Failed to create directory '$d': $( $_.Exception.Message )"
                }
            }

            if ($config['usedocx']['value'] -eq 1) {
                # Remove any existing docx files, don't proceed if it fails
                try {
                    Remove-Item -path $pageCfg['fullexportpath'] -Force -ErrorAction SilentlyContinue
                    "Removing existing docx file: $( $pageCfg['fullexportpath'] )" | Write-Verbose
                }catch {
                    throw "Error removing intermediary '$( $pageCfg['object'].name )' docx file: $( $_.Exception.Message )"
                }
            }

            # Publish OneNote page to Word, don't proceed if it fails
            if (! (Test-Path $pageCfg['fullexportpath']) ) {
                try {
                    $OneNoteConnection.Publish($pageCfg['object'].ID, $pageCfg['fullexportpath'], "pfWord", "")
                    "New docx file: $( $pageCfg['fullexportpath'] )" | Write-Verbose
                }catch {
                    throw "Error while publishing file '$( $pageCfg['object'].name )' to docx: $( $_.Exception.Message )"
                }
            }else {
                "Existing docx file: $( $pageCfg['fullexportpath'] )" | Write-Verbose
            }

            # https://gist.github.com/heardk/ded40b72056cee33abb18f3724e0a580
            # Convert .docx to .md, don't proceed if it fails
            try {
                pandoc.exe -f  docx -t "$( $pageCfg['converter'] )-simple_tables-multiline_tables-grid_tables+pipe_tables" -i $pageCfg['fullexportpath'] -o "$( $pageCfg['fullfilepathwithoutextension'] ).md" --wrap=none --markdown-headings=atx --extract-media="$( $pageCfg['mediaParentPath'] )" # extracts into ./media of the supplied folder
                "Converting docx file to markdown file: $( $pageCfg['fullfilepathwithoutextension'] ).md" | Write-Verbose
            }catch {
                throw "Error while converting file '$( $pageCfg['object'].name )' to md: $( $_.Exception.Message )"
            }

            # Cleanup Word files
            if ($config['keepdocx']['value'] -eq 1) {
                try {
                    Remove-Item -path $pageCfg['fullexportpath'] -Force -ErrorAction Stop
                    "Removing existing docx file: $( $pageCfg['fullexportpath'] )" | Write-Verbose
                }catch {
                    Write-Error "Error removing intermediary '$( $pageCfg['object'].name )' docx file: $( $_.Exception.Message )"
                }
            }

            # Export any attachments
            foreach ($attachmentCfg in $pageCfg['insertedAttachments']) {
                try {
                    Copy-Item -Path $attachmentCfg['source'] -Destination $attachmentCfg['destination'] -Force
                    "Saving inserted attachment: $( $attachmentCfg['destination'] )" | Write-Verbose
                }catch {
                    Write-Error "Error while copying file object '$($pageinsertedfile.InsertedFile.preferredName)' for page '$( $pageCfg['object'].name )': $( $_.Exception.Message )"
                }
            }

            # rename images to have unique names - NoteName-Image#-HHmmssff.xyz
            $timeStamp = (Get-Date -Format HHmmssff).ToString()
            $timeStamp = $timeStamp.replace(':', '')
            $images = Get-ChildItem -Path "$( $pageCfg['mediaPath'] )" -Include "*.png", "*.gif", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name.SubString(0, 5) -match "image" }
            foreach ($image in $images) {
                $newimageName = "$($( $pageCfg['fileName'] ).SubString(0,[math]::min(30,$( $pageCfg['fileName'] ).length)))-$($image.BaseName)-$($timeStamp)$($image.Extension)"
                # Rename Image
                try {
                    $item = Rename-Item -Path "$( $image.FullName )" -NewName $newimageName -ErrorAction SilentlyContinue -PassThru
                    "Renaming image: $( $image.FullName ) to $( $item.FullName )" | Write-Verbose
                }catch {
                    Write-Error "Error while renaming image $( $image.FullName ) to $( $item.FullName ) for page '$( $pageCfg['object'].name )': $( $_.Exception.Message )"
                }
                # Change MD file Image filename References
                try {
                    ((Get-Content -path "$( $pageCfg['fullfilepathwithoutextension'] ).md" -Raw).Replace("$($image.Name)", "$($newimageName)")) | Set-Content -Path "$( $pageCfg['fullfilepathwithoutextension'] ).md"
                    "Mutation of markdown: Rename image references to unique name" | Write-Verbose
                }catch {
                    Write-Error "Error while renaming image file name references to '$( $newimageName )' for file '$( $pageCfg['object'].name )': $( $_.Exception.Message )"
                }
            }

            # Get markdown content
            $orig = @( Get-Content -path "$( $pageCfg['fullfilepathwithoutextension'] ).md" )
            $orig = @(
                if ($orig.Count -gt 6) {
                    # Discard first 6 lines which contain a header, created date, and time. We are going to add our own header
                    $orig[6..($orig.Count - 1)]
                }else {
                    # Empty page
                    ''
                }
            ) -join "`r`n"

            # Perform mutations on markdown content
            foreach ($m in $pageCfg['mutations']) {
                "Mutation of markdown: $( $m['description'] )" | Write-Verbose
                foreach ($r in $m['replacements']) {
                    $orig = $orig -replace $r['searchRegex'], $r['replacement']
                }
            }
            Set-Content -Path "$( $pageCfg['fullfilepathwithoutextension'] ).md" -Value $orig
            "Markdown file ready: $( $pageCfg['fullfilepathwithoutextension'] ).md" | Write-Verbose
        }catch {
            Write-Error "Failed to convert page '$( $pageCfg['uri'] )'"
        }
    }
}

Function Convert-OneNoteSectionGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    ,
        # ConvertOneNote2MarkDown configuration object
        [Parameter(Mandatory)]
        [object]
        $Config
    ,
        # Conversion object
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $ConversionConfig
    ,
        [Parameter()]
        [switch]
        $Recurse

    )

    foreach($cfg in $ConversionConfig) {
        if ($cfg['levelsFromRoot'] -eq 0) {
            "Notebook: $( $cfg.object.name )" | Write-Host -ForegroundColor Green
        }
        foreach ($sectionCfg in $cfg['sections']) {
            "$( '#' * ($sectionCfg['levelsFromRoot']) ) $( $sectionCfg['object'].name ) [$( $sectionCfg['kind'] )]" | Write-Host
            "Uri: $( $sectionCfg['uri'] ) [$( $sectionCfg['kind'] )]" | Write-Verbose
            Convert-OneNoteSection -OneNoteConnection $OneNoteConnection -Config $Config -ConversionConfig $sectionCfg
        }

        if ($Recurse) {
            foreach ($sectionGroupCfg in $cfg['sectionGroups']) {
                "`n$( '#' * ($sectionGroupCfg['levelsFromRoot']) ) $( $sectionGroupCfg['object'].name ) [$( $sectionGroupCfg['kind'] )]" | Write-Host
                "Uri: $( $sectionGroupCfg['uri'] )" | Write-Verbose
                Convert-OneNoteSectionGroup -OneNoteConnection $OneNoteConnection -Config $Config -ConversionConfig $sectionGroupCfg -Recurse
            }
        }
    }
}

Function Print-ConversionErrors {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorCollection
    )

    if ($ErrorCollection.Count -gt 0) {
        "Conversion errors: " | Write-Host
        $ErrorCollection | Where-Object { (Get-Member -InputObject $_ -Name 'CategoryInfo' -Membertype Properties) -and ($_.CategoryInfo.Reason -eq $ExceptionName) } | Write-Host
    }
}

Function Convert-OneNote2MarkDown {
    [CmdletBinding()]
    param ()

    try {
        # Fix encoding problems for languages other than English
        $PSDefaultParameterValues['*:Encoding'] = 'utf8'

        $totalerr = @()

        # Validate dependencies
        Validate-Dependencies

        # Compile and validate configuration
        $config = Compile-Configuration | Validate-Configuration

        # Connect to OneNote
        $OneNote = New-OneNoteConnection

        # Get the hierarchy of OneNote objects as xml
        $hierarchy = Get-OneNoteHierarchy -OneNoteConnection $OneNote

        # Get and validate the notebook(s) to convert
        $notebooks = @(
            if ($config['targetNotebook']['value']) {
                $hierarchy.Notebooks.Notebook | Where-Object { $_.Name -eq $config['targetNotebook']['value'] }
            }else {
                $hierarchy.Notebooks.Notebook
            }
        )
        if ($notebooks.Count -eq 0) {
            if ($config['targetNotebook']['value']) {
                throw "Could not find notebook of name '$( $config['targetNotebook']['value'] )'"
            }else {
                throw "Could not find notebooks"
            }
        }

        # Build a conversion configuration of notebook(s) (i.e. a conversion object representing the conversion)
        "Building a conversion configuration. This might take a while if you have many large notebooks..." | Write-Host -ForegroundColor Cyan
        $conversionConfig = New-SectionGroupConversionConfig -OneNoteConnection $OneNote -NotesDestination $config['notesdestpath']['value'] -Config $config -SectionGroups $notebooks -LevelsFromRoot 0 -ErrorVariable +totalerr
        "Done building conversion configuration." | Write-Host -ForegroundColor Cyan

        # Convert the notebook(s)
        "Converting notes..." | Write-Host -ForegroundColor Cyan
        Convert-OneNoteSectionGroup -OneNoteConnection $OneNote -Config $config -ConversionConfig $conversionConfig -Recurse -ErrorVariable +totalerr
        "Done converting notes." | Write-Host -ForegroundColor Cyan
    }catch {
        throw
    }finally {
        'Cleaning up...' | Write-Host -ForegroundColor Cyan

        # Disconnect OneNote connection
        if (Get-Variable -Name OneNote -ErrorAction SilentlyContinue) {
            Remove-OneNoteConnection -OneNoteConnection $OneNote
        }

        # Print any conversion errors
        Print-ConversionErrors -ErrorCollection $totalerr
        'Exiting.' | Write-Host -ForegroundColor Cyan
    }
}

# Entrypoint
Convert-OneNote2MarkDown

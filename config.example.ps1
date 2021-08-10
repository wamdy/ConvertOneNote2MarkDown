#
# Note: This config file is for those who are lazy to type in configuration everytime you run ./ConvertOneNote2MarkDown-v2.ps1
#
# Steps:
#   1) Rename this file to config.ps1. Ensure it is in the same folder as the ConvertOneNote2MarkDown-v2.ps1 script
#   2) Configure the options below to your liking
#   3) Run the main script: ./ConvertOneNote2MarkDown-v2.ps1. Sit back while the script starts converting immediately.

# Whether to do a dry run
# 1: Convert - Default
# 2: Convert (dry run)
$dryRun = 1

# Specify folder path that will contain your resulting Notes structure - Default: c:\temp\notes
$notesdestpath = 'c:\temp\notes'

# Specify a notebook name to convert
# '': Convert all notebooks - Default
# 'mynotebook': Convert specific notebook named 'mynotebook'
$targetNotebook = ''

# Whether to create new word .docx or reuse existing ones
# 1: Always create new .docx files - Default
# 2: Use existing .docx files (90% faster)
$usedocx = 1

# Whether to discard word .docx after conversion
# 1: Discard intermediate .docx files - Default
# 2: Keep .docx files
$keepdocx = 1

# Whether to use prefix vs subfolders
# 1: Create folders for subpages (e.g. Page\Subpage.md) - Default
# 2: Add prefixes for subpages (e.g. Page_Subpage.md)
$prefixFolders = 1

# Whether to store media in single or multiple folders
# 1: Images stored in single 'media' folder at Notebook-level - Default
# 2: Separate 'media' folder for each folder in the hierarchy
$medialocation = 1

# Specify conversion type
# 1: markdown (Pandoc) - Default
# 2: commonmark (CommonMark Markdown)
# 3: gfm (GitHub-Flavored Markdown)
# 4: markdown_mmd (MultiMarkdown)
# 5: markdown_phpextra (PHP Markdown Extra)
# 6: markdown_strict (original unextended Markdown)
$conversion = 1

# Whether to include page timestamp and separator at top of document
# 1: Include - Default
# 2: Don't include
$headerTimestampEnabled = 1

# Whether to clear double spaces between bullets
# 1: Clear double spaces in bullets - Default
# 2: Keep double spaces
$keepspaces = 1

# Whether to clear escape symbols from md files
# 1: Clear '\' symbol escape character from files - Default
# 2: Keep '\' symbol escape
$keepescape = 1

# Whether to use Line Feed (LF) or Carriage Return + Line Feed (CRLF) for new lines
# 1: LF (unix) - Default
# 2: CRLF (windows)
$newlineCharacter = 1

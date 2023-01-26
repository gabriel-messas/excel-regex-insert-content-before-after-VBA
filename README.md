# excel-regex-insert-content-before-after-VBA

1. Press ALT + F11 in Excel to open the VBA console
2. Copy and add this module
3. The function will be available in the Workbook

Function RegxFunc takes as arguments:
- strInput As String => The content to be matched
- regexPattern As String => The pattern to be matched against
- insertBefore As String => The content to be inserted before the matched text
- insertAfter As String => The content to be inserted after the matched text

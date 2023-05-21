'==========================================================================
'                            Message Box
'==========================================================================

'Syntax: MsgBox "text" [, buttons] [,title]
'vbOKOnly 0
'vbOKCancel 1
'vbAbortRetryIgnore 2
'vbYesNoCancel 3
'vbYesNo 4
'vbRetryCancel 5
'vbCritical 16 - Display Critical Message Icon
'vbQuestion 32 - Display Warning Query Icon
'vbWarning 48 - Display Warning Message Icon
'vbInformation 64 - Display Information Message Icon

'You can sum the desired button/message and add in the buttons position
'Example: 3 + 48
MsgBox "Test Warning", 51, "Warning box with Yes No Cancel buttons"

'Example: 4 + 32
MsgBox "Is it a question?", 36, "Question box with Yes No buttons"

'Example: 2 + 16
MsgBox "Fatal Error", 18, "Critical box with Abort Retry Ignore buttons"
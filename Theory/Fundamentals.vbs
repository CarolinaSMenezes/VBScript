
'==========================================================================
'                      Variables and Constants
'==========================================================================

'Variables, Constants, Dim
'Dim keyword is used to declarate a variable
Dim value1, value2, sum

'Const keyword is used to declarate a constant
Const tax = 1

'Add values to the variables
value1 = 10
value2= 20
sum = value1 + value2 + tax

'MsgBox is used to print messages on a box in the screen
'& is used to concatenate Strings
MsgBox  "The final value of the sum is: " & sum

'.Sleep 3000 makes the code wait for 3 seconds before showing up the second msg box.
WScript.Sleep 3000 

'&_ is used to concatenate lines.
MsgBox "This is a test with sleep and line concatenation. " &_
"The messages will appear together"




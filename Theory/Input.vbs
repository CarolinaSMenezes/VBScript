'==========================================================================
'                                User Input
'==========================================================================
Dim name, age, city
Const additionalValue = 10

' --- First test ---
'name = InputBox("Enter your name: ")
'age = InputBox("Enter your age: ")
'MsgBox "Welcome " & name & "!!! Your age in 10 years will be " & age & ".", 0 , "Welcome Message"

'Add title and default text
name = InputBox("Enter your name: ", "Title", "Enter your name here:" )

'Move the input box around
age = InputBox("Enter your age: ", "Title", "Enter your name here.", 5000, 5000)

city = InputBox("Please inform your city: ", "City Name", "Enter the name of your city here.", 1000, 2000)

'CInt is used to check if the input is an Integer. It will convert the value of the string into an int.
age = CInt(age) + additionalValue

MsgBox "Welcome " & name & "!!! In 10 years you will be " & age & " years old. You will still live in " & city & ".", 0 , "Welcome Message"


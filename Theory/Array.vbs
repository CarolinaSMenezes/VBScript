'==========================================================================
'                                Arrays
'==========================================================================
'Declare the array. We can inform the maximum index in the parameter - it starts in 0 and finishes in the informed index.
Dim arrNames()

'If you do not declare a fixed size, ReDim can be used to resize the array
ReDim arrNames(4)
arrNames(0) = "Carolina"
arrNames(1) = "Carolina2"
arrNames(2) = "Carolina3"
arrNames(3) = "Carolina4"
arrNames(4) = "Carolina5"

'WScript.Echo prints the message in the debug console
WScript.Echo arrNames(0)
WScript.Echo arrNames(1)

Dim inputName1, inputName2
inputName1 = InputBox ("Insert the first name", "Insert First Name", "Insert a name here")
inputName2 = InputBox ("Insert the second name", "Insert Second Name", "Insert a name here")

arrNames(0) = inputName1 
arrNames(1) = inputName2

MsgBox "The first Name inserted was " & arrNames(0) & ". The second name inserted was " & arrNames(1) &_
". Another name in the array is " & arrNames(2) & "!"

'The issue when using only ReDim, is that you will lose all information in the existing positions in the array
ReDim arrNames(5)
arrNames(5) = "Additional Name"

WScript.Echo "New test printing the array"
WScript.Echo arrNames(0)
WScript.Echo arrNames(1)
WScript.Echo arrNames(5)

'ReDim Preserve save the data in an existing array when you change the size of the last dimension
ReDim Preserve arrNames(6)
arrNames(0) = "Carolina"
arrNames(1) = "Carolina2"
arrNames(2) = "Carolina3"
arrNames(3) = "Carolina4"
arrNames(4) = "Carolina5"

WScript.Echo "New test printing the array with ReDim Preserve"
WScript.Echo arrNames(0)
WScript.Echo arrNames(1)
WScript.Echo arrNames(5)

Dim arrAnimals, animal
'add values in the array at once
arrAnimals = Array("Snake", "Bird", "Dog", "Cat", "Cow", "Lion")
animal = arrAnimals(2)

MsgBox "The animal in the third position is:" & animal, 0 , "Animal's Array"

'Split is used to separate a string by the char defined
Const webSite = "www.google.com"

'The variable above is not an array yet.
Dim arrTitle

'After I use the Split function, and use the character "." to delimitate de String, the variable will become an array.
arrTitle = Split(webSite, ".")

WScript.Echo "Using Split"
WScript.Echo arrTitle(0)
WScript.Echo arrTitle(1)
WScript.Echo arrTitle(2)

'The Join function is used if you want to unify the array positions into 1 string again
Dim siteURL
siteURL = Join(arrTitle, ".")
WScript.Echo "Using Join"
WScript.Echo siteURL


'LBound and UBound are very useful functions. LBound returns the lower index of an array, and the UBound return the upper index 
' of an array.
Dim lowerIndex, upperIndex

lowerIndex = LBound(arrNames)
WScript.Echo lowerIndex

upperIndex = UBound(arrNames)
WScript.Echo upperIndex

'IsArray function is used to returns a bolean value if a determinated variable is an array or not.
Dim arrNums(2)
Dim numberTest

'If using WScript.Echo it return -1 and 0.
MsgBox "IsArray funcion - arrNums: " & IsArray(arrNums) & ". IsArray funcion - numberTest: " & IsArray(numberTest)

'Erase Function is used when you want to erase information in an array
WScript.Echo "Erase Function checkpoint 1"
WScript.Echo arrAnimals(1)
WScript.Echo "Erase Function checkpoint 2 - Erased"
erase arrAnimals
'Error message expected
'WScript.Echo arrAnimals(1)

'Creating Filters
Dim matched, notMatched, item

'Filter (array, criteria, determine if you want return what matchs or does not match(true/false), Binary or Text(0/1))
matched = Filter(arrNames, "arol", True, 1)
notMatched = Filter(arrNames, "arol", false, 1)

WScript.Echo "Matched values"
WScript.Echo matched(0)
WScript.Echo matched(1)
WScript.Echo matched(2)

WScript.Echo "Not matched values"
WScript.Echo notMatched(0)



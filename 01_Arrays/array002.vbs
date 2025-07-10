'====================================================================================================================
'Arrays

'====================================================================================================================
' Error Handling
Option Explicit

'variables
Const SITE_TILE = "www.MohanVBScriting.com"

'Array
Dim arrAnimals
Dim disAnimalname

'Assiging value to the Array

arrAnimals = Array("Lion","Tiger","Elephant","Monkey","Mohan", "Frog")

'The value of an element in the array is assigned to a variable
disAnimalname = arrAnimals


'Display the value
MsgBox "The Animal Name is  " & disAnimalname(4), 1, SITE_TILE
# VBA-Strings
A whole bunch of String functions to make it easier and faster coding.

# Public Funtions:
- StringSimilarity
- LevenshteinDistance
- StringInterpolation (also under alias Inject)
- Truncate
- StringBetween
- StringProperLength

# Usage

Import StringFunctions.bas file.
Set Reference to MicroSoft Scripting Runtime (for using Scripting.Dictionary)

Below are some of the examples you can do with single dim arrays. Note, there are several functions for two dim arrays as well.

```vb
Private Sub StringFunctionExamples()
    
    StringSimilarity "Test", "Tester"        '->  66.6666666666667
    LevenshteinDistance "Test", "Tester"     '->  2
    StringInterpolation "${0}\n\t${1}", "First", "Tab and Second" '-> First
                                                                  '->   Tab and Second
                                                                  
    Truncate "This is a long sentence", 10                '-> "This is..."
    StringBetween "Robert Paul Todar", "Robert", "Todar"  '-> "Paul"
    StringProperLength "1001", 6, "0", True               '-> "100100"
    
    
    'Inject is a copy of StringInterpolation, this alias is easier to remember (shorter too!)
    'Here is an example using a dictionary!
    Dim Person As New Scripting.Dictionary
    Person("Name") = "Robert"
    Person("Age") = 30
    
    'REMEMBER, DICTIONARY KEYS ARE CASE SENSITIVE!
    Debug.Print Inject("Hello,\nMy name is ${Name} and I am ${Age}!", Person)
        '-> Hello,
        '-> My name is Robert and I am 30!
End Sub
```

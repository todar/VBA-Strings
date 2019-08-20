# VBA String Functions
A library of String functions to make life easier! =)

## References
| Reference                                  | Object               |
| ------------------------------------------ | -------------------- |
| Microsoft Scripting Runtime                | Scripting.Dictionary |
| Microsoft VBScript Regular Expressions 5.5 | RegExp, Match        |

## Public Funtions

| Function            | Description                                                                              |
| ------------------- | ---------------------------------------------------------------------------------------- |
| StringSimilarity    | This returns a percentage of how similar two strings are using the levenshtein formula.  |
| LevenshteinDistance | The distance between two sequences of words.                                             |
| Inject              | Returns a new cloned string that replaced special {keys} with its associated pair value. |
| Truncate            | Create a max lenght of string and return it with extension.                              |
| StringBetween       | Find string between two words.                                                           |
| StringPadding       | Returns a string with the proper padding on either side.                                 |
| ToString            | Reads any value or object in VBA and returns it in string formatting.                    |

## Example Usage

```vb
'/**
' * Examples of various functions.
' *
' * @author Robert Todar <robert@robertodar.com>
' * @licence MIT
' */
Private Sub testsForStringFunctions()
    Debug.Print StringSimilarity("Test", "Tester")                     '~>  66.6666666666667
    Debug.Print LevenshteinDistance("Test", "Tester")                  '~>  2
    Debug.Print Truncate("This is a long sentence", 10)                '~> "This is..."
    Debug.Print StringBetween("Robert Paul Todar", "Robert", "Todar")  '~> "Paul"
    Debug.Print StringPadding("1001", 6, "0", True)                    '~> "100100"
    Debug.Print Inject("Hello,\nMy name is {Name} and I am {Age}!", "Robert", 31)
        '~> Hello,
        '~> My name is Robert and I am 30!
        
    Debug.Print ToString(Array(1, 2, 3, 4)) '~> [1, 2, 3, 4]
End Sub
```

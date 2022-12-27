--8<-- "../README.md"

## VBA Project Set-up
Each VBA project should include

1. [Config Module](#1-config-module)
2. [Main Module](#2-main-module)
3. [Utils Module (optional)](#3-utils-module-optional)

### 1. Config Module
This module stores meta data of the vba project
???+ note "Inside Config Module"
    Note we use `CAPITALS` for all values as they are constants
    ???+ abstract "What Config module looks like"
        ```vba
        'inside the config module

        Public Const VERSION = "<version>"
        Public Const NAME = "<name>"
        Public Const EMAIL = "<email>"
        Public Const DESCRIPTION = "<example>"
        Public Const REPO = "<repo url>"
        Public Const REPO ADO_FEATURE_NUMBER = "<ADO_number>"
        'feel free to add other meta data you require
        ```

    ??? abstract "Example Config Module"
        ```vba
        'inside the config module
        Public Const VERSION = "0.1.0"
        Public Const NAME = "Ish Hassan"
        Public Const EMAIL = "ihassan@example.com"
        Public Const DESCRIPTION = "export and import VBA code for git versioning"
        Public Const REPO = "https://github.com/deltahedge1/vba-bestpracices"
        Public Const Repo ADO_FEATURE_NUMBER = "<ADO_number>"
        ```

### 2. Main Module
In the main module you should have a `main` subroutine which this logic for the whole solution.
    
???+ note "What is the Main module"

    Understanding the `main` module is best illustrated through an example.

    Say you had a solution to import and transform some data from another workbook.
    
    The steps would be:
    
    1. select sheet to get data from
    2. run validations on data
    3. raise exceptions if any
    4. if no exceptions copy the data
    5. transform the data (e.g. fill blanks)

    ??? abstract "Example Main Module"
        ```vba
        Public Sub Main
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Get data from another sheet and transform it by replacing blanks
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
            
            'subroutines below would be in there own modules
            call utils.get_data
            call utils.validate_data
            call utils.copy_data_to_sheet
            call utils.transform_data

        End Sub
        ```

    **Each step above could be modularised into its own subroutines in another module like `utils` but would be combined in `main` as they all together create the solution.**

    ??? tip "Why do we modularise code into `main` and `other modules`"
        We modularise the code for two reasons. 

        1. More maintainable
            - if your re-using a `function` and there is an error if you abstracted using a `function` you would only need to change it once in that that `function`.
        2. Easier to preform unit tests
            - if all you code is in one long subroutine you would struggle to find which part of the code the error was in, and to test it you would need to run everything before it.

### 3. Utils Module (optional)
This is the module where you can modularise `subroutines` and `functions` which will be used in `main` module

Feel free to add more modules that just `Utils` if it helps manage your code better :smile:

???+ note "Understanding the Utils Module"

    Typically you can put subroutines or functions that need to be used mutiple times. This could be like find the last row in a column or check it a column contains blanks.

    These can then we re-used mutiple time through out the project, and if there is a bug you only need to fix it in one place.

    ??? abstract "Example Utils module"
        ```vba
        'inside utils module

        Sub copy_range_to_another_sheet(from_sheet as worksheet, to_sheet as worksheet)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Copies data from one worksheet to another worksheet
        '
        'ARGS:
        '   from_sheet (worksheet obj): sheet to copy from
        '   to_sheet (worksheet obj): sheet to copy to
        '
        'EXAMPLE:
        '   call copy_range_to_another_sheet activeworkbook.sheets("sheet1") _ 
        '       activeworkbook.sheets("sheet2") 
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
            from_sheet("Sheet1").Range("A1:B10").Copy Destination:=to_sheet("Sheet2").Range("E1")

        End Sub
        ```
    ??? tip "Why do we modularise code into `Utils`"
        We modularise the code for two reasons, same reasons we did it for `Main` module above (are you suprised!)

        1. More maintainable
            - if your re-using a `function` and there is an error if you abstracted using a `function` you would only need to change it once in that that `function`.
        2. Easier to preform unit tests
            - if all you code is in one long subroutine you would struggle to find which part of the code the error was in, and to test it you would need to run everything before it.

## Moving your VBA code to Production
[reference](https://www.spreadsheet1.com/vba-development-best-practices.html)

1. [Documentation](#1-documentation)
2. [Versioning](#2-versioning)
3. [Testing and Testing documentation](#3-testing-and-testing-documentation)
4. [use git to version control](#4-use-git-for-version-control)
5. [Change Log](#5-change-log)
6. [README.md](#6-readmemd)

### 1. Documentation

There are two types of documentation required for code

??? note "1. Internal documentation"
    This documentation lives inside the code and is meant to explain to developers what the code is intended for.

    It includes the following:

    - description
    - arguments and arugments types
    - exceptions raised
    - returns
    - examples
    - references

    === "Subroutine"
        ```
        Sub anonomyse_email(ByRef check_range As Range, _
                Optional ByVal match_str As Variant = "@", _
                Optional ByVal replace_value As Variant = "abc@gmail.com")
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Find emails in the range and then anonymse them with another email
        '
        'ARGS:
        '   check_range (obj:Range): the range to check for emails
        '   Optional match_str (String): string to match against to check if value is an email
        '   Optional replace_value (String): replace the email with this string
        '
        'EXAMPLES:
        '   call with defaults:
        '       call anonomyse_email(activesheet.usedrange)
        '
        '   change anonomysed email:
        '       call anonomyse_email(activesheet.usedrange, "@", "hello@gmail.com") 'change the anonmysed
        '
        '   replace only the gmails with another mail
        '       call anonomyse_email(activesheet.usedrange, "@gmail", "hello@yahoo.com")
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim cell As Range
        For Each cell In check_range
            If InStr(cell.Value2, match_str) > 0 Then
                cell.Value2 = replace_value
            End If
        Next cell
            
        End Sub
        ```

    === "Function"
        ```
        Function add_2(byval num1 as long, byval num2 as long) as long
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'adds to numbers together
        '
        'ARGS:
        '   num1 (long): first number to add
        '   num2 (long): second number to add
        '
        'RETURNS:
        '   long: the total of the two numbers
        '
        'EXAMPLE:
        '   add_2(1, 3) => 4
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        add_2 = num1 + num2
        End Function
        ```

??? note "2. External documentation"
    This documentation lives outside the code and can go into much more depth than internal documentation.

    This documentation usually includes areas suchas :

    1. How to install
    2. Limitations
    3. Quick start guides

### 2. Versioning

Your code should contain at a minimum versioning in the [Config Module](#1-config-module).

Reccomend to use the semantic versioning system.
![semantic versioning explained](https://media.geeksforgeeks.org/wp-content/uploads/semver.png)

### 3. Testing and Testing Documentation

Your code should include unit tests, integration tests, and testing documentation.

Considerations

1. Testing as many inputs and expected outputs for your subroutines and functions
2. Testing any expected exceptions e.g. exception handlers
3. Testing the combination of subroutines and functions integrating together
4. Thinking about the edge cases (as developers its upto to design solutions for the edge cases)
4. Take a risk based approach what is the most probable what could go wrongs and focus on them but also sprinkle some edge cases in there and depending on if the edge case can have serious consoquences testing them vigirously
5. Try to cover as many cases, and combinations as possible and not to over test
5. Testing is hard but neccesiary you will need to use your professional judgement

???+ danger "If your testing is not documented then it is not done"
    Using a test matrix which shows scope on the columns and test cases on the rows is my recomendation

### 4. Use git for version control
Export your code to normal txt files so you can version control them with git.

I know if it annoying but it helps when we collaborate and if we need to roll back to previous versions.

### 5. Change Log

Create a CHANGELOG.md or HISTORY.md in the root directory with

1. version
2. date
3. change in the release 

This will help you so much in the future, you can thank me then. Unless you forgot in which case you will hate yourself!

??? note "Example CHANGELOG.md"

    ```markdown
    # Release History
    
    ## dev
    small changes to the validation
    
    ## 2.28.1 (2022-06-29)
    Improvements

    Speed optimization in `iter_content` with transition to `yield from`. (#6170)
    Dependencies

    Added support for chardet 5.0.0 (#6179)
    Added support for charset-normalizer 2.1.0 (#6169)
    ```

### 6. README.md

create a nice `REAMDE.MD` in the root directory of your project.

Include some highlevel material only. Details should be in the documentation

## Best Practices
1. [Option Explicit](#1-option-explicit)
2. [Full declare variables](#2-full-declare-variables)
3. [Naming convention](#3-naming-convention)
4. [Code commenting](#4-code-commenting)
5. [Turning off ScreenUpdating and Automatic calculation](#5-turning-off-screenupdating-and-automatic-calculation)
6. [Explicitly call the default property of an object](#6-explicitly-call-the-default-property-of-an-object)
7. [Error Handling](#7-error-handling)
8. [Code Readability](#8-code-readability)
9. [Use "" instead of vbNullString](#9-use--instead-of-vbnullstring)
4. [Structure modular code (abstraction)](#4-structure-modular-code-abstraction)

### 1. Option Explicit
### 2. Full declare variables
### 3. Naming convention
### 4. Code commenting
### 5. Turning off ScreenUpdating and Automatic calculation
### 6. Explicitly call the default property of an object
### 7. Error Handling
### 8. Code Readability
### 9. Use "" instead of vbNullString
### 4. Structure modular code (abstraction)
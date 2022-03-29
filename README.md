<div id="repo-header" align="center">
 <img id="repo-logo" alt="repo-logo" src="https://user-images.githubusercontent.com/93403866/160690956-1b725bcc-38af-4316-875c-bd197719b503.png" height="60"></img>
  <h1  id="repo-title"> I 💙 VBScript </h1>
  <h3  id="repo-description" > Template Repository for VBScript Projects </h3>
  <hr>
</div>
<div id="repo-shields"align="center">

</div>

## Table of Contents

- [Introduction](#introduction)
    - [Project Description](#project-description)
    - [Project Goals](#project-goals)
- [Project Details](#project-details)
    - [Prerequisites](#prerequisites)
    - [Features](#features)
    - [Future Improvements](#future-improvements)
- [Installation and Usage](#installation-and-usage)
    - [Installation](#installation)
    - [Usage](#usage)
- [About the Project](#about-the-project)
    - [Code Overview](#code-overview)
- [Licensing and Credits](#licensing-and-credits)
    - [Licensing](#licensing)
    - [Credits](#credits)

## Introduction
### Project Description
I 💙 VBScript is a project that aims to ease the process of creating VBScript code and refine some overly complex VBScript syntax with easier to digest functions and subroutines. Despite its constant hate and its age, VBScript is still a powerful language which is still used to this day by many in many applications. A lot of these applications are rid with old outdated VBScript code that can't handle exceptions and lacks support of modern features. Many VBScript processes are abandonware but some are still in use, even in major companies, but they are very hard to refactor as many don't feel confident to touc. This project aims to solve these problems by providing a set of functions and subroutines that can be used to create VBScript code.

### Project Goals
The project has the following goals:

- **Ease VBScript** : Many people know VBScript as a very pointless language. This project aims to make it easier to use VBScript.
- **Productivity** : I want to be able to write code that is easy to read, maintain, and extend, that follows a template.
- **Collect Useful Functions** : Collect useful functions for VBScript development

## Project Details
### Prerequisites
This project requires the following prerequisites:

- **Operative System** : Any Windows OS ≥ 98

### Features
The project has/will have the following features:

- [x] **Import Local Functions** : Import functions from other VBScript files into your project.
- [ ] **Import External Functions** : Import VBSCript directly from the Web (e.g. from a GitHub repository).
- [x] **Registry Operations** : Read a Value from Registry with a single function call.
- [ ] **Useful Functions** : A collection of useful functions. We will add more as we go!

### Future Improvements
I'm working on the following improvements:

- **Fully Type-Independent Contains** : Allow for _real_ type-independent contains.
- **RW Registry** : Write a Value to Registry with a single function call.
- **Auto-HTA** : Automatically create an HTA file for you with basic functionalities.
- **Security** : Add Security to the project, many functions and improvements are hugely insecure and can let insecure code run without scrutiny. E.g. `ImportFromWeb` could run malicious code on your machine if you don't know what you're doing.
- **Refactoring the code** : Generalize the code to make it more readable and maintainable.

## Installation and Usage
### Installation
To install the project, do the following:
1. Install the required libraries and tools
2. Fork, clone or make a local copy of the repository

### Usage
To use the project, do the following:
1. Read the project documentation, if any is provided;
2. Install the project following the installation instructions;
4. Start using the project!

## About the Project
The codebase for the project is hosted on [GitHub](github.com) in the source code folder.
If you want to contribute to the code, please read the [Contributing](CONTRIBUTING.md) section.

### Code Overview

The code for the project is written in the following languages:

- **[VBScript](https://en.wikipedia.org/wiki/VBScript)** : The codebase if fully written in VBScript.

If you want to see the codebase in a more detailed way, some snippets of code are provided below:

- **Updated Functions** : you can use the ternary operator to simplify if-then-else statements in VBScript with Iif or check whether a variable is contained in another variable without care with a single line of code, and many more functions!
  
```vbscript
Iif(Contains("VBScript","VB"), "Code runs", "Code doesn't run")
```

- **Handling Exceptions** : you can wrap your code in an exception handler to catch any exceptions that may occur within your code and create custom handling for them, by wrapping your subroutine in an `Exception_Wrapper` block with the following syntax:

```vbscript
Call Exception_Wrapper("My_Subroutine", "SubroutineHandler")
Call Exception_Wrapper("Another_Subroutine", "Another_Handler")

Sub HandleException(Subroutine, Exception)
    '    Wrap a Subroutine inside a given Exception.
    On Error Resume Next
    Dim Num
    Execute "Call " & Subroutine
    Num = Err.Number
    If CBool(Num) And VBool(Exception) Then
        On Error Goto 0
        Execute "Call " & Exception & "(" & Num &")"
        Err.Clear
    End If
    On Error Goto 0
End Sub
```

- **Import Dependencies** : You can import function and subroutines from other vbscript files with the following syntax:

```vbscript
Import "src/Functions"

MsgBox Contains("the brown fox jumps over the lazy dog", "fox")

Sub Import(Filename)
     '    Import Code from VBS File.
    Dim Lib, Code, FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Lib = FSO.OpenTextFile(Filename & ".vbs")
    Code = Lib.ReadAll
    Lib.Close
    ExecuteGlobal Code
    Set Lib = Nothing
    Set FSO = Nothing
End Sub
```

- **Random Numbers** : VBScript has a built-in random number generator `Rnd` which can be used to generate random numbers, however you are required to use the `Randomize` function to seed the random number generator. In order to avoid the random number generator being seeded with the same value every time the script is run, you can import `Functions` which has a built-in `Randomize`.

## Licensing and Credits
### Licensing
This project is licensed under the [GNU GPL v3](https://www.gnu.org/licenses/gpl-3.0.html) license. The code is open source and available for anyone to use and modify. Please see the [LICENSE](LICENSE.md) file for more information. 

### Credits
If you have any questions, please feel free to contact me, clicking on the link below.
<div align="center">
    <a href="mailto:fabioflorey@hackermail.com?subject=I%F0%9F%92%99VBSript"><kbd>
    <img src="https://avatars.githubusercontent.com/u/93403866?s=96&v=4" height="80" alt=""/></kbd><br>
    <sub><b>Fabio Craig Wimmer Florey</b></sub><br>
    <sup><b>Author</b></sup>
</div>


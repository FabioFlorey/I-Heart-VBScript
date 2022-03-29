Option Explicit

'                                 VBScript Main File Model
' Author:                            
'      Fabio Craig Wimmer Florey (fabioflorey@hackermail.com)
'
' Reviewed By:                                                  Last Review:
'      Fabio Craig Wimmer Florey (fabioflorey@hackermail.com)     2022-03-29
'
' Description:
'      Main Subroutine


Import "src/Functions"

Sub Main()
  ' Main Subroutine
End Sub

Sub Import(Filename)
  '    Import Code from VBS File, DO NOT DELETE
    Dim Lib, Code, FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Lib = FSO.OpenTextFile(Filename & ".vbs")
    Code = Lib.ReadAll
    Lib.Close
    ExecuteGlobal Code
    Set Lib = Nothing
    Set FSO = Nothing
End Sub

Call Main

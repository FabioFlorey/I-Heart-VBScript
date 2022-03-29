Option Explicit

'                               VBScript Functions Model
' Author:                            
'      Fabio Craig Wimmer Florey
'
' Reviewed By:                                                  Last Review:
'      Fabio Craig Wimmer Florey                                    2022-03-29
'
' Description:
'      Useful function for general purposes.

Randomize

Const StringTypeCode = 8
Const ArrayTypeCode  = 8191

' ----------------------------------------------------------------------------
'                              Generic Operations
' ----------------------------------------------------------------------------

Function IIf(Condition, True_Value, False_Value)
    ' Ternary Operator Function for VBScript.
    If Condition Then
        IIf = True_Value
   Else
        IIf = False_Value
    End If
End Function

Function VBool(Element)
    ' Convert to Boolean different DataTypes.
    Dim ElementType
    ElementType = VarType(Element)
    Select Case ElementType
        Case 0, 1
            Vbool = False
        Case 2, 3, 4, 5
            Vbool = CBool(Element)
        Case StringTypeCode
            VBool = CBool(Len(Element))
        Case Else
        VBool = IIf(ElementType>=ArrayTypeCode, _ 
            CBool(Ubound(Element)+1), _
            Element)
    End Select
End Function

Function Contains(Element, SubElement)
    ' Check if a given element of type String or 
    ' Array contains another given SubElement.
    On Error Resume Next
    Dim ElementType, ArrayCode
    ElementType = VarType(Element)
    Select Case ElementType
        Case StringTypeCode
            Contains = IIf(Cbool(InStr(Element,SubElement)), _
            True, IIf(Not Vbool(Element) And _
                        Not Vbool(SubElement), _ 
                        True, False))
        Case Else
            Contains = IIf(ElementType > ArrayTypeCode, _
                           ArrayContains(Element, SubElement), False)
    End Select
    If CBool(Err.Number) Then
        Contains = False
        Err.Clear
    End If
End Function


Function ArrayContains(ArrayElement, SubElement)
    ' Check if a given Array contains a given SubElement.
    Dim IsContained, SubArrayElement, ArrayCode
    For Each SubArrayElement in ArrayElement
        IsContained = False
        If(VarType(SubElement)) > ArrayTypeCode Then
            SubElement = Join(SubElement)
            SubArrayElement = IIf(VarType(SubArrayElement) > ArrayTypeCode, _
                              Join(SubArrayElement), SubArrayElement)
        End If              
        If SubArrayElement = SubElement Then
            IsContained = True
            Exit For
        End If
    Next
    ArrayContains = IsContained
End Function

' ----------------------------------------------------------------------------
'                               String Operations
' ----------------------------------------------------------------------------

Function StartsWith(Element, SubElement)
    ' Check if Element starts with SubElement
    StartsWith = IIf(InStr(Element, SubElement) = 1, True, False)
End Function


Function EndsWith(Element, SubElement)
    ' Check if Element ends with SubElement
    EndsWith = IIf(InStr(Element, SubElement) = _ 
                   Len(Element)-Len(SubElement) + 1 , True, False)
End Function


Function LeadingSubElement(Element, SubElement, Size):
    ' Add a leading SubElement to a given Element.
    LeadingSubElement = IIf(Len(Element) < Size, Replace(Space(Size - Len( _ 
                            IIf(Len(Element) < Size, Element,0)))," ", _ 
                            SubElement) & Element, Element)
End Function


Function TrailingSubElement(Element, SubElement, Size):
    ' Add a leading SubElement to a given Element.
    TrailingSubElement = IIf(Len(Element) < Size, _
                                Element & Replace(Space(Size - Len(IIf( _ 
                                Len(Element) < Size, Element,0))), " ", _
                                SubElement), Element)
End Function

Sub WebImport(URL)
    ' Import the VBS code at a given URL.
    ' and run it globally pushing the functions into the Main
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Request = createobject ("MSXML2.ServerXMLHTTP")
    FileName = "temp.vbs"
    
    Request.Open "GET", URL, false
    Request.Send
    TextResponse = Request.responSetext
    If FSO.FileExists(FileName) Then
    FSO.DeleteFile FileName, True
    End If
    Set WriteVBFile = FSO.CreateTextFile(FileName, True)
    WriteVBFile.Write TextResponse
    WriteVBFile.Close
    Set WriteVBFile = Nothing
    Set ReadVBFile = FSO.OpenTextFile(FileName, 1)
    Script = ReadVBFile.ReadAll
    ReadVBFile.Close
    Set Request = Nothing
    Set ReadVBFile = Nothing
    ExecuteGlobal Script    
    FSO.DeleteFile Filename, True
    Set FSO = Nothing
End Sub

    
' ----------------------------------------------------------------------------
'                               Registry Operations
' ----------------------------------------------------------------------------

Function ReadFromRegistry(KeyElement, DefaultValue)
    ' Read a given Value from a Registry Key.
    On Error Resume Next
    Dim Value
    Set Shell = CreateObject("WScript.Shell")
    Value = Shell.RegRead(KeyElement)
    ReadFromRegistry = IIf(Cbool(Err.Number), DefaultValue,Value)
    Set Shell = Nothing
End Function

' ----------------------------------------------------------------------------
'                            Basic Exception Handling
' ----------------------------------------------------------------------------


Sub CustomHandler(Number)
    ' Custom Exception Handler
      Select Case Number
      '  Do Something
        Case Else
            Err.Raise(Number)
      End Select
  End Sub
  
  
  Sub HandleException(Subroutine, Exception)
    ' Exception Handler
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
  
  

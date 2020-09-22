<div align="center">

## clsINI\.cls


</div>

### Description

Lately I've seen some posts for editing ini files that involve opening the ini file directly as a text file, looping line by line thru the file until locating the line desired and then altering that line. There is a much easier and more reliable way using the Windows API. This class module makes that easy. It also shows the proper way to handle errors that happen in a class module by raising custom error codes to be handled by the application that using using the class.
 
### More Info
 
See the comments in the code.

It assumes a basic familiarity with how to use class modules.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bryan Johns](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bryan-johns.md)
**Level**          |Intermediate
**User Rating**    |4.3 (39 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bryan-johns-clsini-cls__1-33363/archive/master.zip)





### Source Code

```
Option Explicit
' API Declarations
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
' Local variables to hold property values.
Private mstrINIPath As String
Private mstrFileName As String
Private mstrWindowsPath As String
Public Property Get WindowsPath() As String
 WindowsPath = mstrWindowsPath
End Property
Private Property Let WindowsPath(ByVal strWindowsPath As String)
 mstrWindowsPath = strWindowsPath
End Property
'***************************
'* Procedure: WindowsPathGet
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Uses an API call to set the read only WindowsPath property.
'****************************
Private Sub WindowsPathGet()
 Dim Y As String
 On Error GoTo Error
 mstrWindowsPath = Space(255)
 Y = GetWindowsDirectory(mstrWindowsPath, 255)
 mstrWindowsPath = Left$(mstrWindowsPath, Y)
 Exit Sub
Error:
 Err.Raise 10001, "clsINI.cls", "Unable to read the windows path."
End Sub
Public Property Get FileName() As String
 FileName = mstrFileName
End Property
Public Property Let FileName(ByVal strFileName As String)
 mstrFileName = strFileName
End Property
'***************************
'* Procedure: WriteINI
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Exposes the private WriteTo sub.
'****************************
Public Sub WriteINI(Section As String, Field As String, Value As String)
 WriteTo Section, Field, Value
End Sub
'***************************
'* Function : ReadINI
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Exposes the Private ReadFrom function.
'***************************
Public Function ReadINI(Section As String, Field As String) As String
 ReadINI = ReadFrom(Section, Field)
End Function
Public Property Get INIPath() As String
 INIPath = mstrINIPath
End Property
Public Property Let INIPath(ByVal strINIPath As String)
 mstrINIPath = strINIPath
End Property
'***************************
'* Function : ReadFrom
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Returns values read from the INI file.
'***************************
Private Function ReadFrom(lstrSection As String, lstrField As String) As String
 Dim varReturnedString As Integer
 Dim lstrResults As String
 lstrResults = Space(255)
 varReturnedString = GetPrivateProfileString&(lstrSection, lstrField, "", lstrResults, 255, mstrINIPath & "\" & mstrFileName)
 lstrResults = Left$(lstrResults, varReturnedString)
 If Len(lstrResults) < 1 Then
  Err.Raise 10000, "ReadFrom()", "Unable to read ini file entry."
  Exit Function
 End If
 ReadFrom = lstrResults
End Function
'***************************
'* Procedure: WriteTo
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Writes values to the INI file.
'****************************
Private Sub WriteTo(lstrSection As String, lstrField As String, lstrDefaultValue As String)
 Dim X As Boolean
 X = WritePrivateProfileString&(lstrSection, lstrField, lstrDefaultValue, mstrINIPath & "\" & mstrFileName)
 If X = False Then
  Err.Raise 10002, "WriteTo()", "There was a critical error writing to the" & mstrFileName & " file."
 End If
End Sub
Private Sub Class_Initialize()
 ' get the windows path and assign it to the INIPath property so that if the user of this
 ' class module doesn't supply a path it's defaulted to the windows path.
 WindowsPathGet
 mstrINIPath = mstrWindowsPath
End Sub
```


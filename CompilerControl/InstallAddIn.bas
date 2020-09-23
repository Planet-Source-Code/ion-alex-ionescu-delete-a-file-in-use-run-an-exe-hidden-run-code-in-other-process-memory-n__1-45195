Attribute VB_Name = "InstallAddIn"
Option Explicit

Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Declare Function GetLastError Lib "Kernel32" () As Long

Public Const ADDIN_NAME = "CompileController"

'To define the caption that appears in the Add-In Manager window go to the
'Object Browser (F2), select clsConnect, right click, select "Properties ..."
'VB's "Member Options" dialog should appear.  In the "Description" text box
'enter the caption you want to appear in the Add-Manager window.

Sub Main()
    Dim sExternalError As String
    If AddToINI(sExternalError) Then
        MsgBox "Add-In called """ & ADDIN_NAME & """ has been installed."
    Else
        MsgBox "Failed to install add-in: " & sExternalError
    End If
End Sub

'This procedure must be executed before VB's Add-In Manager will
'recognize the add-in as available.  Normally the procedure should be
'executed by the setup program.  During program development you will need
'to run it once in the immediate window to make the add-in available in
'your local environment.
Function AddToINI(sError As String) As Boolean
    Dim lngErrorCode As Long, lngErrorValue As Long
    On Error GoTo EH
    lngErrorValue = WritePrivateProfileString("Add-Ins32", ADDIN_NAME & ".clsConnect", "0", "vbaddin.ini")
    If lngErrorValue = 0 Then
        lngErrorCode = GetLastError
        sError = "WritePrivateProfileString generated error code: " & lngErrorCode
    Else
        AddToINI = True
    End If
    Exit Function
EH:
    sError = "Unexpected error writing private profile string to vbaddin.ini: " & Err.Description
End Function


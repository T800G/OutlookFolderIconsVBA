Attribute VB_Name = "OpenFileDlg"
Option Explicit

' >>>>>>>> test x64!!! /StrConv(**, vbUnicode)?
#If Win64 Then
Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
#Else
Private Declare Function GetOpenFileNameWin32 Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
#End If

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'http://msdn.microsoft.com/en-us/library/windows/desktop/ms646839%28v=vs.85%29.aspx
Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Function GetOpenFileName(Optional ApphInstance As Long, Optional hwndOwner As Long, _
                                Optional sFilter As String, Optional dlgTitle As String, _
                                Optional strInitialDir As String) As String
On Error GoTo Err
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
Redo:
    lReturn = 0
    GetOpenFileName = vbNullString
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = hwndOwner
    OpenFile.hInstance = ApphInstance
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = strInitialDir
    OpenFile.lpstrTitle = dlgTitle
    OpenFile.Flags = OFN_FILEMUSTEXIST
    lReturn = GetOpenFileNameWin32(OpenFile)
    If lReturn <> 0 Then
       GetOpenFileName = Trim(StrConv(OpenFile.lpstrFile, vbUnicode))
       If InStr(1, GetOpenFileName, "?", vbTextCompare) <> 0 Then
            GetOpenFileName = vbNullString
            GoTo Redo
       End If
    End If
    Exit Function
Err:
    GetOpenFileName = vbNullString
    Debug.Print "OpenFileDlg::GetOpenFileName : error " & Err.Number & ": " & Err.Description
End Function


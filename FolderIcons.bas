Attribute VB_Name = "FolderIcons"
Option Explicit

Private g_fiEvents As FolderIconsEvents

'MSXML3 reference must be added to the project!
Private g_xmlDoc As MSXML2.DOMDocument

'settings save location
Private Const xmlPath = "D:\T800\Programming\VBA\Outlook Folder Icons\OutlookFolderIcons.xml"

Private Const strDefaultXML = "<?xml version='1.0'?><OUTLOOK></OUTLOOK>"

Private Const olFolderMIN = 3
Private Const olFolderMAX = 30
Public Enum OlDefaultFolders2010
    olFolderDeletedItems = 3 'The Deleted Items folder.
    olFolderOutbox = 4 'The Outbox folder.
    olFolderSentMail = 5 'The Sent Mail folder.
    olFolderInbox = 6 'The Inbox folder.
    olFolderCalendar = 9 'The Calendar folder.
    olFolderContacts = 10 'The Contacts folder.
    olFolderJournal = 11 'The Journal folder.
    olFolderNotes = 12 'The Notes folder.
    olFolderTasks = 13 'The Tasks folder.
    olFolderDrafts = 16 'The Drafts folder.
    olPublicFoldersAllPublicFolders = 18 'The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
    olFolderConflicts = 19 'The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    olFolderSyncIssues = 20 'The Sync Issues folder. Only available for an Exchange account.
    olFolderLocalFailures = 21 'The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    olFolderServerFailures = 22 'The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    olFolderJunk = 23 'The Junk E-Mail folder.
    olFolderRssFeeds = 25 'The RSS Feeds folder.
    olFolderToDo = 28 'The To Do folder.
    olFolderManagedEmail = 29 'The top-level folder in the Managed Folders group. Only available for an Exchange account.
    'Outlook 2013+
    olFolderSuggestedContacts = 30 'The Suggested Contacts folder.
End Enum


'TODO: test md5 on x64

'//////////////////////////////////////////////////////////////////////
' storeID string is too long for XML, use string hash instead
Private Type MD5_CTX
  i(1 To 2) As Long
  buf(1 To 4) As Long
  inp(1 To 64) As Byte
  digest(1 To 16) As Byte
End Type
#If Win64 Then
Private Declare PtrSafe Sub MD5Init Lib "cryptdll" (Context As MD5_CTX)
Private Declare PtrSafe Sub MD5Update Lib "cryptdll" (Context As MD5_CTX, ByVal strInput As String, ByVal lLen As Long)
Private Declare PtrSafe Sub MD5Final Lib "cryptdll" (Context As MD5_CTX)
#Else
Private Declare Sub MD5Init Lib "cryptdll" (Context As MD5_CTX)
Private Declare Sub MD5Update Lib "cryptdll" (Context As MD5_CTX, ByVal strInput As String, ByVal lLen As Long)
Private Declare Sub MD5Final Lib "cryptdll" (Context As MD5_CTX)
#End If
'//////////////////////////////////////////////////////////////////////
#If Win64 Then
Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongLong
Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long
Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
#Else
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
#End If
'//////////////////////////////////////////////////////////////////////
Private Enum SystemMetrics
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXSMICON = 49
    SM_CYSMICON = 50
End Enum
#If Win64 Then
Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
'//////////////////////////////////////////////////////////////////////
Public Sub SetFolderIcon()
    If ActiveExplorer Is Nothing Then Exit Sub
    If IsDefaultFolder(ActiveExplorer.CurrentFolder) Then
        MsgBox "Custom icon cannot be set for a default or special folder", vbExclamation, "Set Folder Icon"
        Exit Sub
    End If
    If g_xmlDoc Is Nothing Then LoadSettings
    Dim path As String
    'get dialog parent window as early as possible (user might switch to other process/window)
    Dim hWnd As Long
    hWnd = GetForegroundWindow()
    Dim threadID As Long
    threadID = GetWindowThreadProcessId(hWnd, vbNull)
    If threadID <> GetCurrentThreadId() Then hWnd = 0
    Dim fileFilter As String
    fileFilter = "Icons (*.ico)" & Chr(0) & "*.ico" & Chr(0) & _
                "Pictures (*.bmp, *.dib, *.jpg, *.jpeg, *.jpe, *.jfif, *.gif)" & Chr(0) & _
                "*.bmp;*.dib;*.jpg;*.jpeg;*.jpe;*.jfif;*.gif" & Chr(0) '& _
                '"All (*.*)" & Chr(0) & "*.*" & Chr(0)
    path = OpenFileDlg.GetOpenFileName(0, hWnd, fileFilter, "Select icon")
    If path <> "" Then
        path = Left(path, InStr(1, path, vbNullChar) - 1) 'trim trailing nulls
        'check if bad filetype is selected anyway
        If InStr(1, ".ico.bmp.dib.jpg.jpeg.jpe.jfif.gif", LCase(Mid(path, InStrRev(path, "."), 5))) = 0 Then
            MsgBox "Unsupported file type", vbExclamation, "Set Folder Icon"
            Exit Sub
        End If
        SetFolderIconImpl ActiveExplorer.CurrentFolder, path
        g_xmlDoc.Save xmlPath
    End If
End Sub
'//////////////////////////////////////////////////////////////////////
Public Sub RemoveFolderIcon()
    If ActiveExplorer Is Nothing Then Exit Sub
    If ActiveExplorer.CurrentFolder.GetCustomIcon Is Nothing Then Exit Sub
    If g_xmlDoc Is Nothing Then LoadSettings
    If MsgBox("Remove icon from" & vbCrLf & ActiveExplorer.CurrentFolder.folderPath, _
                vbOKCancel Or vbDefaultButton2 Or vbQuestion, "Remove Folder Icon") = vbCancel Then Exit Sub
    DeleteFolderIconSetting ActiveExplorer.CurrentFolder
    g_xmlDoc.Save xmlPath
    ActiveExplorer.CurrentFolder.SetCustomIcon Nothing 'undocumented way to get default icon back!
End Sub
'//////////////////////////////////////////////////////////////////////
Public Function Initialize() As Variant
    If g_fiEvents Is Nothing Then
        Set g_fiEvents = New FolderIconsEvents
    End If
    If ActiveExplorer Is Nothing Then Exit Function
    LoadSettings
    Dim oStore As Outlook.Store
    Dim oRoot As Outlook.folder
    On Error Resume Next
    For Each oStore In Application.Session.Stores
        Set oRoot = oStore.GetRootFolder
        EnumerateFolders oRoot
    Next
    g_xmlDoc.Save xmlPath
End Function
'//////////////////////////////////////////////////////////////////////
Private Sub EnumerateFolders(ByRef oFolder As Outlook.folder)
    On Error Resume Next
    Dim c As Long
    c = SetFolderIconImpl(oFolder)
    If c <> 0 Then Exit Sub
    Dim fldrs As Outlook.Folders
    Set fldrs = oFolder.Folders
    c = fldrs.count
    If c Then
        Dim fldr As Outlook.folder
        Dim icopath As String
        For Each fldr In fldrs
            EnumerateFolders fldr 'recursion!
        Next
    End If
End Sub
'called by enumerator (no path) and by command macro (filepath)
Private Function SetFolderIconImpl(ByRef fldr As folder, Optional ByVal filePath As String = "") As Long
    On Error GoTo ErrHandler
    Dim iconPath As String
    If filePath = "" Then
        iconPath = GetFolderIconSetting(fldr)
        If iconPath = "" Then
            Exit Function
        End If
    Else
        iconPath = filePath
    End If
    If Not FileExists(iconPath) Then
        Debug.Print "icon (" & iconPath & ") not found"
        DeleteFolderIconSetting fldr
        Exit Function
    End If
    Dim hresult As Long, w As Long, h As Long
    w = GetSystemMetrics(SystemMetrics.SM_CXSMICON)
    h = GetSystemMetrics(SystemMetrics.SM_CYSMICON)
    Dim pDisp As IPictureDisp
    Set pDisp = LoadPicture(iconPath, w, h, LoadPictureConstants.Default) 'throws exception on errors!
    If pDisp Is Nothing Then
        Debug.Print "IPictureDisp is nothing " & Err.Number & ": " & Err.Description
        SetFolderIconImpl = Err.Number
        Exit Function
    End If
    fldr.SetCustomIcon pDisp
    Set pDisp = Nothing
    'no errors so far
    SaveFolderIconSetting fldr, iconPath
    On Error GoTo 0
    Exit Function
ErrHandler:
    SetFolderIconImpl = Err.Number
    If Not pDisp Is Nothing Then Set pDisp = Nothing
    Debug.Print "FolderIcons::SetFolderIconImpl:  Line " & Erl & " error " & Err.Number & ": " & Err.Description
    MsgBox Err.Description, vbExclamation, "Set Folder Icon"
End Function
Private Sub LoadSettings()
    'this sub should be called only once
    On Error GoTo ErrHandler
    If g_xmlDoc Is Nothing Then
        Set g_xmlDoc = New MSXML2.DOMDocument
    End If
    g_xmlDoc.async = False
    If FileExists(xmlPath) Then
        g_xmlDoc.Load xmlPath
    Else
        g_xmlDoc.LoadXML strDefaultXML
    End If
    If g_xmlDoc.parseError.ErrorCode <> 0 Then
        Debug.Print "g_xmlDoc.parseError=" & g_xmlDoc.parseError.reason
        Exit Sub
    End If
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "FolderIcons::LoadSettings Line " & Erl & " error " & Err.Number & ": " & Err.Description
End Sub
Private Function GetFolderIconSetting(ByRef fldr As folder) As String
    On Error GoTo ErrHandler
    If g_xmlDoc Is Nothing Then Exit Function
    Dim pNode As MSXML2.IXMLDOMNode
    Set pNode = g_xmlDoc.SelectSingleNode("//OUTLOOK/STORE[@storeidMD5='" & CalcMD5(fldr.Store.storeID) & "']/FOLDER[@entryID='" & fldr.entryID & "']")
    If pNode Is Nothing Then Exit Function
    GetFolderIconSetting = pNode.Text
    On Error GoTo 0
    Exit Function
ErrHandler:
    Debug.Print "FolderIcons::GetFolderIconSetting Line " & Erl & " error " & Err.Number & ": " & Err.Description
End Function
Private Sub SaveFolderIconSetting(ByRef fldr As folder, ByVal iconPath As String)
    On Error GoTo ErrHandler
    If g_xmlDoc Is Nothing Then Exit Sub
    Dim pRoot As MSXML2.IXMLDOMNode, pStore As MSXML2.IXMLDOMNode, pFolder As MSXML2.IXMLDOMNode
    Set pRoot = g_xmlDoc.SelectSingleNode("//OUTLOOK")
    If pRoot Is Nothing Then
        Debug.Print "critical error: no OUTLOOK node" 'not my xml?
        Exit Sub
    End If
    Dim md5 As String
    md5 = CalcMD5(fldr.Store.storeID)
    Set pStore = pRoot.SelectSingleNode("//OUTLOOK/STORE[@storeidMD5='" & md5 & "']")
    If pStore Is Nothing Then
        Set pStore = AddXMLNode(pRoot, "STORE", vbNullString, vbNullString, "storeidMD5", md5)
        If pStore Is Nothing Then
            Debug.Print "AddXMLNode(STORE) failed"
            Exit Sub
        End If
    End If
    Set pFolder = pStore.SelectSingleNode("//STORE/FOLDER[@entryID='" & fldr.entryID & "']")
    If pFolder Is Nothing Then
        Set pFolder = AddXMLNode(pStore, "FOLDER", vbNullString, vbNullString, "entryID", fldr.entryID)
        If pFolder Is Nothing Then
            Debug.Print "AddXMLNode(FOLDER) failed"
            Exit Sub
        End If
    End If
    pFolder.Text = iconPath
    'Debug.Print g_xmlDoc.XML
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "FolderIcons::SaveFolderIconSetting Line " & Erl & " error " & Err.Number & ": " & Err.Description
End Sub
Private Sub DeleteFolderIconSetting(ByRef fldr As folder)
    On Error GoTo ErrHandler
    If g_xmlDoc Is Nothing Then Exit Sub
    Dim pNode As MSXML2.IXMLDOMNode
    Set pNode = g_xmlDoc.SelectSingleNode("//OUTLOOK/STORE[@storeidMD5='" & CalcMD5(fldr.Store.storeID) & "']/FOLDER[@entryID='" & fldr.entryID & "']")
    If pNode Is Nothing Then Exit Sub
    Dim parent As MSXML2.IXMLDOMNode
    Set parent = pNode.parentNode
    parent.RemoveChild pNode
    If parent.ChildNodes.Length = 0 Then parent.parentNode.RemoveChild parent 'no child nodes left, remove self (store level)
    'Debug.Print g_xmlDoc.XML
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "FolderIcons::DeleteFolderIconSetting Line " & Erl & " error " & Err.Number & ": " & Err.Description
End Sub
Private Function FileExists(ByVal fName As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function
Private Function IsDefaultFolder(ByRef fldr As folder) As Boolean
    On Error Resume Next
    Dim oFldr As folder, nsp As namespace
    Set nsp = GetNamespace("MAPI")
    If nsp Is Nothing Then Exit Function
    Dim i As Integer
    For i = olFolderMIN To olFolderMAX
        Set oFldr = nsp.GetDefaultFolder(i)
        If Not oFldr Is Nothing Then
            If (oFldr.entryID = fldr.entryID) And (oFldr.storeID = fldr.storeID) Then
                IsDefaultFolder = True
                Set oFldr = Nothing
                Exit Function
            End If
            Set oFldr = Nothing
        End If
    Next i
End Function
Private Function AddXMLNode(ByRef parentNode As MSXML2.IXMLDOMNode, ByVal nodeName As String, ByVal nodeText As String, ByVal namespaceURI As String, _
                            ByVal attributeName As String, ByVal attributeValue As String) As MSXML2.IXMLDOMNode
    On Error GoTo ErrHandler
    'msoCustomXMLNodeElement=1,msoCustomXMLNodeAttribute=2
    Set AddXMLNode = parentNode.OwnerDocument.createNode(1, nodeName, namespaceURI)
    Dim attr As MSXML2.IXMLDOMAttribute
    Set attr = parentNode.OwnerDocument.createAttribute(attributeName)
    attr.Value = attributeValue
    AddXMLNode.Attributes.setNamedItem attr
    AddXMLNode.Text = nodeText
    parentNode.appendChild AddXMLNode
    On Error GoTo 0
    Exit Function
ErrHandler:
    Set AddXMLNode = Nothing
    Debug.Print "FolderIcons::AddXMLNode Line " & Erl & " error " & Err.Number & ": " & Err.Description
End Function
Private Function CalcMD5(strBuffer As String) As String
    Dim ctx As MD5_CTX
    MD5Init ctx
    MD5Update ctx, strBuffer, Len(strBuffer)
    MD5Final ctx
    Dim result As String
    result = StrConv(ctx.digest, vbUnicode)
    Dim lp As Long
    For lp = 1 To Len(result)
            CalcMD5 = CalcMD5 & Right("00" & Hex(Asc(Mid(result, lp, 1))), 2)
    Next
End Function

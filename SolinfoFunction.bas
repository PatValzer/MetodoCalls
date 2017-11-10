Attribute VB_Name = "SolinfoFunction"
Option Explicit

Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As Long)
'Type SYSTEMTIME
'    wYear As Integer
'    wMonth As Integer
'    wDayOfWeek As Integer
'    wDay As Integer
'    wHour As Integer
'    wMinute As Integer
'    wSecond As Integer
'    wMilliseconds As Integer
'End Type
'
'codice sul tuo form
'codice:
'Dim MyTime As SYSTEMTIME
'
'Private Sub Command1_Click()
'End
'End Sub
'
'Private Sub Form_Initialize()
'' valore dell'interval va da 1 a 100
'Timer1.Interval = 10
'End Sub
'
'Private Sub Timer1_Timer()
'GetLocalTime MyTime
'Text$ = MyTime.wHour & ":" & MyTime.wMinute & ":" & _
'        MyTime.wSecond & ":" & MyTime.wMilliseconds
'Text1.Text = Text$
'End Sub



Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function UuidCreate& Lib "rpcrt4" (lpGUID As GUID)
Private Declare Function UuidToString& Lib "rpcrt4" Alias "UuidToStringA" (lpGUID As GUID, lpGUIDString&)
Private Declare Function RpcStringFree& Lib "rpcrt4" Alias "RpcStringFreeA" (lpGUIDString&)

Private Declare Function lstrlen& Lib "kernel32" Alias "lstrlenA" (ByVal lpString&)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Const BIF_RETURNONLYFSDIRS = &H1
'Private shlShell As shell32.Shell
'Private shlFolder As shell32.Folder


Private Const GW_HWNDPREV = 3

Public Const SWP_FRAMECHANGED = &H20
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Private Const RPC_S_OK As Long = 0&
Private Const RPC_S_UUID_LOCAL_ONLY As Long = 1824&
Private Const RPC_S_UUID_NO_ADDRESS As Long = 1739&

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Private Type BrowseInfo
 hWndOwner As Long
 pIDLRoot As Long
 pszDisplayName As Long
 lpszTitle As Long
 ulFlags As Long
 lpfnCallback As Long
 lParam As Long
 iImage As Long
 End Type
 
 Const BIF_RETURNONLYFSDIRS = 1
 Const MAX_PATH = 260
 Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
 Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
 Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
 Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
 


Public Sub FormOnTop(frm As Form, VF As Boolean)
    ' Questa procedura permette di impostare sempre in primo piano un form rispetto alle altre finestre
    ' frm ---> form su cui agisce la procedura
    ' vf  ---> se true il form è visualizzato in primo piano
    
'    Select Case VF
'        Case True
'          SetWindowPos FRM.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
'        Case False
'          SetWindowPos FRM.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE
'    End Select
End Sub

Public Function CreateGUID() As String

  Dim nGUID As GUID, lpStrGUID&, nStrLen&, sGUID$
  Dim sBuffer() As Byte, bNoError As Boolean
  
  ' create the GUID
  If UuidCreate(nGUID) <> RPC_S_UUID_NO_ADDRESS Then

    ' if the GUID was created, convert it to a string since we
    ' can't do much with the struct that holds it now
    If UuidToString(nGUID, lpStrGUID) = RPC_S_OK Then
    
      ' get the length of the string that was returned
      nStrLen = lstrlen(lpStrGUID)
  
      ' dimention our byte array to the proper size
      ReDim sBuffer(nStrLen - 1) As Byte
      
      ' copy the string to the byte array
      CopyMemory sBuffer(0), ByVal lpStrGUID, nStrLen
    
      ' release the memory for the string that Windows returned
      Call RpcStringFree(lpStrGUID)
  
      ' convert the string to unicode so that we can display it.
      ' if you don't need to display the string, you may want to
      ' remove this call since converting to unicode doubles the
      ' amount of the storage space required.
      sGUID = StrConv(sBuffer, vbUnicode)

      ' show it in the label
      CreateGUID = UCase$(sGUID)
      
      ' exit before the error message
      Exit Function
    End If
  End If
  
  ' if we get here an error occurred
  CreateGUID = "Error, unable to create GUID"
  
End Function


Public Sub Attendi(ByVal Sec As Double)
Dim Inizio As Single

Inizio = VBA.Timer
Do While (VBA.Timer - Inizio < Sec)
    DoEvents
Loop
End Sub

' The name of the interactive user
Public Function UserName() As String
On Error GoTo Err_UserName
    Dim buffer As String * 512, length As Long
    length = Len(buffer)
    If GetUserName(buffer, length) Then
        ' returns non-zero if successful, and modifies the length argument
        UserName = Left$(buffer, length - 1)
    End If
Err_UserName:
    If Err <> 0 Then
        UserName = ""
    End If
End Function



Public Function GetValueINI(ByVal strSection As String, ByVal strKey As String, ByVal varDefault As Variant, Optional ByVal bWriteIfAbsent As Boolean = True) As String
On Error Resume Next
Dim fs As New filesystemobject
    Dim lLen As Long
    Dim strVal As String * 256
    Dim strINI As String
    strINI = fs.BuildPath(App.Path, App.EXEName & ".ini")
    'MsgBox strINI

    lLen = GetPrivateProfileString(strSection, strKey, "", strVal, Len(strVal), strINI)
    If lLen <> 0 Then
        GetValueINI = (RTrim$(Left$(strVal, lLen)))
    Else
        GetValueINI = varDefault
        If bWriteIfAbsent Then WritePrivateProfileString strSection, strKey, CStr(varDefault), strINI
    End If

Set fs = Nothing
End Function

Public Function SetValueINI(ByVal strSection As String, ByVal strKey As String, ByVal varValue As Variant) As String
Dim fs As New filesystemobject
    Dim lLen As Long
    Dim strVal As String * 256
    Dim strINI As String
    strINI = fs.BuildPath(App.Path, App.EXEName & ".ini")
    WritePrivateProfileString strSection, strKey, CStr(varValue), strINI
End Function


Public Function GetValueINI_File(ByVal strSection As String, ByVal strKey As String, ByVal varDefault As Variant, Optional ByVal bWriteIfAbsent As Boolean = True, Optional ByVal filename As String = "") As String
On Error Resume Next
Dim fs As New filesystemobject
    Dim lLen As Long
    Dim strVal As String * 256
    Dim strINI As String
    
    
    If fs.FileExists(filename) Then
        strINI = filename
    Else
        strINI = fs.BuildPath(App.Path, App.EXEName & ".ini")
    End If

    lLen = GetPrivateProfileString(strSection, strKey, "", strVal, Len(strVal), strINI)
    If lLen <> 0 Then
        GetValueINI_File = (RTrim$(Left$(strVal, lLen)))
    Else
        GetValueINI_File = varDefault
        If bWriteIfAbsent Then WritePrivateProfileString strSection, strKey, CStr(varDefault), strINI
    End If

Set fs = Nothing
End Function

Public Function SetValueINI_File(ByVal strSection As String, ByVal strKey As String, ByVal varValue As Variant, Optional ByVal filename As String = "") As String
Dim fs As New filesystemobject
    Dim lLen As Long
    Dim strVal As String * 256
    Dim strINI As String
    If fs.FileExists(filename) Then
        strINI = filename
    Else
        strINI = fs.BuildPath(App.Path, App.EXEName & ".ini")
    End If
    WritePrivateProfileString strSection, strKey, CStr(varValue), strINI
End Function


Public Sub AggiornaLog(strTesto As String)
On Error Resume Next
Dim fs As New filesystemobject
Dim txt As TextStream
Dim strPercorso As String
    If GetValueINI("DEBUG", "ATTIVALOG", 1, True) = 1 Then
        strPercorso = fs.BuildPath(App.Path, "LOG")
        If Not fs.FolderExists(strPercorso) Then fs.CreateFolder strPercorso
        Set txt = fs.OpenTextFile(fs.BuildPath(strPercorso, Format(Now, "YYMMDD") & "_" & App.EXEName & "_" & UserName & ".txt"), ForAppending, True)
        txt.WriteLine strTesto
        txt.Close
    End If
Set fs = Nothing
End Sub

Public Sub VisualizzaLog(Optional strFileName As String)
Dim fs As New filesystemobject
Dim txt As TextStream
Dim strPercorso As String
    If Len(strFileName) = 0 Then
        strPercorso = fs.BuildPath(App.Path, "LOG")
        strPercorso = fs.BuildPath(strPercorso, Format(Now, "YYMMDD") & "_" & App.EXEName & "_" & UserName & ".txt")
    Else
        strPercorso = strFileName
    End If
    Shell "notepad.exe " & strPercorso, vbNormalFocus
    Set fs = Nothing
End Sub

Public Sub ApriFile(strFile As String)
    Shell "notepad.exe " & strFile, vbNormalFocus
End Sub

'
Public Sub VisualizzaTesto(strTeso As String)
    'MsgBox strTeso
End Sub

Public Function LeggiLog(strPercorso As String) As String
On Error Resume Next
Dim fs As New filesystemobject
Dim txt As TextStream
    LeggiLog = ""
    Set txt = fs.OpenTextFile(strPercorso, ForReading, False)
    LeggiLog = txt.ReadAll
    Set fs = Nothing
End Function


'Dim fs As New FileSystemObject
'Dim txt As TextStream
'Dim strPercorso As String
'    strPercorso = MXNU.GetTempFile
'    Set txt = fs.OpenTextFile(strPercorso, ForWriting, True)
'    txt.Write strTeso
'    Shell "notepad.exe " & strPercorso, vbNormalFocus
'    Set fs = Nothing
'End Sub

Public Function DialogBoxOpen(cmddlg As CommonDialog, Optional strTitle As String = "", Optional strFilter As String = "Tutti i file (*.*) |*.*|", Optional strInitDir As String = "", Optional strFileName As String = "", Optional blnCancelError As Boolean = True, Optional blnRelative As Boolean = False) As String
On Error GoTo Err_OpenDialogBox
Dim fs As New filesystemobject
    cmddlg.DialogTitle = strTitle
    cmddlg.Filter = strFilter
    cmddlg.CancelError = blnCancelError
    cmddlg.InitDir = strInitDir
    cmddlg.filename = strFileName
    cmddlg.ShowOpen
Err_OpenDialogBox:
    Select Case Err.Number
        Case Is = 0
        'L'utente ha selezionato un file
            If blnRelative <> False Then
                If "" <> fs.GetParentFolderName(cmddlg.filename) Then
                    DialogBoxOpen = cmddlg.filename
                Else
                    DialogBoxOpen = fs.GetFileName(cmddlg.filename)
                End If
            Else
                DialogBoxOpen = cmddlg.filename
            End If
        Case Is = 32755
        'L'utente non ha selezionato nessun file
        'Ritorno per default il valore originale
            If blnRelative <> False Then
                If "" <> fs.GetParentFolderName(fs.BuildPath(strInitDir, strFileName)) Then
                    DialogBoxOpen = fs.BuildPath(strInitDir, strFileName)
                Else
                    DialogBoxOpen = fs.GetFileName(fs.BuildPath(strInitDir, strFileName))
                End If
            Else
                DialogBoxOpen = fs.BuildPath(strInitDir, strFileName)
            End If
        Case Else
            If Err.Number <> 0 Then
                MsgBox LoadResString(502), vbCritical
                If blnRelative <> False Then
                    If "" <> fs.GetParentFolderName(fs.BuildPath(strInitDir, strFileName)) Then
                        DialogBoxOpen = fs.BuildPath(strInitDir, strFileName)
                    Else
                        DialogBoxOpen = fs.GetFileName(fs.BuildPath(strInitDir, strFileName))
                    End If
                Else
                    DialogBoxOpen = fs.BuildPath(strInitDir, strFileName)
                End If
            End If
    End Select
    Set fs = Nothing
End Function

Public Function DialogBoxSave(cmddlg As CommonDialog, Optional strTitle As String = "", Optional strFilter As String = "Tutti i file (*.*) |*.*|", Optional strInitDir As String = "", Optional strFileName As String = "FileSenzaNome", Optional blnCancelError As Boolean = True) As String
On Error GoTo Err_SalvaDialogBox
    cmddlg.DialogTitle = strTitle
    cmddlg.Filter = strFilter
    cmddlg.CancelError = blnCancelError
    cmddlg.InitDir = strInitDir
    cmddlg.filename = strFileName
    cmddlg.ShowSave
Err_SalvaDialogBox:
    Select Case Err.Number
        Case Is = 0
        'L'utente ha selezionato un file
            DialogBoxSave = cmddlg.filename
        Case Is = 32755
        'L'utente non ha selezionato nessun file
        'Ritorno per default il valore originale
            DialogBoxSave = ""
        Case Else
            If Err.Number <> 0 Then
                MsgBox LoadResString(502), vbCritical
                DialogBoxSave = ""
            End If
    End Select
End Function

Public Sub Clessidra()
    Screen.MousePointer = 11
End Sub

Public Sub Freccia()
    Screen.MousePointer = 0
End Sub

Sub ActivatePrevInstance()
    Dim OldTitle As String
    Dim PrevHndl As Long
    Dim result As Long
    'Save the title of the application.
    OldTitle = App.Title
    'Rename the title of this application so FindWindow
    'will not find this application instance.
    App.Title = "unwanted instance"
    'Attempt to get window handle using VB4 class name.
    PrevHndl = FindWindow("ThunderRTMain", OldTitle)
    'Check for no success.
    If PrevHndl = 0 Then
        'Attempt to get window handle using VB5 class name.
        PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
    End If
    'Check if found
    If PrevHndl = 0 Then
        'Attempt to get window handle using VB6 class name
        PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
    End If
    'Check if found
    If PrevHndl = 0 Then
        'No previous instance found.
        Exit Sub
    End If
    'Get handle to previous window.
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    'Restore the program.
    result = OpenIcon(PrevHndl)
    'Activate the application.
    result = SetForegroundWindow(PrevHndl)
    'End the application.
    'End
End Sub

'  Public Function SelezionaCartella(hwnd As Long) As String
'  Dim fs As New FileSystemObject
'      If shlShell Is Nothing Then
'          Set shlShell = New shell32.Shell
'      End If
'      Set shlFolder = shlShell.BrowseForFolder(hwnd, "Select a Directory", BIF_RETURNONLYFSDIRS)
'      If Not shlFolder Is Nothing Then
'          SelezionaCartella = shlFolder.Self.Path
'      Else
'          SelezionaCartella = ""
'      End If
'  Set fs = Nothing
'  End Function


 
 Public Function SelezionaCartella() As String
 Dim iNull As Integer, lpIDList As Long, lResult As Long
 Dim strPath As String, udtBI As BrowseInfo
 
With udtBI
 .hWndOwner = 0&
 .lpszTitle = lstrcat("Seleziona cartella", "")
 'Return only if the user selected a directory
 .ulFlags = BIF_RETURNONLYFSDIRS
 End With
 
lpIDList = SHBrowseForFolder(udtBI)
 If lpIDList Then
 strPath = String$(MAX_PATH, 0)
 SHGetPathFromIDList lpIDList, strPath
 'free the block of memory
 CoTaskMemFree lpIDList
 iNull = InStr(strPath, vbNullChar)
 If iNull Then
 strPath = Left$(strPath, iNull - 1)
 End If
 End If
 
    SelezionaCartella = strPath
 End Function
 


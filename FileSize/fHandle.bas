Attribute VB_Name = "fHandle"
Option Explicit
'**************************************************************************************
'API function declarations for the shell execute function to open a file with its associated app
'**************************************************************************************
Private Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

'***App Window Constants***
Public Const WIN_NORMAL = 1         'Open Normal
Public Const WIN_MAX = 3            'Open Maximized
Public Const WIN_MIN = 2            'Open Minimized
'AGP - I added the following two constants
Public Const WIN_EXPLORE = "Explore" 'Open with Explorer
Public Const WIN_OPEN = "Open"       'Open with a Folder View

'***Error Codes***
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_ACCESSDENIED = 5&
Private Const ERROR_OUT_OF_MEM2 = 8&
Private Const ERROR_BAD_FORMAT = 11&
Private Const ERROR_SHARE = 26&
Private Const ERROR_ASSOCINCOMPLETE = 27&
Private Const ERROR_DDETIMEOUT = 28&
Private Const ERROR_DDEFAIL = 29&
Private Const ERROR_DDEBUSY = 30&
Private Const ERROR_NO_ASSOC = 31&
Private Const SHELL_SUCCESS = 32&
Private Const ERROR_DLLNOTFOUND = 32&

Public fRet As Long

'Open a folder:
'Print fHandleFile("f:\power\", WIN_NORMAL, WIN_EXPLORE)
'Call Email app:
'Print fHandleFile("mailto:dash10@hotmail.com", WIN_NORMAL)
'Open URL:
'Print fHandleFile("http://www.mvps.org/access/", WIN_NORMAL)
'Handle Unknown extensions (call Open With Dialog):
'
'Print fHandleFile("F:\Documents and Settings\Administrator\My Documents\My VB\Ruby\ruby.don", WIN_NORMAL)
'Start Access instance:
'Print fHandleFile("F:\Documents and Settings\Administrato
Public Function fHandleFile(stFile As String, lShowHow As Long, Optional OpenType As String)
    Dim lRet As Long, varTaskID As Variant
    Dim stRet As String
    
    If OpenType <> WIN_EXPLORE And OpenType <> WIN_OPEN Then OpenType = vbNullString
    
    'First try ShellExecute
    lRet = apiShellExecute(&O0, OpenType, stFile, vbNullString, vbNullString, lShowHow)
    
    If lRet > SHELL_SUCCESS Then
        stRet = vbNullString
        lRet = -1
    Else
        Select Case lRet
            Case ERROR_NO_ASSOC:
                'Try the OpenWith dialog
                varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & stFile, WIN_NORMAL)
                lRet = (varTaskID <> 0)
            Case ERROR_OUT_OF_MEM, ERROR_OUT_OF_MEM2:
                stRet = "Out of Memory/Resources"
            Case ERROR_FILE_NOT_FOUND:
                stRet = "File not found"
            Case ERROR_PATH_NOT_FOUND:
                stRet = "Path not found"
            Case ERROR_BAD_FORMAT:
                stRet = "Bad File Format. Invalid EXE file or error in EXE image"
            'new error messages
            Case ERROR_ACCESSDENIED
                stRet = "Access denied"
            Case ERROR_DLLNOTFOUND
                stRet = "DLL not found"
            Case ERROR_SHARE
                stRet = "A sharing violation occurred"
            Case ERROR_ASSOCINCOMPLETE
                stRet = "Incomplete or invalid file association"
            Case ERROR_DDETIMEOUT
                stRet = "DDE Time out"
            Case ERROR_DDEFAIL
                stRet = "DDE transaction failed"
            Case ERROR_DDEBUSY
                stRet = "DDE busy"
            Case Else
                stRet = "Unknown error"
        End Select
        stRet = "Error number: " & lRet & vbCrLf & _
                "Error description: " & stRet
        If lRet <> -1 Then MsgBox stFile & vbCrLf & stRet, vbCritical, "Shell execute failed..."
    End If
    fHandleFile = lRet '& IIf(stRet = "", vbNullString, ", " & stRet)
End Function


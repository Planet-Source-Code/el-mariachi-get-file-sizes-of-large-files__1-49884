VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Size"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   5160
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warp Engine Software"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      MouseIcon       =   "Form1.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   14
      ToolTipText     =   "Go to the official website"
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "API Method (Revised):"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "API Method (WinNT):"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "API Method:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "FileLen Method:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "API Method (Currency):"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The file size function included in VB will give a wrong value when used on large files.
'I believe a file larger than about 2.1GB will return a wrong value. The traditional
'way of getting the file size via API will also not work as detailed in the MS article
'http://support.microsoft.com/default.aspx?scid=kb;en-us;185476

'After a bit of research I ended up tweaking the API function to calculate the proper
'file size. I picked up bits and pieces here and there and if some of this looks like
'something that others have done...well there's only really a couple ways to do it.
'
'To see the differences, browse to a large file. The first two methods will give a negative
'value on some occasions. I believe this is because the API is expecting an unsigned long integer
'and VB can only do a signed long integer. On using the Currency (64 bit) type of variable we are
'able to do the needed calcualtion without overflowing the variable and return it as a variant.
'
'Let me know if there are any bugs or suggestions to improve performance.
'
'AGP
'Warp Engine Software
'www.WarpEngine.com


Option Explicit
Dim fnp As String

Private Sub Command1_Click()
    On Error GoTo OpenErr
    
    Text1_Change
    
    ' CancelError is True.
    cdg1.CancelError = True
    'dialog title
    cdg1.DialogTitle = "Browse to file..."
    'initialize filename
    cdg1.FileName = ""
    'the default directory
    'If OpenCnt = 1 Then cdgOpenDB.InitDir = App.Path Else cdgOpenDB.InitDir = ""
    'cdg1.InitDir = App.Path
    ' Set filters.
    cdg1.Filter = "All Files (*.*)|*.*"
    ' Specify default filter
    cdg1.FilterIndex = 1
    ' Specify default flags (checkmark on Read Only)
    'cdg1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNNoReadOnlyReturn '+ cdlOFNNoChangeDir
    cdg1.Flags = cdlOFNReadOnly
    ' Display the Open dialog box.
    cdg1.ShowOpen
    ' Get the file name and path
    fnp = cdg1.FileName
    If Trim(fnp) = "" Then Exit Sub
    Me.Text1.Text = fnp

OpenErr:
    ' User pressed Cancel button.
    If Err.Number <> 0 And Err.Number <> 32755 Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error in opening file"
        Err.Clear
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    On Error GoTo GetErr

    'get the file path
    fnp = Trim(Text1.Text)
    'If fnp = "" Or Dir(fnp) = "" Then Exit Sub
    
    'use the function built into VB
    Label1.Caption = VBFileSize(fnp) & " bytes"

    'the file size via traditional API
    Label2.Caption = Format(APIFileSize(fnp), "#,##0") & " bytes"

    'the revised traditional file size via API
    Label8.Caption = Format(APIFileSize2(fnp), "#,##0") & " bytes"

    'the file size that can be used only on NT-based systems
    Label6.Caption = Format(WinNTSize(fnp), "#,##0") & " bytes"
    
    'the method that i settled on. it works on all Windows systems
    'and the only limitation I see is that of the Currency type variable.
    'Ive tested this method on a Win98 and WinXP machine on files with
    'size 1 byte to 7.3GB and it calculated the right value.
    Label9.Caption = Format(MyFileSize(fnp), "#,##0") & " bytes"
    
GetErr:
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error in getting size"
        Err.Clear
    End If
End Sub

Private Sub lblURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        lblURL.ForeColor = vbRed
        fRet = fHandleFile("http://www.WarpEngine.com", WIN_NORMAL)
    End If
End Sub

Private Sub Text1_Change()
    Me.Label1.Caption = ""
    Me.Label2.Caption = ""
    Me.Label8.Caption = ""
    Me.Label6.Caption = ""
    Me.Label9.Caption = ""
End Sub



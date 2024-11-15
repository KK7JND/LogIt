VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "JS8Call Monitor LogIt"
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton8 
      Caption         =   "Button 8"
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin MSWinsockLib.Winsock ctlWinsock 
      Left            =   3240
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   2217
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtComments 
      Height          =   1455
      Left            =   1320
      TabIndex        =   8
      Top             =   3960
      Width           =   4215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   3480
      Width           =   4215
   End
   Begin VB.TextBox txtGrid 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtSrRx 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtSrTx 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdButton7 
      Caption         =   "Button 7"
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdButton6 
      Caption         =   "Button 6"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdButton5 
      Caption         =   "Button 5"
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdButton4 
      Caption         =   "Button 4"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdButton1 
      Caption         =   "Button 1"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdButton2 
      Caption         =   "Button 2"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdButton3 
      Caption         =   "Button 3"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cboBand 
      Height          =   315
      ItemData        =   "frmMain.frx":030A
      Left            =   1320
      List            =   "frmMain.frx":0350
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      ItemData        =   "frmMain.frx":03CA
      Left            =   1320
      List            =   "frmMain.frx":054E
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "Now"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtCall 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblToken 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Token"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   30
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblHost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Host"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblComments 
      Alignment       =   1  'Right Justify
      Caption         =   "Comments:"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblGrid 
      Alignment       =   1  'Right Justify
      Caption         =   "Grid:"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblSrRx 
      Alignment       =   1  'Right Justify
      Caption         =   "SNR-R"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblSrTx 
      Alignment       =   1  'Right Justify
      Caption         =   "SNR-S"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblBand 
      Alignment       =   1  'Right Justify
      Caption         =   "Band:"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblmode 
      Alignment       =   1  'Right Justify
      Caption         =   "Mode:"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label LblStart 
      Alignment       =   1  'Right Justify
      Caption         =   "Start:"
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblCall 
      Alignment       =   1  'Right Justify
      Caption         =   "Call:"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu pop 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu open 
         Caption         =   "&Open"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo Err_Form_Load

    basUtilites.getSettings
    cmdClear_Click
    Hook Me.hwnd
    AddIconToTray Me.hwnd, Me.Icon, Me.Icon.Handle, "JS8Call Monitor LogIt"
    Me.Hide

Exit_Form_Load:
    Exit Sub
    
Err_Form_Load:
    MsgBox "Error in frmMain:Form_Load: " & Err.Description
    Resume Exit_Form_Load

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Form_Unload

    Call Unhook
    
Exit_Form_Unload:
    Exit Sub
    
Err_Form_Unload:
    MsgBox "Error in frmMain:Form_Unload: " & Err.Description
    Resume Exit_Form_Unload

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Err_Form_QueryUnload

    Select Case UnloadMode
        Case 1, 2, 3 'If the program is being terminated by Code, Windows shutting down, or Task Manager
            Cancel = False 'Allow the program termination
            Unload Me
        Case Else
            Cancel = True 'Else disallow the termination
    End Select
    
Exit_Form_QueryUnload:
    Exit Sub
    
Err_Form_QueryUnload:
    MsgBox "Error in frmMain:Form_QueryUnload: " & Err.Description
    Resume Exit_Form_QueryUnload

End Sub
Private Sub Form_Resize()
On Error GoTo Err_Form_Resize

    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If
    
Exit_Form_Resize:
    Exit Sub
    
Err_Form_Resize:
    MsgBox "Error in frmMain:Form_Resize: " & Err.Description
    Resume Exit_Form_Resize

End Sub
Private Sub open_Click()
On Error GoTo Err_open_Click

    If Me.WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    frmMain.Show
    
Exit_open_Click:
    Exit Sub
    
Err_open_Click:
    MsgBox "Error in frmMain:open_Click: " & Err.Description
    Resume Exit_open_Click

End Sub
Private Sub about_Click()
On Error GoTo Err_about_Click

    MsgBox "Version: " & App.Major & ":" & App.Minor & ":" & App.Revision

Exit_about_Click:
    Exit Sub
    
Err_about_Click:
    MsgBox "Error in frmMain:about_Click: " & Err.Description
    Resume Exit_about_Click

End Sub
Private Sub exit_Click()
On Error GoTo Err_exit_Click

    Unload Me

Exit_exit_Click:
    Exit Sub
    
Err_exit_Click:
    MsgBox "Error in frmMain:exit_Click: " & Err.Description
    Resume Exit_exit_Click
 
End Sub
Public Sub SysTrayMouseEventHandler()
On Error GoTo Err_SysTrayMouseEventHandler

    SetForegroundWindow Me.hwnd
    PopupMenu pop, vbPopupMenuRightButton

Exit_SysTrayMouseEventHandler:
    Exit Sub
    
Err_SysTrayMouseEventHandler:
    MsgBox "Error in frmMain:SysTrayMouseEventHandler: " & Err.Description
    Resume Exit_SysTrayMouseEventHandler

End Sub
Private Sub cmdNow_Click()
On Error GoTo Err_cmdNow_Click

    basUtilites.getDate

Exit_cmdNow_Click:
    Exit Sub
    
Err_cmdNow_Click:
    MsgBox "Error in frmMain:cmdNow_Click: " & Err.Description
    Resume Exit_cmdNow_Click

End Sub
Private Sub cmdButton1_Click()
On Error GoTo Err_cmdButton1_Click

    cboMode.Text = cmdButton1.Tag

Exit_cmdButton1_Click:
    Exit Sub
    
Err_cmdButton1_Click:
    MsgBox "Error in frmMain:cmdButton1_Click: " & Err.Description
    Resume Exit_cmdButton1_Click

End Sub
Private Sub cmdButton2_Click()
On Error GoTo Err_cmdButton2_Click

    cboMode.Text = cmdButton2.Tag

Exit_cmdButton2_Click:
    Exit Sub
    
Err_cmdButton2_Click:
    MsgBox "Error in frmMain:cmdButton2_Click: " & Err.Description
    Resume Exit_cmdButton2_Click

End Sub
Private Sub cmdButton3_Click()
On Error GoTo Err_cmdButton3_Click

    cboMode.Text = cmdButton3.Tag

Exit_cmdButton3_Click:
    Exit Sub
    
Err_cmdButton3_Click:
    MsgBox "Error in frmMain:cmdButton3_Click: " & Err.Description
    Resume Exit_cmdButton3_Click

End Sub
Private Sub cmdButton4_Click()
On Error GoTo Err_cmdButton4_Click

    cboMode.Text = cmdButton4.Tag

Exit_cmdButton4_Click:
    Exit Sub
    
Err_cmdButton4_Click:
    MsgBox "Error in frmMain:cmdButton4_Click: " & Err.Description
    Resume Exit_cmdButton4_Click

End Sub
Private Sub cmdButton5_Click()
On Error GoTo Err_cmdButton5_Click

    cboMode.Text = cmdButton5.Tag

Exit_cmdButton5_Click:
    Exit Sub
    
Err_cmdButton5_Click:
    MsgBox "Error in frmMain:cmdButton5_Click: " & Err.Description
    Resume Exit_cmdButton5_Click

End Sub
Private Sub cmdButton6_Click()
On Error GoTo Err_cmdButton6_Click

    cboMode.Text = cmdButton6.Tag

Exit_cmdButton6_Click:
    Exit Sub
    
Err_cmdButton6_Click:
    MsgBox "Error in frmMain:cmdButton6_Click: " & Err.Description
    Resume Exit_cmdButton6_Click

End Sub
Private Sub cmdButton7_Click()
On Error GoTo Err_cmdButton7_Click

    cboMode.Text = cmdButton7.Tag
    
Exit_cmdButton7_Click:
    Exit Sub
    
Err_cmdButton7_Click:
    MsgBox "Error in frmMain:cmdButton7_Click: " & Err.Description
    Resume Exit_cmdButton7_Click

End Sub
Private Sub cmdButton8_Click()
On Error GoTo Err_cmdButton8_Click

    cboMode.Text = cmdButton8.Tag
    
Exit_cmdButton8_Click:
    Exit Sub
    
Err_cmdButton8_Click:
    MsgBox "Error in frmMain:cmdButton8_Click: " & Err.Description
    Resume Exit_cmdButton8_Click

End Sub
Private Sub cmdClear_Click()
On Error GoTo Err_cmdClear_Click

    ' Sets the form fields to defaults
    frmMain.txtCall = ""
    frmMain.txtDate = ""
    frmMain.cboMode.Text = frmMain.cboMode.Tag
    frmMain.cboBand.Text = frmMain.cboBand.Tag
    frmMain.txtSrTx.Text = frmMain.txtSrTx.Tag
    frmMain.txtSrRx.Text = frmMain.txtSrRx.Tag
    frmMain.txtGrid = ""
    frmMain.txtName = ""
    frmMain.txtComments = ""

Exit_cmdClear_Click:
    Exit Sub
    
Err_cmdClear_Click:
    MsgBox "Error in frmMain:cmdClear_Click: " & Err.Description
    Resume Exit_cmdClear_Click

End Sub
Private Sub cmdLog_Click()
On Error GoTo Err_cmdLog_Click

    basUtilites.sendLog

Exit_cmdLog_Click:
    Exit Sub
    
Err_cmdLog_Click:
    MsgBox "Error in frmMain:cmdLog_Click: " & Err.Description
    Resume Exit_cmdLog_Click

End Sub


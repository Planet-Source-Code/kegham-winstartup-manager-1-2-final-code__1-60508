VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00CB7834&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winstartup Manager 1.2"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6360
   FillColor       =   &H00CB7834&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00CB7834&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4770
      Left            =   0
      TabIndex        =   10
      Top             =   825
      Width           =   6360
      Begin VB.Timer tmrforce1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4665
         Top             =   1980
      End
      Begin VB.Timer tmrback 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4665
         Top             =   1065
      End
      Begin VB.Timer tmrforce 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4665
         Top             =   1515
      End
      Begin VB.ListBox lstName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2565
         ItemData        =   "frmmain.frx":27A2
         Left            =   4215
         List            =   "frmmain.frx":27A4
         TabIndex        =   14
         ToolTipText     =   $"frmmain.frx":27A6
         Top             =   330
         Width           =   1980
      End
      Begin VB.ListBox lstCmdLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1005
         ItemData        =   "frmmain.frx":2846
         Left            =   165
         List            =   "frmmain.frx":2848
         TabIndex        =   13
         ToolTipText     =   "Command line and executable path"
         Top             =   3270
         Width           =   6030
      End
      Begin VB.PictureBox mainpic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   2565
         Left            =   165
         ScaleHeight     =   2535
         ScaleWidth      =   3795
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   $"frmmain.frx":284A
         Top             =   330
         Width           =   3825
         Begin VB.OptionButton chk9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "User Startup"
            Height          =   195
            Left            =   -195
            MouseIcon       =   "frmmain.frx":2901
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   1680
            Width           =   4100
         End
         Begin VB.OptionButton chk1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_CURRENT_USER    Run"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2A53
            MousePointer    =   99  'Custom
            TabIndex        =   1
            Top             =   150
            Value           =   -1  'True
            Width           =   4100
         End
         Begin VB.OptionButton chk2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_CURRENT_USER    Run Once"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2BA5
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   405
            Width           =   4100
         End
         Begin VB.OptionButton chk3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_LOCAL_MACHINE  Run"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2CF7
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   660
            Width           =   4100
         End
         Begin VB.OptionButton chk4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_LOCAL_MACHINE  Run Once"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2E49
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   915
            Width           =   4100
         End
         Begin VB.OptionButton chk5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HKEY_LOCAL_MACHINE  Run Services"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":2F9B
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   1170
            Width           =   4100
         End
         Begin VB.OptionButton chk6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Common Startup"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":30ED
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   1425
            Width           =   4100
         End
         Begin VB.OptionButton chk7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Win.ini (Manual Edit)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":323F
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   1935
            Width           =   4100
         End
         Begin VB.OptionButton chk8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "System.ini (Manual Edit)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -195
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frmmain.frx":3391
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   2190
            Width           =   4100
         End
      End
      Begin VB.FileListBox filelistbox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   2565
         Left            =   4215
         System          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "You can always press the DEL key on your keyboard to delete the selected item or [ CTRL + C  ] to copy the selected item ."
         Top             =   330
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00D7C4AE&
         X1              =   -15
         X2              =   6345
         Y1              =   4755
         Y2              =   4755
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Executable Filename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008AD2FD&
         Height          =   210
         Left            =   4245
         TabIndex        =   18
         Top             =   75
         Width           =   1740
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Registry Run Section Key Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008AD2FD&
         Height          =   210
         Left            =   180
         TabIndex        =   17
         Top             =   75
         Width           =   3810
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Executable Path  [Start In]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008AD2FD&
         Height          =   210
         Left            =   180
         TabIndex        =   16
         Top             =   3015
         Width           =   5670
      End
      Begin VB.Label lblinf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008AD2FD&
         Height          =   240
         Left            =   180
         TabIndex        =   15
         Top             =   4365
         Width           =   5685
      End
   End
   Begin VB.Label lbldate 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   5715
      Width           =   5700
   End
   Begin VB.Image imgregg 
      Height          =   435
      Left            =   3195
      Picture         =   "frmmain.frx":34E3
      Stretch         =   -1  'True
      Top             =   360
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   -15
      Picture         =   "frmmain.frx":37ED
      Top             =   0
      Width           =   6375
   End
   Begin VB.Menu mnufle 
      Caption         =   "&File"
      Begin VB.Menu mnuCopyToClipboard 
         Caption         =   "Copy To Clipboard"
      End
      Begin VB.Menu mnuDeleteEntry 
         Caption         =   "Delete Selected Item"
      End
      Begin VB.Menu mnuterm 
         Caption         =   "Terminate Selected Task"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuRestoreBackup 
         Caption         =   "Restore Backup"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnualwaysontop 
         Caption         =   "Always On Top"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuCreatenewBackup 
         Caption         =   "Create New Backup"
      End
      Begin VB.Menu mnusepback 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlwayscreatebackuponexit 
         Caption         =   "Always Create Backup On Exit"
      End
      Begin VB.Menu mnuRestoreaprevoussavedbackup 
         Caption         =   "Restore Previous Saved Backup"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnureadme 
         Caption         =   "Read Me"
      End
      Begin VB.Menu mnuseptip 
         Caption         =   "-"
      End
      Begin VB.Menu mnucontact 
         Caption         =   "Contact Support"
      End
      Begin VB.Menu mnuvisithomepage 
         Caption         =   "Visit Homepage"
      End
      Begin VB.Menu mnusepa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Dim reg As CRegistry
Dim env As CEnvironment

Dim hKey As Long, LCount As Long, i As Long

Private Sub Command1_Click()

End Sub

Private Sub Form_Initialize()

On Error Resume Next
Dim X As Variant
X = InitCommonControls

End Sub
Private Sub Form_Load()

'On Error Resume Next

'Me.Height = 6720
'Me.Width = 6090
checkinstances
readsetting

Set reg = New CRegistry
Set env = New CEnvironment

lblinf.Caption = chk1.Caption
firstrun

End Sub

Sub checkinstances()
On Error Resume Next

If App.PrevInstance = True Then
'MsgBox Me.Caption & " Is already running", vbInformation, "No more instances"
End
  
End If
End Sub

Sub readsetting()
Dim checkback As Variant
checkback = GetSetting(App.EXEName, "settings", "AlwaysBackup", 0)
mnuAlwayscreatebackuponexit.Checked = checkback
lbldate.Caption = "Last backup " & GetSetting(App.EXEName, "settings", "BackupDate", 0)
chk1_Click
End Sub

Sub firstrun()
On Error Resume Next

Dim f As Variant
f = GetSetting(App.EXEName, "settings", "FirstRun", 0)
If f = "Done" Then
Exit Sub
Else
MsgBox "You are not allowed to compile and submit to anywhere as yours", vbInformation, "Please acept it as special note"

Me.Hide
frmwait.Show
End If

End Sub


Private Sub chk1_Click()

On Error Resume Next

If chk1.Value = True Then
lblinf.Caption = chk1.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun1
End If

End Sub
Private Sub chk2_Click()

On Error Resume Next

If chk2.Value = True Then
lblinf.Caption = chk2.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun2
End If

End Sub
Private Sub chk3_Click()

On Error Resume Next

If chk3.Value = True Then
lblinf.Caption = chk3.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun3
End If

End Sub
Private Sub chk4_Click()

On Error Resume Next

If chk4.Value = True Then
lblinf.Caption = chk4.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun4
End If

End Sub
Private Sub chk5_Click()

On Error Resume Next

If chk5.Value = True Then
lblinf.Caption = chk5.Caption
lstName.Clear
lstCmdLine.Clear
filelistbox.FileName = vbNullString
filelistbox.Visible = False
lstName.Visible = True
EnumRegRun5
End If

End Sub
Private Sub chk6_Click()

On Error Resume Next

lstName.Clear
lstCmdLine.Clear
lblinf.Caption = chk6.Caption
lstName.Visible = False

filelistbox.Refresh
filelistbox.Visible = True
Dim selfile2 As Variant
selfile2 = filelistbox.Selected(0)
filelistbox.FileName = CheckFolderID(Common_StartUp)
lstCmdLine.AddItem CheckFolderID(Common_StartUp)
filelistbox.Selected(selfile2) = True

End Sub
Private Sub chk7_Click()

On Error Resume Next
lblinf.Caption = chk7.Caption
lstName.Clear
lstCmdLine.Clear

filelistbox.Visible = False
lstName.Visible = True

MousePointer = 11
ShellExecute 0, "open", "notepad.exe", env.WindowsDirectory & "\win.ini", "", 1
MousePointer = 0

End Sub
Private Sub chk8_Click()

On Error Resume Next
filelistbox.Visible = False
lblinf.Caption = chk8.Caption
lstName.Clear
lstCmdLine.Clear
lstName.Visible = True
MousePointer = 11
ShellExecute 0, "open", "notepad.exe", env.WindowsDirectory & "\system.ini", "", 1
MousePointer = 0

End Sub
Private Sub chk9_Click()
On Error Resume Next

lstName.Clear
lstCmdLine.Clear
lblinf.Caption = chk9.Caption
lstName.Visible = False
filelistbox.Refresh
filelistbox.Visible = True
Dim selfile1 As Variant
selfile1 = filelistbox.Selected(0)
filelistbox.FileName = CheckFolderID(StartUp)
lstCmdLine.AddItem CheckFolderID(StartUp)
filelistbox.Selected(selfile1) = True

End Sub

Private Sub chk1_LostFocus()
chk1.FontBold = False
End Sub

Private Sub chk2_LostFocus()
chk2.FontBold = False
End Sub

Private Sub chk3_LostFocus()
chk3.FontBold = False
End Sub

Private Sub chk4_LostFocus()
chk4.FontBold = False
End Sub

Private Sub chk5_LostFocus()
chk5.FontBold = False
End Sub

Private Sub chk6_LostFocus()
chk6.FontBold = False
End Sub

Private Sub chk7_LostFocus()
chk7.FontBold = False
End Sub

Private Sub chk8_LostFocus()
chk8.FontBold = False
End Sub

Private Sub chk9_LostFocus()
chk9.FontBold = False
End Sub

Private Sub chk1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk1.FontBold = True
End Sub

Private Sub chk2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk2.FontBold = True
End Sub

Private Sub chk3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk3.FontBold = True
End Sub

Private Sub chk4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk4.FontBold = True
End Sub

Private Sub chk5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk5.FontBold = True
End Sub

Private Sub chk6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk6.FontBold = True
End Sub

Private Sub chk7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk7.FontBold = True
End Sub

Private Sub chk8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk8.FontBold = True
End Sub

Private Sub chk9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk9.FontBold = True

End Sub

Private Sub chk1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If
End Sub

Private Sub chk2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If
End Sub

Private Sub chk3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If
End Sub

Private Sub chk4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If
End Sub

Private Sub chk5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lstName.SetFocus
End If
End Sub

Private Sub chk6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
filelistbox.SetFocus
End If
End Sub
Private Sub chk9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
filelistbox.SetFocus
End If
End Sub

Private Sub filelistbox_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

If filelistbox.ListCount = 0 Then
Exit Sub
Else

Dim tmp As Variant
tmp = filelistbox.FileName

If KeyCode = "{CTRL + C}" Then
Clipboard.Clear
Clipboard.SetText (tmp), 1

If KeyCode = vbKeyDelete Then
If chk6.Value = True Then
mnuDeletestartupfiles
Else

If KeyCode = vbKeyDelete Then
If chk9.Value = True Then
mnuDeletestartupfilesa

End If
End If
End If
End If
End If
End If
End Sub

Private Sub filelistbox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If filelistbox.ListCount = 0 Then
Exit Sub
Else

If Button = 2 Then
PopupMenu mnufle
End If
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
chk1.FontBold = False
chk2.FontBold = False
chk3.FontBold = False
chk4.FontBold = False
chk5.FontBold = False
chk6.FontBold = False
chk7.FontBold = False
chk8.FontBold = False
chk9.FontBold = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
chk1.FontBold = False
chk2.FontBold = False
chk3.FontBold = False
chk4.FontBold = False
chk5.FontBold = False
chk6.FontBold = False
chk7.FontBold = False
chk8.FontBold = False
chk9.FontBold = False


End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
chk1.FontBold = False
chk2.FontBold = False
chk3.FontBold = False
chk4.FontBold = False
chk5.FontBold = False
chk6.FontBold = False
chk7.FontBold = False
chk8.FontBold = False
chk9.FontBold = False
End Sub

Private Sub lstCmdLine_Click()
On Error Resume Next
lstName.ListIndex = lstCmdLine.ListIndex
End Sub

Private Sub lstCmdLine_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then
mnuDeleteEntry_Click

End If
End Sub

Private Sub lstName_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Dim tmp1 As Variant
tmp1 = lstName.Text
If KeyCode = vbKeyDelete Then
mnuDeleteEntry_Click
Else
If KeyCode = "{CTRL + C}" Then
Clipboard.Clear
Clipboard.SetText (tmp1), 1
End If
End If

End Sub

Private Sub lstName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If lstName.ListCount = 0 Then
Exit Sub
Else
If Button = 2 Then
PopupMenu mnufle
End If
End If

End Sub

Private Sub mainpic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
chk1.FontBold = False
chk2.FontBold = False
chk3.FontBold = False
chk4.FontBold = False
chk5.FontBold = False
chk6.FontBold = False
chk7.FontBold = False
chk8.FontBold = False
chk9.FontBold = False
End Sub

Private Sub mnuAlwayscreatebackuponexit_Click()

On Error Resume Next
If mnuAlwayscreatebackuponexit.Checked = True Then
mnuAlwayscreatebackuponexit.Checked = False
SaveSetting App.EXEName, "Settings", "AlwaysBackup", "False"
Else
SaveSetting App.EXEName, "Settings", "AlwaysBackup", "True"
mnuAlwayscreatebackuponexit.Checked = True
End If

End Sub

Private Sub mnucontact_Click()
On Error Resume Next
ShellExecute hwnd, "open", "mailto:kegham_d@hotmail.com", vbNullString, vbNullString, 1


End Sub

Private Sub mnuCopyToClipboard_Click()

On Error Resume Next
If lstName.Visible = True Then
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lstName.Text
Else
If lstName.Visible = False Then
On Error Resume Next
Clipboard.Clear
Clipboard.SetText filelistbox.FileName
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
Dim backupopt As Variant
backupopt = GetSetting(App.EXEName, "settings", "AlwaysBackup", 0)
If backupopt = "True" Then
createbackup

'Clear everything and end here!
'Set env = Nothing
'Set reg = Nothing


End If
Exit Sub
End
End Sub

Private Sub lstName_Click()
On Error Resume Next

lstCmdLine.ListIndex = lstName.ListIndex

End Sub

Private Sub mnuabout_Click()

On Error Resume Next
frmAbout.Show

End Sub

Private Sub mnualwaysontop_Click()

On Error Resume Next
Dim RetVal As Variant

If mnualwaysontop.Checked = False Then
mnualwaysontop.Checked = True
RetVal = SetWindowPos(frmmain.hwnd, -1, 0, 0, 0, 0, 3)
Else
If mnualwaysontop.Checked = True Then
mnualwaysontop.Checked = False
RetVal = SetWindowPos(frmmain.hwnd, -2, 0, 0, 0, 0, 3)
Else
End If
End If

End Sub

Private Sub mnuCreatenewBackup_Click()
On Error Resume Next
Me.Enabled = False

lstCmdLine.Clear
lstName.Clear
createbackup

End Sub

Private Sub mnuDeleteEntry_Click()

On Error Resume Next

Dim lstnamesel As Variant
Dim cmdlinesel As Variant
Dim Val As String

lstnamesel = lstName.ListIndex
cmdlinesel = lstCmdLine.ListIndex
Val = Dir(env.SystemDirectory & "\ntoskrnl.exe")

'If startup files option box checked here
'*****************************************************
If chk6.Value = True And lstName.Visible = False Then
mnuDeletestartupfiles
chk6_Click
Else

If chk9.Value = True And lstName.Visible = False Then
mnuDeletestartupfilesa
chk9_Click
Else

'If Execution names listbox count greater than 0 then
'******************************************************
If lstName.ListCount > 0 Then
Dim ask As String
ask = MsgBox("You are about to remove this item from execution." & vbCrLf & "Item Name: " & lstName.Text, vbYesNo, "Please confirm removing it if you sure")
If ask = vbYes Then

'Begin checking which option button clicked
'*******************************************

If chk1.Value = True Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyCurrentUser, "Software\Microsoft\Windows\CurrentVersion\Run", lstName.Text
lstCmdLine.RemoveItem cmdlinesel
lstName.RemoveItem lstnamesel
lstCmdLine.Clear
lstName.Clear
lstName.Refresh
chk1_Click


Else

If chk2.Value = True Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyCurrentUser, "Software\Microsoft\Windows\CurrentVersion\RunOnce", lstName.Text
lstCmdLine.RemoveItem cmdlinesel
lstName.RemoveItem lstnamesel
lstCmdLine.Clear
lstName.Clear
lstName.Refresh
chk2_Click

Else

If chk3.Value = True Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\Run", lstName.Text
lstCmdLine.RemoveItem cmdlinesel
lstName.RemoveItem lstnamesel
lstCmdLine.Clear
lstName.Clear
lstName.Refresh
chk3_Click

Else

If chk4.Value = True Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\RunOnce", lstName.Text
lstCmdLine.RemoveItem cmdlinesel
lstName.RemoveItem lstnamesel
lstCmdLine.Clear
lstName.Clear
lstName.Refresh
chk4_Click

Else

If chk5.Value = True Then
ModDel.DeleteValue ERegRoot.eRegRoot_HKeyLocalMachine, "Software\Microsoft\Windows\CurrentVersion\RunServices", lstName.Text
lstCmdLine.RemoveItem cmdlinesel
lstName.RemoveItem lstnamesel
lstCmdLine.Clear
lstName.Clear
lstName.Refresh
chk5_Click

Else
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Sub mnuDeletestartupfiles()

If filelistbox.ListCount = 0 Then
Exit Sub
Else

Dim Val As String
Dim askfiledel As Variant
Dim startupfolder As Variant
Dim sFile As Variant

Val = Dir(env.SystemDirectory & "\ntoskrnl.exe")
startupfolder = CheckFolderID(Common_StartUp)
sFile = startupfolder & "\" & (filelistbox.FileName)

askfiledel = MsgBox("Are you sure you want to delete the selected file from startup directory", vbYesNo, "Confirm please")
If askfiledel = vbNo Then
Exit Sub
Else

On Error GoTo frunning

Kill (sFile)
filelistbox.Refresh
chk6_Click

Exit Sub

frunning:
If Val = "ntoskrnl.exe" Then
Dim askforce
askforce = MsgBox("This item must be terminated in order to delete it", vbYesNo, "Do you want to forcefully terminate it")
If askforce = vbYes Then
Shell "Taskkill -F -IM " & filelistbox.FileName, 0
tmrforce.Enabled = True

End If
End If
End If
End If


End Sub
Sub mnuDeletestartupfilesa()

If filelistbox.ListCount = 0 Then
Exit Sub
Else

Dim Vala As String
Dim askfiledela As Variant
Dim startupfoldera As Variant
Dim sFilea As Variant

Vala = Dir(env.SystemDirectory & "\ntoskrnl.exe")
startupfoldera = CheckFolderID(StartUp)
sFilea = startupfoldera & "\" & (filelistbox.FileName)

askfiledela = MsgBox("Are you sure you want to delete the selected file from startup directory", vbYesNo, "Confirm please to delete")
If askfiledela = vbNo Then
Exit Sub
Else

On Error GoTo frunninga

Kill (sFilea)
filelistbox.Refresh
chk9_Click

Exit Sub

frunninga:
If Vala = "ntoskrnl.exe" Then
Dim askforcea
askforcea = MsgBox("This item must be terminated in order to delete it", vbYesNo, "Do you want to forcefully terminate it")
If askforcea = vbYes Then
Shell "Taskkill -F -IM " & filelistbox.FileName, 0
tmrforce1.Enabled = True

End If
End If
End If
End If


End Sub

Private Sub mnuexit_Click()
On Error Resume Next

Unload Me
End
End Sub



Private Sub mnureadme_Click()
On Error Resume Next
Dim fl
fl = Dir("readme.txt")
If fl <> "readme.txt" Then
MsgBox "Readme file has been modified please reinstall Winstartup Manager", vbInformation, "File not found"
Else

Shell "readme.txt", 3
End If

End Sub

Private Sub mnuRestoreaprevoussavedbackup_Click()
On Error Resume Next

Dim askrestore As Variant
askrestore = MsgBox("This will restore a pre saved backup do you wish to continue", vbInformation + vbYesNo, "Yes to restore no to not")
If askrestore = vbYes Then
On Error Resume Next

restorebackup
Me.Enabled = False
lstCmdLine.Clear
lstName.Clear



'Exit Sub
End If
End Sub

Private Sub mnuterm_Click()

If filelistbox.Visible = False Then

Dim Val As String
Dim selfile As Variant

Val = Dir(env.SystemDirectory & "\ntoskrnl.exe")
selfile = lstName.Text
If Val = "ntoskrnl.exe" Then

Dim askterm
askterm = MsgBox("This will terminate the process named: " & lstName.Text, vbYesNo, "Do you want to continue?")
If askterm = vbYes Then
Shell "Taskkill -f -IM " & lstName.Text, 0
lbldate.Caption = " Selected task has been terminated"
tmrback.Enabled = True
'Help if win9x method how to end task it from command line

Exit Sub
End If
End If
End If


End Sub

Private Sub mnuvisithomepage_Click()
On Error Resume Next
ShellExecute hwnd, "open", "http://www.vbdotlb.connect.to", vbNullString, vbNullString, 1

End Sub

Sub createbackup()

On Error Resume Next

Me.Enabled = False


Dim curdate As Variant
Dim curtime As Variant

curdate = Date
curtime = Time

Shell "regedit /e HKCU_Run.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
Shell "regedit /e HKCU_RunOnce.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
Shell "regedit /e HKLM_Run.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run"
Shell "regedit /e HKLM_RunOnce.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce"
Shell "regedit /e HKLM_RunServices.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices"

SaveSetting App.EXEName, "Settings", "BackupDate", "Date " & Date & " Time: " & Time

lbldate.Caption = "A new backup has been created"
tmrback.Enabled = True

End Sub
Sub restorebackup()
On Error Resume Next


Shell "regedit /is HKCU_Run.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
Shell "regedit /is HKCU_RunOnce.reg HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
Shell "regedit /is HKLM_Run.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run"
Shell "regedit /is HKLM_RunOnce.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunOnce"
Shell "regedit /is HKLM_RunServices.reg HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices"
'lblinf.Caption = "Restoring complete"

'chk1_Click

lbldate.Caption = "A pre saved Backup has been restored"
tmrback.Enabled = True

Exit Sub
End Sub

Sub EnumRegRun1()
On Error Resume Next
    
hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1
lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))

Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0
    
End Sub

Sub EnumRegRun2()
On Error Resume Next

hKey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1
        
lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0
    
End Sub

Sub EnumRegRun3()
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1
lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0

End Sub

Sub EnumRegRun4()
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1
lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))

Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0

End Sub

Sub EnumRegRun5()
On Error Resume Next

hKey = OpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices")
LCount = GetCount(hKey, Values)
For i = 0 To LCount - 1
lstName.AddItem EnumValue(hKey, i)
lstCmdLine.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
Next i
lstName.ListIndex = 0
lstCmdLine.ListIndex = 0
End Sub

Private Sub tmrback_Timer()
chk1_Click
lbldate.Caption = "Last backup " & GetSetting(App.EXEName, "settings", "BackupDate", 0)

tmrback.Enabled = False
lstName.Refresh
filelistbox.Refresh
lstCmdLine.Refresh

Me.Enabled = True

End Sub

Private Sub tmrforce_Timer()

Dim startupfolder, sFile

startupfolder = CheckFolderID(Common_StartUp)
sFile = startupfolder & "\" & (filelistbox.FileName)

On Error Resume Next
Kill (sFile)
filelistbox.Refresh
tmrforce.Enabled = False
chk6_Click

End Sub

Private Sub tmrforce1_Timer()

Dim startupfoldera, sFilea

startupfoldera = CheckFolderID(StartUp)
sFilea = startupfoldera & "\" & (filelistbox.FileName)

On Error Resume Next
Kill (sFilea)
filelistbox.Refresh
tmrforce1.Enabled = False
chk9_Click

End Sub

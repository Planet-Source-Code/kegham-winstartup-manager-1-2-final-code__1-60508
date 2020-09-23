VERSION 5.00
Begin VB.Form frmwait 
   Appearance      =   0  'Flat
   BackColor       =   &H00CB7834&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Please Wait ..."
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   945
      Top             =   1245
   End
   Begin VB.Label lblprb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   45
      TabIndex        =   1
      Top             =   525
      Width           =   30
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Creating first runtime backup."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   4230
   End
   Begin VB.Shape shmain 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   30
      Top             =   1005
      Width           =   4245
   End
End
Attribute VB_Name = "frmwait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long


Private Sub Form_Initialize()
On Error Resume Next

Dim X As Variant
X = InitCommonControls

End Sub

Private Sub Form_Load()

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

On Error Resume Next

lblprb.Visible = True
lblprb.Width = lblprb.Width + 10

If lblprb.Width > shmain.Width Then
SaveSetting App.EXEName, "Settings", "FirstRun", "Done"

frmmain.createbackup
Timer1.Enabled = False
Unload Me
frmmain.lbldate.Caption = " First backup created" & GetSetting(App.EXEName, "settings", "BackupDate", 0)
frmmain.Show
End If

End Sub


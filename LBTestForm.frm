VERSION 5.00
Object = "{F3731EEC-A7CE-4247-8095-37823063E354}#3.0#0"; "LightButton.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin LightButtonControl.LightButton LightButton1 
      Height          =   975
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1720
      BackColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelPicture      =   "LBTestForm.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   3840
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  LightButton1.Flash 5
End Sub

Private Sub LightButton1_Click()
  MsgBox "ok"
End Sub


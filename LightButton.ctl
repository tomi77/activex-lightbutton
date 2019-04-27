VERSION 5.00
Begin VB.UserControl LightButton 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   ScaleHeight     =   990
   ScaleWidth      =   3870
   Begin VB.Timer tmrChkStatus 
      Interval        =   250
      Left            =   3000
      Top             =   120
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "LightButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum lbModeTypes
  [Text Only Mode] = 0
  [Image Mode] = 1
End Enum

Public Enum lbBorderStyleTypes
  None = 0
  [Fixed Single]
End Enum

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event StatusChanged(NewStatus As Integer)
Attribute StatusChanged.VB_MemberFlags = "200"

Private mbooButtonLighted As Boolean
Private mfonFont As StdFont
Private mmodButtonMode As lbModeTypes
Private molcBackColor As OLE_COLOR
Private molcSelColor As OLE_COLOR
Private mpicPicture As New StdPicture
Private mpicSelPicture As New StdPicture
Private mpoiCursorPos As POINTAPI

Private Sub tmrChkStatus_Timer()

  Dim lonCStat As Long
  Dim lonCurrhWnd As Long

  tmrChkStatus.Enabled = False
  
  lonCStat = GetCursorPos(mpoiCursorPos)
  lonCurrhWnd = WindowFromPoint(mpoiCursorPos.x, mpoiCursorPos.y)
  
  If mbooButtonLighted = False Then
    If lonCurrhWnd = UserControl.hWnd Then
      mbooButtonLighted = True
      If mmodButtonMode = [Text Only Mode] Then
        UserControl.BackColor = molcSelColor
      Else
        Set UserControl.Picture = mpicSelPicture
      End If
      RaiseEvent StatusChanged(1)
    End If
  Else
    If lonCurrhWnd <> UserControl.hWnd Then
      mbooButtonLighted = False
      If mmodButtonMode = [Text Only Mode] Then
        UserControl.BackColor = molcBackColor
      Else
        Set UserControl.Picture = mpicPicture
      End If
      RaiseEvent StatusChanged(0)
    End If
  End If
  
  tmrChkStatus.Enabled = True

End Sub

Private Sub UserControl_Click()

  RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()

  RaiseEvent DblClick

End Sub


Private Sub UserControl_Initialize()

  mbooButtonLighted = False
  
  RaiseEvent StatusChanged(0)
  
  molcBackColor = UserControl.BackColor

End Sub

Private Sub UserControl_InitProperties()

  tmrChkStatus.Enabled = Ambient.UserMode
  
  Caption = Ambient.DisplayName

  SelColor = &H80FFFF
  
  ButtonMode = [Text Only Mode]

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

  RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

  RaiseEvent KeyPress(KeyAscii)

End Sub


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

  RaiseEvent KeyUp(KeyCode, Shift)

End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  RaiseEvent MouseDown(Button, Shift, x, y)

End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  RaiseEvent MouseMove(Button, Shift, x, y)

End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  RaiseEvent MouseUp(Button, Shift, x, y)

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  tmrChkStatus.Enabled = Ambient.UserMode

  BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
  BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  ButtonMode = PropBag.ReadProperty("ButtonMode", mmodButtonMode)
  Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
  Enabled = PropBag.ReadProperty("Enabled", Enabled)
  ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
  SelColor = PropBag.ReadProperty("SelColor", &H80FFFF)
  
  Set Font = PropBag.ReadProperty("Font", mfonFont)
  Set Picture = PropBag.ReadProperty("Picture", Nothing)
  Set SelPicture = PropBag.ReadProperty("SelPicture", Nothing)

End Sub

Private Sub UserControl_Resize()

  lblCaption.Top = (Height - lblCaption.Height) / 2
  lblCaption.Left = (Width - lblCaption.Width) / 2

End Sub

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"

  Caption = lblCaption.Caption

End Property

Public Property Let Caption(ByVal NewValue As String)

  lblCaption.Caption = NewValue
  UserControl.PropertyChanged "Caption"

End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  PropBag.WriteProperty "BackColor", BackColor, &HFFFFFF
  PropBag.WriteProperty "BorderStyle", BorderStyle, 1
  PropBag.WriteProperty "ButtonMode", ButtonMode, mmodButtonMode
  PropBag.WriteProperty "Caption", Caption, Ambient.DisplayName
  PropBag.WriteProperty "Enabled", Enabled, True
  PropBag.WriteProperty "ForeColor", ForeColor, &H80000008
  PropBag.WriteProperty "SelColor", SelColor, &H80FFFF
  
  PropBag.WriteProperty "Font", Font, mfonFont
  PropBag.WriteProperty "Picture", Picture, Nothing
  PropBag.WriteProperty "SelPicture", SelPicture, Nothing

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"

  BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)

  UserControl.BackColor = NewValue
  UserControl.PropertyChanged "BackColor"
  
  molcBackColor = NewValue

End Property

Public Property Get BorderStyle() As lbBorderStyleTypes
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"

  BorderStyle = UserControl.BorderStyle

End Property

Public Property Let BorderStyle(ByVal NewValue As lbBorderStyleTypes)

  If NewValue = None Or NewValue = [Fixed Single] Then
    UserControl.BorderStyle = NewValue
    UserControl.PropertyChanged "BorderStyle"
  Else
    Err.Raise vbObjectError + 32112, , "Invalid BorderStyle value (0 or 1 only)"
  End If

End Property

Public Property Get ButtonMode() As lbModeTypes
Attribute ButtonMode.VB_ProcData.VB_Invoke_Property = ";Behavior"

  ButtonMode = mmodButtonMode

End Property

Public Property Let ButtonMode(ByVal NewValue As lbModeTypes)

  If NewValue = [Text Only Mode] Or NewValue = [Image Mode] Then
    mmodButtonMode = NewValue
    If mmodButtonMode = [Text Only Mode] Then _
      lblCaption.Visible = True
    If mmodButtonMode = [Image Mode] Then _
      lblCaption.Visible = False
    UserControl.PropertyChanged "ButtonMode"
  Else
    Err.Raise vbObjectError + 32113, , "Invalid ButtonMode value (0 or 1 only)"
  End If

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"

  Set Font = lblCaption.Font

End Property

Public Property Set Font(ByVal NewValue As StdFont)

  Set lblCaption.Font = NewValue
  UserControl.PropertyChanged "Font"

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"

  ForeColor = lblCaption.ForeColor

End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)

  lblCaption.ForeColor = NewValue
  UserControl.PropertyChanged "ForeColor"

End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"

  Set Picture = UserControl.Picture

End Property

Public Property Set Picture(ByVal NewValue As StdPicture)

  Set UserControl.Picture = NewValue
  
  Set mpicPicture = NewValue
  
  If NewValue Is Nothing Then
    ButtonMode = [Text Only Mode]
  Else
    ButtonMode = [Image Mode]
  End If
  
  UserControl.PropertyChanged "Picture"

End Property

Public Property Get SelColor() As OLE_COLOR
Attribute SelColor.VB_ProcData.VB_Invoke_Property = ";Appearance"

  SelColor = molcSelColor

End Property

Public Property Let SelColor(ByVal NewValue As OLE_COLOR)

  molcSelColor = NewValue
  UserControl.PropertyChanged "SelColor"

End Property

Public Property Get SelPicture() As StdPicture
Attribute SelPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"

  Set SelPicture = mpicSelPicture

End Property

Public Property Set SelPicture(ByVal NewValue As StdPicture)

  Set mpicSelPicture = NewValue
  UserControl.PropertyChanged "SelPicture"

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514

  Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal NewValue As Boolean)

  UserControl.Enabled = NewValue
  UserControl.PropertyChanged "Enabled"

End Property

Public Sub Refresh()

  UserControl.Refresh

End Sub

Public Sub Flash(NumTimes As Integer)

  Dim booButtonLighted As Boolean
  Dim intFlashLoop As Integer
  Dim sinOldTimer As Single
  
  If NumTimes <= 0 Then Exit Sub
  
  booButtonLighted = mbooButtonLighted
  
  For intFlashLoop = 1 To (NumTimes * 2)
    booButtonLighted = Not booButtonLighted
    If booButtonLighted = True Then
      If mmodButtonMode = [Text Only Mode] Then
        UserControl.BackColor = molcSelColor
      Else
        Set UserControl.Picture = mpicSelPicture
      End If
    Else
      If mmodButtonMode = [Text Only Mode] Then
        UserControl.BackColor = molcBackColor
      Else
        Set UserControl.Picture = mpicPicture
      End If
    End If
    sinOldTimer = Timer
    Do
      DoEvents
    Loop Until (Timer >= sinOldTimer + 0.5)
  Next intFlashLoop

End Sub

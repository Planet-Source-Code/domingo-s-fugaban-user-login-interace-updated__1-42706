VERSION 5.00
Begin VB.UserControl DomsWinXP 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer tmrMouseOver 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1080
      Top             =   0
   End
   Begin VB.Image imgXPHovL 
      Height          =   315
      Left            =   0
      Picture         =   "DomsWinXP.ctx":0000
      Top             =   2520
      Width           =   90
   End
   Begin VB.Image imgXPHovM 
      Height          =   315
      Left            =   120
      Picture         =   "DomsWinXP.ctx":022F
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   690
   End
   Begin VB.Image imgXPHovR 
      Height          =   315
      Left            =   840
      Picture         =   "DomsWinXP.ctx":0455
      Top             =   2520
      Width           =   90
   End
   Begin VB.Image imgXPDisR 
      Height          =   315
      Left            =   840
      Picture         =   "DomsWinXP.ctx":0684
      Top             =   2160
      Width           =   90
   End
   Begin VB.Image imgXPDisM 
      Height          =   315
      Left            =   120
      Picture         =   "DomsWinXP.ctx":07C9
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   690
   End
   Begin VB.Image imgXPDisL 
      Height          =   315
      Left            =   0
      Picture         =   "DomsWinXP.ctx":090F
      Top             =   2160
      Width           =   90
   End
   Begin VB.Image imgXPDefL 
      Height          =   315
      Left            =   0
      Picture         =   "DomsWinXP.ctx":0B39
      Top             =   1800
      Width           =   90
   End
   Begin VB.Image imgXPDefM 
      Height          =   315
      Left            =   120
      Picture         =   "DomsWinXP.ctx":0D68
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   690
   End
   Begin VB.Image imgXPDefR 
      Height          =   315
      Left            =   840
      Picture         =   "DomsWinXP.ctx":0F85
      Top             =   1800
      Width           =   90
   End
   Begin VB.Image imgXPNorR 
      Height          =   315
      Left            =   840
      Picture         =   "DomsWinXP.ctx":11B2
      Top             =   1440
      Width           =   90
   End
   Begin VB.Image imgXPNorM 
      Height          =   315
      Left            =   120
      Picture         =   "DomsWinXP.ctx":13E1
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   690
   End
   Begin VB.Image imgXPNorL 
      Height          =   315
      Left            =   0
      Picture         =   "DomsWinXP.ctx":1609
      Top             =   1440
      Width           =   90
   End
   Begin VB.Image imgXPPreR 
      Height          =   315
      Left            =   840
      Picture         =   "DomsWinXP.ctx":1838
      Top             =   1080
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgXPPreM 
      Height          =   315
      Left            =   120
      Picture         =   "DomsWinXP.ctx":1A66
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgXPPreL 
      Height          =   315
      Left            =   0
      Picture         =   "DomsWinXP.ctx":1C8A
      Top             =   1080
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image imgXP 
      Height          =   255
      Left            =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DomsWinXP"
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
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   840
   End
   Begin VB.Image imgXPL 
      Height          =   315
      Left            =   0
      Picture         =   "DomsWinXP.ctx":1EB6
      Top             =   0
      Width           =   90
   End
   Begin VB.Image imgXPM 
      Height          =   315
      Left            =   120
      Picture         =   "DomsWinXP.ctx":20E5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   690
   End
   Begin VB.Image imgXPR 
      Height          =   315
      Left            =   840
      Picture         =   "DomsWinXP.ctx":230D
      Top             =   0
      Width           =   90
   End
End
Attribute VB_Name = "DomsWinXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variable Declarations
Private blnOnHover As Boolean
Private blnOnFocus As Boolean
Private blnOnKeyUp As Boolean

'Event Declarations
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub imgXP_Click()
    RaiseEvent Click
End Sub

Private Sub imgXP_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub imgXP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEvents

    imgXPL.Picture = imgXPPreL.Picture
    imgXPM.Picture = imgXPPreM.Picture
    imgXPR.Picture = imgXPPreR.Picture
    
    blnOnKeyUp = False
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgXP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X >= 0 And Y >= 0) And (X < imgXP.Width And Y < imgXP.Height) And (Not blnOnHover) Then
        imgXPL.Picture = imgXPHovL.Picture
        imgXPM.Picture = imgXPHovM.Picture
        imgXPR.Picture = imgXPHovR.Picture
        
        blnOnHover = True
        blnOnKeyUp = True
        tmrMouseOver.Enabled = True
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgXP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgXPL.Picture = imgXPDefL.Picture
    imgXPM.Picture = imgXPDefM.Picture
    imgXPR.Picture = imgXPDefR.Picture
    
    blnOnKeyUp = True
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub tmrMouseOver_Timer()
    Dim pntAPI As POINTAPI
    
    GetCursorPos pntAPI
    
    If (hWnd <> WindowFromPoint(pntAPI.X, pntAPI.Y)) And (blnOnKeyUp) Then
        If (blnOnFocus) Then
            imgXPL.Picture = imgXPDefL.Picture
            imgXPM.Picture = imgXPDefM.Picture
            imgXPR.Picture = imgXPDefR.Picture
        Else
            imgXPL.Picture = imgXPNorL.Picture
            imgXPM.Picture = imgXPNorM.Picture
            imgXPR.Picture = imgXPNorR.Picture
        End If
        
        blnOnHover = False
        tmrMouseOver.Enabled = False
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If (PropertyName = "BackColor") Then
        UserControl.BackColor = Ambient.BackColor
    End If
End Sub

Private Sub UserControl_GotFocus()
    imgXPL.Picture = imgXPDefL.Picture
    imgXPM.Picture = imgXPDefM.Picture
    imgXPR.Picture = imgXPDefR.Picture
    
    blnOnFocus = True
End Sub

Private Sub UserControl_LostFocus()
    imgXPL.Picture = imgXPNorL.Picture
    imgXPM.Picture = imgXPNorM.Picture
    imgXPR.Picture = imgXPNorR.Picture
    
    blnOnFocus = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "DomsWinXP")
    
    If (UserControl.Enabled = True) Then
        imgXPL.Picture = imgXPNorL.Picture
        imgXPM.Picture = imgXPNorM.Picture
        imgXPR.Picture = imgXPNorR.Picture
    Else
        imgXPL.Picture = imgXPDisL.Picture
        imgXPM.Picture = imgXPDisM.Picture
        imgXPR.Picture = imgXPDisR.Picture
    End If
End Sub

Private Sub UserControl_Resize()
    If (UserControl.Width < (imgXPL.Width + imgXPR.Width)) Then
        UserControl.Width = (imgXPL.Width + imgXPR.Width)
    End If
    
    UserControl.Height = imgXPM.Height

    imgXPL.Top = 0
    imgXPL.Left = 0
    
    imgXPR.Top = 0
    imgXPR.Left = UserControl.Width - imgXPL.Width
    
    imgXPM.Top = 0
    imgXPM.Left = imgXPL.Width
    imgXPM.Width = imgXPR.Left - imgXPL.Width
      
    lblCaption.Width = UserControl.Width
    lblCaption.Top = (UserControl.Height / 2) - (lblCaption.Height / 2)
    
    imgXP.Top = 0
    imgXP.Left = 0
    imgXP.Height = imgXPM.Height
    imgXP.Width = UserControl.Width
    
    blnOnHover = False
    blnOnFocus = False
    blnOnKeyUp = False
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "DomsWinXP")
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeySpace) Then
        imgXPL.Picture = imgXPPreL.Picture
        imgXPM.Picture = imgXPPreM.Picture
        imgXPR.Picture = imgXPPreR.Picture
    End If
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeySpace) Then
        imgXPL.Picture = imgXPDefL.Picture
        imgXPM.Picture = imgXPDefM.Picture
        imgXPR.Picture = imgXPDefR.Picture
    End If
        
    RaiseEvent Click
End Sub


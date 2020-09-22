VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Doms Video Rental Software"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "l"
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin prjLogin.DomsWinXP cmdCancel 
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "Cancel"
   End
   Begin prjLogin.DomsWinXP cmdOK 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "OK"
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Clearance Required"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4215
   End
   Begin VB.Line lneDivider 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   4680
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Shape shpBanner 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdbLogin As Database

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim rstUsers As Recordset
    
    If (Len(Trim(txtUserName.Text)) > 0) Then
        Set mdbLogin = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
        Set rstUsers = mdbLogin.OpenRecordset("SELECT * FROM Users WHERE UserName='" & _
            Trim(txtUserName.Text) & "'")
    Else
        MsgBox "User Name field is empty!", vbInformation, "Login Error"
        txtUserName.SetFocus
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
        Exit Sub
    End If
    
    With rstUsers
        If (.RecordCount > 0) Then
            If (.Fields("Password").Value = txtPassword.Text) Then
                'Access Granted! Your code goes here
                
                MsgBox "Access Granted! Your code goes here"
            Else
                MsgBox "Invalid Password! Try again!", vbInformation, "Login Error"
                txtPassword.SetFocus
                txtPassword.SelStart = 0
                txtPassword.SelLength = Len(txtPassword.Text)
            End If
        Else
            MsgBox "User Name does not exists!", vbInformation, "Login Error"
            txtUserName.SetFocus
            txtUserName.SelStart = 0
            txtUserName.SelLength = Len(txtUserName.Text)
        End If
    End With
    
    rstUsers.Close
    Set rstUsers = Nothing
    
    mdbLogin.Close
    Set mdbLogin = Nothing
End Sub

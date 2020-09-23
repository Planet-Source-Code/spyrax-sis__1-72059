VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmAdminSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to SIS!"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdminSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList MyImage 
      Left            =   5370
      Top             =   2790
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminSetup.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdminSetup.frx":0EE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdsave 
      Height          =   465
      Left            =   390
      TabIndex        =   16
      Top             =   3690
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "&Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16443093
      FCOL            =   16711680
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmAdminSetup.frx":147E
      ALIGN           =   1
      IMGLST          =   "MyImage"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox txtverifypassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2010
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2730
      Width           =   3285
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2010
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2265
      Width           =   3285
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2010
      TabIndex        =   4
      Top             =   1860
      Width           =   3285
   End
   Begin VB.ComboBox cmbuser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      ItemData        =   "frmAdminSetup.frx":149A
      Left            =   2010
      List            =   "frmAdminSetup.frx":14A4
      TabIndex        =   7
      Text            =   "Please Select!"
      Top             =   3120
      Width           =   3285
   End
   Begin VB.TextBox txtFname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2025
      TabIndex        =   0
      Top             =   105
      Width           =   3285
   End
   Begin VB.TextBox txtLname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2025
      TabIndex        =   2
      Top             =   960
      Width           =   3285
   End
   Begin VB.TextBox txtContact 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2025
      TabIndex        =   3
      Top             =   1380
      Width           =   3285
   End
   Begin VB.TextBox txtMname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   2025
      TabIndex        =   1
      Top             =   525
      Width           =   3285
   End
   Begin VB.Timer timerblink 
      Interval        =   525
      Left            =   5220
      Top             =   690
   End
   Begin LVbuttons.LaVolpeButton cmdcancel 
      Height          =   465
      Left            =   3870
      TabIndex        =   17
      Top             =   3690
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   16443093
      FCOL            =   192
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmAdminSetup.frx":14BD
      ALIGN           =   1
      IMGLST          =   "MyImage"
      IMGICON         =   "2"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C&onfirm Password:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   375
      TabIndex        =   8
      Top             =   2745
      Width           =   1500
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1050
      TabIndex        =   13
      Top             =   2325
      Width           =   825
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   990
      TabIndex        =   15
      Top             =   1905
      Width           =   900
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Level:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   975
      TabIndex        =   14
      Top             =   3165
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&First Name:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   900
      TabIndex        =   12
      Top             =   210
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Last Name:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   930
      TabIndex        =   11
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Contact #:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   10
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Middle Name:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   9
      Top             =   630
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   3465
      Left            =   30
      Top             =   60
      Width           =   1935
   End
End
Attribute VB_Name = "frmAdminSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const Cap As String = "Welcome to SIS!"

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
If MyTxtEmpty(txtFname) = True Then Exit Sub
If MyTxtEmpty(txtMname) = True Then Exit Sub
If MyTxtEmpty(txtLname) = True Then Exit Sub
If MyTxtEmpty(txtUsername) = True Then Exit Sub
If MyTxtEmpty(txtPassword) = True Then Exit Sub
If MyTxtEmpty(txtContact) = True Then Exit Sub


Set MyRsUser = New ADODB.Recordset
MyRsUser.Open "TblUsers", libcon, adOpenDynamic, adLockOptimistic

With MyRsUser
    .AddNew
    .Fields("username") = txtUsername.Text
    .Fields("password") = txtPassword.Text
    .Fields("fname") = UCase$(txtLname.Text)
    .Fields("gname") = UCase$(txtFname.Text)
    .Fields("mname") = UCase$(txtMname.Text)
    .Fields("cpnum") = txtContact.Text
    .Fields("userlevel") = cmbuser.Text
    .Update
    .Close
End With
MsgBox "New USER ACCOUNT has been added!", vbInformation, "SIS"
Unload Me
End Sub

Private Sub Form_Load()
MyDBcon
FormCenter Me

End Sub

Private Sub timerblink_Timer()
Static i As Boolean
If i = True Then
 frmAdminSetup.Caption = ""
 i = False
Else
 frmAdminSetup.Caption = Cap
 i = True
End If
End Sub


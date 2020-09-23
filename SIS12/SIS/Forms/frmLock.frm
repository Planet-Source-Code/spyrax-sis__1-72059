VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmLock 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1836
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin LVbuttons.LaVolpeButton cmdunlock 
         Height          =   405
         Left            =   1920
         TabIndex        =   4
         Top             =   1350
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&Unlock"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
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
         MICON           =   "frmLock.frx":0000
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Timer Timer2 
         Interval        =   525
         Left            =   270
         Top             =   1470
      End
      Begin VB.TextBox txtPass 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   216
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   840
         Width           =   5565
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   2835
         Top             =   3420
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4770
         Picture         =   "frmLock.frx":001C
         Top             =   240
         Width           =   480
      End
      Begin VB.Shape Shape1 
         Height          =   1065
         Left            =   150
         Top             =   240
         Width           =   5685
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER PASSWORD TO UNLOCK"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1260
         TabIndex        =   2
         Top             =   300
         Width           =   3435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER PASSWORD TO UNLOCK"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1260
         TabIndex        =   3
         Top             =   270
         Width           =   3435
      End
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const Cap As String = "ENTER PASSWORD TO UNLOCK"


Private Sub cmdUnlock_Click()
    If txtPass.Text = Mypass Then
        Unload Me
    Else
        MsgBox "Wrong password supplied. Attempt to unlock failed.", vbOKOnly + vbExclamation, "Library System"
       txtPass.Text = ""
       txtPass.SetFocus
        Exit Sub
    End If


End Sub



Private Sub Form_Load()

FormCenter Me

End Sub

Private Sub Timer1_Timer()


    If Trim(txtPass.Text) = "" Then
        cmdunlock.Visible = False
    Else
        cmdunlock.Visible = True
    End If



End Sub

Private Sub Timer2_Timer()
Static i As Boolean
If i = True Then
 Label2.Caption = ""
 i = False
Else
 Label2.Caption = Cap
 i = True
End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdUnlock_Click
    End If

End Sub

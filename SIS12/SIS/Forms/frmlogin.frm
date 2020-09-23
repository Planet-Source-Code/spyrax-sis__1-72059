VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2805
      Left            =   30
      TabIndex        =   2
      Top             =   -30
      Width           =   5835
      Begin VB.TextBox txtpassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1410
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   1530
         Width           =   4215
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4350
         TabIndex        =   4
         Top             =   2190
         Width           =   1305
      End
      Begin VB.TextBox txtusername 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   1410
         TabIndex        =   0
         Top             =   990
         Width           =   4215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Enter your username and Password to Login."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1440
         TabIndex        =   5
         Top             =   180
         Width           =   4125
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   930
         Picture         =   "frmlogin.frx":0000
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   60
         TabIndex        =   6
         Top             =   1530
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   90
         TabIndex        =   3
         Top             =   990
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2805
      Left            =   30
      TabIndex        =   8
      Top             =   -30
      Visible         =   0   'False
      Width           =   5835
      Begin VB.CommandButton cmdok 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4350
         TabIndex        =   9
         Top             =   2190
         Width           =   1305
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   360
         Top             =   2640
      End
      Begin VB.Label lblcounter 
         Alignment       =   2  'Center
         Caption         =   "counter"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1155
         Left            =   990
         TabIndex        =   12
         Top             =   900
         Width           =   4335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ACCESS DENIED!"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   24
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   1020
         TabIndex        =   11
         Top             =   300
         Width           =   4395
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   450
         Picture         =   "frmlogin.frx":076A
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ACCESS DENIED!"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   24
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   960
         TabIndex        =   10
         Top             =   300
         Width           =   4395
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESS DENIED!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   4395
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Cap As String = "ACCESS DENIED!"
Dim mycounter As Integer

Private Sub cmdcancel_Click()
End
End Sub

Private Sub cmdok_Click()
Frame1.Visible = True
txtusername.Text = ""
txtpassword.Text = ""
txtusername.SetFocus

End Sub

Private Sub Form_Load()
MyDBcon
FormCenter Me


Set MyRsUser = New ADODB.Recordset
SqlStr = "Select * from tblusers"
MyRsUser.Open SqlStr, libcon, adOpenDynamic, adLockOptimistic
mycounter = 3

End Sub


Private Sub Timer2_Timer()
Static i As Boolean
If i = True Then
 Me.Label7.Caption = ""
 i = False
Else
 Me.Label7.Caption = Cap
 i = True
End If
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
 
If KeyAscii = 13 Then
    If MyTxtEmpty(txtusername) = True Then Exit Sub
    If MyTxtEmpty(txtpassword) = True Then Exit Sub


With MyRsUser
    .Requery
    .Find "username = '" & txtusername.Text & "' "
    If .EOF Then
        mycounter = mycounter - 1
        If mycounter = 0 Then
            Frame1.Visible = False
            Frame2.Visible = True
            lblcounter.Caption = "You already used all attempt." & vbCrLf & "This will terminate the application."""
            End
        End If
            Frame1.Visible = False
            Frame2.Visible = True
            lblcounter.Caption = "The User USERNAME/PASSWORD you entered is not valid." & vbCrLf & "Please try again." & vbCrLf & "Warning: You only have " & mycounter & " attempt."
           
        Exit Sub
   Else
        
        If .Fields("password") = txtpassword.Text Then
               Mypass = MyRsUser.Fields("Password") 'to hold the user's password to be to unlock the system
               Unload Me
        Else
            mycounter = mycounter - 1
            If mycounter = 0 Then
                Frame1.Visible = False
                Frame2.Visible = True
                MsgBox "You already used all attempt." & vbCrLf & "This will terminate the application.", vbCritical, "SIS"
                End
            End If
                Frame1.Visible = False
                Frame2.Visible = True
                lblcounter.Caption = "The User USERNAME/PASSWORD you entered is not valid." & vbCrLf & "Please try again." & vbCrLf & "Warning: You only have " & mycounter & " attempt."
               
            Exit Sub
    
        End If
    End If
   
End With
End If
End Sub


Private Sub txtusername_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtpassword.SetFocus
End If
End Sub


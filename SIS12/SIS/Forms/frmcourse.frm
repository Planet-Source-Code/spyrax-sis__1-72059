VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmcourse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD COURSE"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   3405
   Icon            =   "frmcourse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   3405
   Begin MSComctlLib.ImageList MyImage 
      Left            =   2370
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcourse.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcourse.frx":0EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcourse.frx":165E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcourse.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdclose 
      Height          =   465
      Left            =   2340
      TabIndex        =   1
      Top             =   90
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "&Close"
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
      MICON           =   "frmcourse.frx":2552
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
   Begin VB.ListBox lstCourse 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3870
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   2235
   End
   Begin LVbuttons.LaVolpeButton cmdedit 
      Height          =   465
      Left            =   2340
      TabIndex        =   3
      Top             =   1050
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "&Edit"
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
      MICON           =   "frmcourse.frx":256E
      ALIGN           =   1
      IMGLST          =   "MyImage"
      IMGICON         =   "2"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmddelete 
      Height          =   465
      Left            =   2340
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "&Delete"
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
      MICON           =   "frmcourse.frx":258A
      ALIGN           =   1
      IMGLST          =   "MyImage"
      IMGICON         =   "3"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdadd 
      Height          =   465
      Left            =   2340
      TabIndex        =   2
      Top             =   570
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "&Add"
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
      MICON           =   "frmcourse.frx":25A6
      ALIGN           =   1
      IMGLST          =   "MyImage"
      IMGICON         =   "4"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
End
Attribute VB_Name = "frmcourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
frmAddCourse.Show
End Sub

Private Sub cmdclose_Click()
Unload Me

End Sub

Private Sub cmddelete_Click()

   Set MyRsCourse = New ADODB.Recordset
   MyRsCourse.Open "tblcourse", libcon, adOpenKeyset, adLockOptimistic
   
   If lstCourse.Text = "" Then
        MsgBox "Please select COURSE from the list to delete.", vbExclamation, "Warning!"
        Exit Sub
   Else
            With MyRsCourse
                .Filter = "course='" & lstCourse.Text & "'"
                .Delete
                .Requery
                .Close
            End With
            lstCourse.Clear
            MyListRefresh
            If lstCourse.Text = "" Then
               cmdedit.Visible = False
               cmddelete.Visible = False
            End If
            
    End If
End Sub

Private Sub Form_Load()
MyDBcon
FormCenter Me
MyListRefresh
End Sub

Private Sub lstCourse_Click()

If lstCourse.Text <> Empty Then
    cmddelete.Visible = True
    cmdedit.Visible = True
End If

End Sub

Private Sub MyListRefresh()
Set MyRsCourse = New ADODB.Recordset
MyRsCourse.Open "tblcourse", libcon, adOpenKeyset, adLockReadOnly
While MyRsCourse.EOF <> True
    lstCourse.AddItem MyRsCourse!course
    MyRsCourse.MoveNext
Wend
End Sub

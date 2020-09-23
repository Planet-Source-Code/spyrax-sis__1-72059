VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmAddCourse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD NEW COURSE"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4470
   Icon            =   "frmAddCourse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4470
   Begin LVbuttons.LaVolpeButton cmdsave 
      Height          =   435
      Left            =   480
      TabIndex        =   1
      Top             =   930
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
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
      MICON           =   "frmAddCourse.frx":076A
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox txtcourse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   480
      TabIndex        =   0
      Top             =   420
      Width           =   3525
   End
   Begin LVbuttons.LaVolpeButton cmdclose 
      Height          =   435
      Left            =   2520
      TabIndex        =   2
      Top             =   930
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
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
      MICON           =   "frmAddCourse.frx":0786
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label lblcourse 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter course:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   480
      TabIndex        =   3
      Top             =   150
      Width           =   2565
   End
End
Attribute VB_Name = "frmAddCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
Set MyRsCourse = New ADODB.Recordset

SqlStr = "select * from Tblcourse where course='" & txtcourse.Text & "'"
MyRsCourse.Open SqlStr, libcon, adOpenKeyset, adLockOptimistic
   

    If Not MyRsCourse.EOF And Not MyRsCourse.BOF Then
        MsgBox "Course number already exist!", vbExclamation, "SIS"
        txtcourse.Text = ""
        txtcourse.SetFocus
    Else
        With MyRsCourse
            .AddNew
            .Fields("course") = txtcourse.Text
            .Update
            .Close
        End With
    End If
    
    Set MyRsCourse = New ADODB.Recordset
    SqlStr = "select course from Tblcourse"
    MyRsCourse.Open SqlStr, libcon, adOpenKeyset, adLockOptimistic
    frmcourse.lstCourse.Clear
    While MyRsCourse.EOF <> True
        frmcourse.lstCourse.AddItem MyRsCourse!course
        MyRsCourse.MoveNext
    Wend
    frmcourse.lstCourse.Refresh
    Unload Me
   
   
End Sub

Private Sub Form_Load()
MyDBcon
FormCenter Me
End Sub

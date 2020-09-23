VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4680
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   Begin VB.Frame fra1 
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   4575
      Begin LVbuttons.LaVolpeButton cmdsearch 
         Height          =   405
         Left            =   390
         TabIndex        =   4
         Top             =   1350
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "&Search"
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
         MICON           =   "frmSearch.frx":076A
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
      Begin VB.TextBox txtsearch 
         Alignment       =   2  'Center
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
         Height          =   435
         Left            =   390
         TabIndex        =   3
         Text            =   "Enter String to search for:"
         Top             =   690
         Width           =   3855
      End
      Begin VB.OptionButton optfname 
         Caption         =   "FAMILY NAME"
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
         Height          =   465
         Left            =   2610
         TabIndex        =   2
         Top             =   150
         Width           =   1425
      End
      Begin VB.OptionButton optId 
         Caption         =   "ID NUMBER"
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
         Height          =   465
         Left            =   690
         TabIndex        =   1
         Top             =   150
         Width           =   1155
      End
      Begin LVbuttons.LaVolpeButton cmdCancel 
         Height          =   405
         Left            =   2790
         TabIndex        =   5
         Top             =   1380
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   714
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
         MICON           =   "frmSearch.frx":0786
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
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdsearch_Click()
    Set MyRsSearch = New ADODB.Recordset
    libcon.CursorLocation = adUseClient
    
If optId.Value = True Then
    SqlStr = "Select * from TblProf where id = " & txtsearch.Text & ""
    MyRsSearch.Open SqlStr, libcon, adOpenKeyset, adLockOptimistic
    If MyRsSearch.EOF Then
      MsgBox "Record not found!", vbInformation, "SIS"
      txtsearch.Text = ""
      txtsearch.SetFocus
    Else
        Unload Me
        Set frmView.dgridview.DataSource = MyRsSearch
        frmView.Show
        
    End If
End If
If optfname.Value = True Then
    SqlStr = "Select * from TblProf where fname = " & UCase$(txtsearch.Text) & ""
    MyRsSearch.Open SqlStr, libcon, adOpenKeyset, adLockOptimistic
    If MyRsSearch.EOF Then
      MsgBox "Record not found!", vbInformation, "SIS"
      txtsearch.Text = ""
      txtsearch.SetFocus
    Else
    Unload Me
    Set frmView.dgridview.DataSource = MyRsSearch
    frmView.Show
    
    End If
End If

    
End Sub

Private Sub Form_Load()
MyDBcon
End Sub

Private Sub txtsearch_Click()
SendKeys "{home}+{End}"
End Sub

Private Sub txtsearch_GotFocus()
SendKeys "{home}+{End}"
End Sub

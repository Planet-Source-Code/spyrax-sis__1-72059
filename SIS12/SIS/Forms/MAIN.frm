VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MAIN 
   BackColor       =   &H8000000C&
   Caption         =   "Student Information System ver 1.0"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "MAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   2700
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   5477
            Text            =   "Student Information System ver 1.0"
            TextSave        =   "Student Information System ver 1.0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   970
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   979
            MinWidth        =   970
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "4/20/2009"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList MyImage 
      Left            =   1650
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":0EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":165E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":342C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar MyToolbar 
      Align           =   1  'Align Top
      Height          =   2070
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   3651
      ButtonWidth     =   2223
      ButtonHeight    =   1164
      Appearance      =   1
      ImageList       =   "MyImage"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Student Profile"
            Object.ToolTipText     =   "Student Profile"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User Accounts"
            Object.ToolTipText     =   "User Accounts"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "System Lock"
            Object.ToolTipText     =   "System Lock"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            Object.ToolTipText     =   "Reports"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "7.1"
                  Text            =   "View All Records"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "7.2"
                  Text            =   "Print Records"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "System Settings"
            Object.ToolTipText     =   "System Settings"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "8.1"
                  Text            =   "Add Course"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "8.2"
                  Text            =   "Add College"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "8.3"
                  Text            =   "SY Settings"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnustudprof 
         Caption         =   "Student Profile"
      End
      Begin VB.Menu mnuuser 
         Caption         =   "User Accounts"
      End
      Begin VB.Menu mnusearch 
         Caption         =   "Search Record"
      End
      Begin VB.Menu break 
         Caption         =   "-"
      End
      Begin VB.Menu mnulock 
         Caption         =   "Lock System"
      End
      Begin VB.Menu break1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuviewAll 
         Caption         =   "View all Records"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print Records"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&System Settings"
      Begin VB.Menu mnuCourse 
         Caption         =   "Add Course"
      End
      Begin VB.Menu mnucollege 
         Caption         =   "Add College"
      End
      Begin VB.Menu mnuSY 
         Caption         =   "SY Settings"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuDev 
         Caption         =   "System Developer"
      End
   End
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'programmer: cadbisquera
'address: nueva vizcaya
'this application is for beginners that  i want to share as a basis on how to
'write an application in VB using MySql as the database


Private Sub MDIForm_Load()

Me.BackColor = RGB(153, 153, 204)
Me.Picture = LoadPicture(App.Path & "\MyBckgrd.jpg")
Me.Show
frmlogin.Show 1
End Sub

Private Sub mnuCourse_Click()
frmcourse.Show
End Sub

Private Sub mnuexit_Click()
Dim ANS As String
ANS = MsgBox("Do you really want to exit? ", vbInformation + vbYesNo, "SIS")
If ANS = vbYes Then
    End
Else
    Exit Sub
End If

End Sub

Private Sub mnulock_Click()
frmLock.Show 1
End Sub

Private Sub mnusearch_Click()
frmSearch.Show
End Sub

Private Sub mnustudprof_Click()
frmProfile.Show
End Sub

Private Sub mnuuser_Click()
frmAdminSetup.Show
End Sub

Private Sub mnuviewAll_Click()
frmView.Show
End Sub

Private Sub MyToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
          frmProfile.Show
    Case 2
          'show your form to view user accounts
    Case 3
          frmSearch.Show
    Case 5
          frmLock.Show 1
End Select

End Sub

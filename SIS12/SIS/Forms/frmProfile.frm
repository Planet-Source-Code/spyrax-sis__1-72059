VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Profile"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   Icon            =   "frmProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   9360
   Begin MSComctlLib.ImageList MyImage 
      Left            =   8160
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
            Picture         =   "frmProfile.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProfile.frx":0EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProfile.frx":165E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProfile.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton cmdclose 
      Height          =   435
      Left            =   7980
      TabIndex        =   46
      Top             =   420
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmProfile.frx":2CB2
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3795
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   6694
      _Version        =   393216
      TabHeight       =   520
      Enabled         =   0   'False
      BackColor       =   -2147483632
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Personal"
      TabPicture(0)   =   "frmProfile.frx":2CCE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblprof"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Contact"
      TabPicture(1)   =   "frmProfile.frx":2CEA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Emergency Contact"
      TabPicture(2)   =   "frmProfile.frx":2D06
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   -74910
         TabIndex        =   42
         Top             =   480
         Width           =   7725
         Begin VB.TextBox txtcontact 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   1740
            TabIndex        =   3
            Top             =   1320
            Width           =   3105
         End
         Begin VB.TextBox txtrelation 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   1740
            TabIndex        =   2
            Top             =   810
            Width           =   5445
         End
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   1740
            TabIndex        =   1
            Top             =   300
            Width           =   5445
         End
         Begin VB.Label lblcontact 
            BackStyle       =   0  'Transparent
            Caption         =   "CONTACT NUMBER"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   150
            TabIndex        =   45
            Top             =   1320
            Width           =   1605
         End
         Begin VB.Label lblrelation 
            BackStyle       =   0  'Transparent
            Caption         =   "RELATIONSHIP"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   510
            TabIndex        =   44
            Top             =   810
            Width           =   1215
         End
         Begin VB.Label lblname 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NAME"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1080
            TabIndex        =   43
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3135
         Left            =   -74910
         TabIndex        =   34
         Top             =   480
         Width           =   7725
         Begin VB.TextBox txtemail 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   1830
            TabIndex        =   10
            Top             =   2220
            Width           =   4605
         End
         Begin VB.TextBox txthphone 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   4710
            TabIndex        =   9
            Top             =   1410
            Width           =   2925
         End
         Begin VB.TextBox txtzcode 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   810
            TabIndex        =   8
            Top             =   1380
            Width           =   2595
         End
         Begin VB.TextBox txtprov 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   4530
            TabIndex        =   7
            Top             =   840
            Width           =   3105
         End
         Begin VB.TextBox txttown 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   810
            TabIndex        =   6
            Top             =   840
            Width           =   2595
         End
         Begin VB.TextBox txtbrgy 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   4530
            TabIndex        =   5
            Top             =   300
            Width           =   3105
         End
         Begin VB.TextBox txtstreet 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   810
            TabIndex        =   4
            Top             =   300
            Width           =   2595
         End
         Begin VB.Label lblemail 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "EMAIL ADD."
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   360
            TabIndex        =   41
            Top             =   2220
            Width           =   1395
         End
         Begin VB.Label lblHphone 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "HOME PHONE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   3540
            TabIndex        =   40
            Top             =   1410
            Width           =   1155
         End
         Begin VB.Label lblzcode 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ZIPCODE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   60
            TabIndex        =   39
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label lblprov 
            BackStyle       =   0  'Transparent
            Caption         =   "PROVINCE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   3570
            TabIndex        =   38
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label lbltown 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TOWN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   210
            TabIndex        =   37
            Top             =   840
            Width           =   555
         End
         Begin VB.Label lblbrgy 
            BackStyle       =   0  'Transparent
            Caption         =   "BARANGAY"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   3570
            TabIndex        =   36
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lblstreet 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "STREET"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   150
            TabIndex        =   35
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame fra1 
         Height          =   3135
         Left            =   90
         TabIndex        =   22
         Top             =   480
         Width           =   7725
         Begin VB.TextBox txtid 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   1050
            TabIndex        =   11
            Top             =   300
            Width           =   1695
         End
         Begin VB.ComboBox cmbcourse 
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   390
            Left            =   3630
            TabIndex        =   12
            Text            =   "COURSE"
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox txtyrsec 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   5940
            TabIndex        =   13
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txtfname 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   90
            TabIndex        =   14
            Top             =   960
            Width           =   2505
         End
         Begin VB.TextBox txtgname 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   2610
            TabIndex        =   15
            Top             =   960
            Width           =   2505
         End
         Begin VB.TextBox txtmname 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   5130
            TabIndex        =   16
            Top             =   960
            Width           =   2505
         End
         Begin VB.TextBox txtgender 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   1065
            TabIndex        =   17
            Top             =   1830
            Width           =   1695
         End
         Begin VB.TextBox txtcstatus 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   3960
            TabIndex        =   18
            Top             =   1830
            Width           =   1695
         End
         Begin VB.TextBox txtbdate 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   6315
            TabIndex        =   19
            Top             =   1830
            Width           =   1095
         End
         Begin VB.TextBox txtnationality 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   1560
            TabIndex        =   20
            Top             =   2430
            Width           =   2385
         End
         Begin VB.TextBox txttribe 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   4650
            TabIndex        =   21
            Top             =   2430
            Width           =   2775
         End
         Begin VB.Label lblid 
            BackStyle       =   0  'Transparent
            Caption         =   "ID NUMBER"
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
            Height          =   285
            Left            =   90
            TabIndex        =   33
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lblcourse 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "COURSE"
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
            Height          =   285
            Left            =   2820
            TabIndex        =   32
            Top             =   300
            Width           =   765
         End
         Begin VB.Label lbyrsec 
            BackStyle       =   0  'Transparent
            Caption         =   "YEAR/SEC."
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
            Height          =   285
            Left            =   4980
            TabIndex        =   31
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lblfname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "FAMILY NAME"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   90
            TabIndex        =   30
            Top             =   1380
            Width           =   2475
         End
         Begin VB.Label lblgname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "GIVEN NAME"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   2610
            TabIndex        =   29
            Top             =   1380
            Width           =   2475
         End
         Begin VB.Label lblmname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "MIDDLE NAME"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   5130
            TabIndex        =   28
            Top             =   1380
            Width           =   2475
         End
         Begin VB.Label lblgender 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "GENDER"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   285
            TabIndex        =   27
            Top             =   1860
            Width           =   735
         End
         Begin VB.Label lblcstatus 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CIVIL STATUS"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2805
            TabIndex        =   26
            Top             =   1830
            Width           =   1125
         End
         Begin VB.Label lblbdate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "BDATE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   5685
            TabIndex        =   25
            Top             =   1830
            Width           =   585
         End
         Begin VB.Label lblnationality 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NATIONALITY"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   2430
            Width           =   1245
         End
         Begin VB.Label lbltribe 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TRIBE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   3930
            TabIndex        =   23
            Top             =   2430
            Width           =   675
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Emergency Contact"
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
         Height          =   255
         Left            =   -69210
         TabIndex        =   53
         Top             =   30
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
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
         Height          =   255
         Left            =   -71370
         TabIndex        =   52
         Top             =   60
         Width           =   765
      End
      Begin VB.Label lblprof 
         BackStyle       =   0  'Transparent
         Caption         =   "Personal"
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
         Height          =   255
         Left            =   960
         TabIndex        =   51
         Top             =   60
         Width           =   765
      End
   End
   Begin LVbuttons.LaVolpeButton cmdNew 
      Height          =   435
      Left            =   7980
      TabIndex        =   47
      Top             =   870
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "&New"
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
      MICON           =   "frmProfile.frx":2D22
      ALIGN           =   1
      IMGLST          =   "MyImage"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton cmdedit 
      Height          =   435
      Left            =   7980
      TabIndex        =   48
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
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
      MICON           =   "frmProfile.frx":2D3E
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
   Begin LVbuttons.LaVolpeButton cmdsave 
      Height          =   435
      Left            =   7980
      TabIndex        =   49
      Top             =   1770
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmProfile.frx":2D5A
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
   Begin LVbuttons.LaVolpeButton cmdsearch 
      Height          =   435
      Left            =   7980
      TabIndex        =   50
      Top             =   2220
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
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
      FCOL            =   192
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmProfile.frx":2D76
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
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Dim ANS As String
If cmdclose.Caption = "&Close" Then
    Unload Me
Else
    ANS = MsgBox("Do you really want to close this form? ", vbInformation + vbYesNo, "SIS")
    If ANS = vbYes Then
        Unload Me
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub cmdedit_Click()
cmdsave.Visible = True
cmdedit.Visible = False
SSTab1.Enabled = True
txtid.SetFocus
End Sub

Private Sub cmdNew_Click()
SSTab1.Enabled = True
txtid.SetFocus
cmdNew.Enabled = False
cmdsearch.Enabled = False
cmdsave.Visible = True
cmdclose.Caption = "&Cancel"
End Sub

Private Sub cmdsave_Click()
Set MyRsProf = New ADODB.Recordset

SqlStr = "select * from Tblprof where id='" & txtid.Text & "'"
MyRsProf.Open SqlStr, libcon, adOpenKeyset, adLockOptimistic
   

    If Not MyRsProf.EOF And Not MyRsProf.BOF Then
        MsgBox "Student ID number already exist!", vbExclamation, "SIS"
        txtid.Text = ""
        txtid.SetFocus
        Exit Sub
    Else

    With MyRsProf
    .AddNew
    .Fields("id") = txtid.Text
    .Fields("course") = cmbcourse.Text
    .Fields("yrsec") = txtyrsec.Text
    .Fields("fname") = txtFname.Text
    .Fields("gname") = txtgname.Text
    .Fields("mname") = txtMname.Text
    .Fields("gender") = txtgender.Text
    .Fields("cstatus") = txtcstatus.Text
    .Fields("bdate") = txtbdate.Text
    .Fields("nationality") = txtnationality.Text
    .Fields("tribe") = txttribe.Text
    .Fields("street") = txtstreet.Text
    .Fields("brgy") = txtbrgy.Text
    .Fields("town") = txttown.Text
    .Fields("prov") = txtprov.Text
    .Fields("zcode") = txtzcode.Text
    .Fields("hphone") = txthphone.Text
    .Fields("email") = txtemail.Text
    .Fields("name") = txtname.Text
    .Fields("relation") = txtrelation.Text
    .Fields("contact") = txtContact.Text
    .Update
    .Close
    End With
    End If
    MsgBox "Record has been save..", vbInformation, "SIS"
End Sub

Private Sub cmdsearch_Click()
cmdsearch.Visible = False
cmdNew.Visible = False
cmdedit.Visible = True
End Sub

Private Sub Form_Load()
MyDBcon
FormCenter Me
Set MyRsCourse = New ADODB.Recordset
SqlStr = "Select * from tblcourse"
MyRsCourse.Open SqlStr, libcon, adOpenKeyset, adLockOptimistic
While MyRsCourse.EOF <> True
    cmbcourse.AddItem MyRsCourse!course
    MyRsCourse.MoveNext
Wend

End Sub

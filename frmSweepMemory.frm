VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSweepMemory 
   Caption         =   "Sweep The System..."
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   Icon            =   "frmSweepMemory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3060
      Top             =   4440
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4140
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSweepMemory.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSweepMemory.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSweepMemory.frx":079E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSweepMemory.frx":0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSweepMemory.frx":1052
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   573
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Operations"
      TabPicture(0)   =   "frmSweepMemory.frx":8554
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Line2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvwFilesAndFolders"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Settin&gs"
      TabPicture(1)   =   "frmSweepMemory.frx":8570
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Miscelleno&us"
      TabPicture(2)   =   "frmSweepMemory.frx":858C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   "Internet and Background Activities..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4875
         Left            =   4440
         TabIndex        =   41
         Top             =   540
         Width           =   5175
         Begin VB.CommandButton cmdGrtInformation 
            Caption         =   "G&et Current Information"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2820
            TabIndex        =   42
            Top             =   4380
            Width           =   2235
         End
         Begin VB.Label lblInternetConnectionState 
            AutoSize        =   -1  'True
            Caption         =   "Network Connection"
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
            Left            =   2460
            TabIndex        =   46
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Internet Connection State"
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
            Left            =   240
            TabIndex        =   45
            Top             =   720
            Width           =   1890
         End
         Begin VB.Label lblNetworkConnection 
            AutoSize        =   -1  'True
            Caption         =   "Network Connection"
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
            Left            =   2460
            TabIndex        =   44
            Top             =   420
            Width           =   1455
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Network Connection"
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
            Left            =   240
            TabIndex        =   43
            Top             =   420
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4755
         Left            =   -74820
         TabIndex        =   33
         Top             =   660
         Width           =   4335
         Begin VB.CheckBox chkSettings 
            Caption         =   "Empty Recycle Bin On Application Start Up"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   37
            Top             =   1620
            Width           =   3915
         End
         Begin VB.CheckBox chkSettings 
            Caption         =   "Empty Temp Folder on  Application Start Up"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   36
            Top             =   1260
            Width           =   3435
         End
         Begin VB.CheckBox chkSettings 
            Caption         =   "Add Application to Start Up"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   35
            Top             =   900
            Width           =   2295
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Apply"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3000
            TabIndex        =   34
            Top             =   4200
            Width           =   1155
         End
         Begin VB.Label Label5 
            Caption         =   "It is recommended that you restart your machine after applying settings"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   495
            Left            =   180
            TabIndex        =   38
            Top             =   420
            Width           =   3570
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Sweep Options ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3435
         Left            =   -74820
         TabIndex        =   13
         Top             =   2220
         Width           =   6015
         Begin VB.CommandButton cmdListFiles 
            Height          =   315
            Index           =   3
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   2580
            Width           =   375
         End
         Begin VB.CheckBox chkSweepOptions 
            Caption         =   "Temporary Internet Explorer Files."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   29
            Top             =   2580
            Width           =   2955
         End
         Begin VB.CommandButton cmdListFiles 
            Height          =   315
            Index           =   2
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2160
            Width           =   375
         End
         Begin VB.CheckBox chkSweepOptions 
            Caption         =   "Empty Clip Board."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   27
            Top             =   2160
            Width           =   2355
         End
         Begin VB.CheckBox chkSweepOptions 
            Caption         =   "Temporary Files in Temp Folder."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   600
            Width           =   2715
         End
         Begin VB.CheckBox chkSweepOptions 
            Caption         =   "Delete Folders and Sub Folders in."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   25
            Top             =   960
            Width           =   2955
         End
         Begin VB.TextBox txtFolderPath 
            Height          =   315
            Left            =   180
            TabIndex        =   24
            Top             =   1320
            Width           =   5055
         End
         Begin VB.CommandButton cmdBrowseFolder 
            Height          =   315
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdListFiles 
            Height          =   315
            Index           =   0
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   540
            Width           =   375
         End
         Begin VB.CommandButton cmdListFiles 
            Height          =   315
            Index           =   1
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   900
            Width           =   375
         End
         Begin VB.CheckBox chkSweepOptions 
            Caption         =   "File having the extension as "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   20
            Top             =   1740
            Width           =   2355
         End
         Begin VB.TextBox txtExtension 
            Height          =   315
            Left            =   2640
            TabIndex        =   19
            Text            =   "*.tmp"
            Top             =   1740
            Width           =   975
         End
         Begin VB.CommandButton cmdSweep 
            Caption         =   "&Sweep"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4860
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2940
            Width           =   1035
         End
         Begin VB.CheckBox chkSweepOptions 
            Caption         =   "Empty Recycle Bin."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   17
            Top             =   2940
            Width           =   2415
         End
         Begin VB.CommandButton cmdListFiles 
            Height          =   315
            Index           =   4
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2940
            Width           =   375
         End
         Begin VB.CommandButton cmdListFiles 
            Height          =   315
            Index           =   5
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1740
            Width           =   375
         End
         Begin VB.DriveListBox Drive1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3780
            TabIndex        =   14
            Top             =   1740
            Width           =   915
         End
         Begin VB.Label Label6 
            Caption         =   "Drive"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   32
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label lblMessage 
            AutoSize        =   -1  'True
            Caption         =   "Warning:Permanent Deletion"
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
            Left            =   180
            TabIndex        =   31
            Top             =   300
            Width           =   2070
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "System Information..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1515
         Left            =   -74820
         TabIndex        =   6
         Top             =   720
         Width           =   6015
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Computer Name"
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
            Left            =   180
            TabIndex        =   12
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label lblCompName 
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   2100
            TabIndex        =   11
            Top             =   300
            Width           =   90
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "User Name"
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
            Left            =   180
            TabIndex        =   10
            Top             =   600
            Width           =   780
         End
         Begin VB.Label lblUserName 
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   2100
            TabIndex        =   9
            Top             =   600
            Width           =   90
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Temporary Files Folder"
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
            Left            =   180
            TabIndex        =   8
            Top             =   900
            Width           =   1635
         End
         Begin VB.Label lblTempFolderPath 
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   1140
            Width           =   90
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "Miscellenous API Demos..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4875
         Left            =   240
         TabIndex        =   1
         Top             =   540
         Width           =   3975
         Begin VB.CommandButton cmdShowDesktop 
            Caption         =   "Sho&w Desktop"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   4
            Top             =   660
            Width           =   3555
         End
         Begin VB.TextBox txtNewComputerName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            MaxLength       =   31
            TabIndex        =   3
            Top             =   1620
            Width           =   3555
         End
         Begin VB.CommandButton cmdChangeCompName 
            Caption         =   "Chan&ge Computer Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   2
            Top             =   2100
            Width           =   3555
         End
         Begin VB.Label Label7 
            Caption         =   "Enter &New name for your computer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   1260
            Width           =   3495
         End
      End
      Begin MSComctlLib.ListView lvwFilesAndFolders 
         Height          =   4995
         Left            =   -68700
         TabIndex        =   39
         Top             =   660
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   8811
         SortKey         =   1
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   882
         EndProperty
      End
      Begin VB.Label lblOSName 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3960
         TabIndex        =   47
         Top             =   -360
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "File Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -67980
         TabIndex        =   40
         Top             =   420
         Width           =   1350
      End
      Begin VB.Line Line1 
         X1              =   -68820
         X2              =   -68820
         Y1              =   720
         Y2              =   2100
      End
      Begin VB.Line Line2 
         X1              =   -68880
         X2              =   -68880
         Y1              =   2280
         Y2              =   5340
      End
   End
   Begin VB.Label lblOSVersion 
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3720
      TabIndex        =   48
      Top             =   180
      Width           =   105
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMaximise 
         Caption         =   "Maximise"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
      Begin VB.Menu mnuMinimise 
         Caption         =   "Minimize"
      End
   End
End
Attribute VB_Name = "frmSweepMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer : Malik Iftikhar Hussain
'Email: humsafar_ak@yahoo.com

Option Explicit
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long

Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2

Const WM_LBUTTONDOWN As Long = &H201
Const WM_LBUTTONDBLCLK As Long = &H203
Const WM_RBUTTONDBLCLK As Long = &H206
Const WM_RBUTTONDOWN As Long = &H204


Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4

'shell notification icon

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'to apply this settings in ini file
'if startup is selected we need to put our app in registry

Private Sub cmdApply_Click()
On Error GoTo err

    'application at OS Startup
    If chkSettings(0).Value = 1 Then
        WritePrivateProfileString "Application", "Startup", "True", App.Path & "\SweepMemory.ini"
        RegisterForStartUp "Register"
    Else
        WritePrivateProfileString "Application", "Startup", "False", App.Path & "\SweepMemory.ini"
        RegisterForStartUp "Unregister"
    End If
    
    'application Temp
    If chkSettings(1).Value = 1 Then
        WritePrivateProfileString "Application", "Temp", "True", App.Path & "\SweepMemory.ini"
    Else
        WritePrivateProfileString "Application", "Temp", "False", App.Path & "\SweepMemory.ini"
    End If
    
    'application Recycle
    If chkSettings(2).Value = 1 Then
        WritePrivateProfileString "Application", "Recycle", "True", App.Path & "\SweepMemory.ini"
    Else
        WritePrivateProfileString "Application", "Recycle", "False", App.Path & "\SweepMemory.ini"
    End If
    
    MsgBox "Changes Applied." & vbCrLf & "Restart machine for apply changes to take effects. ", vbInformation
     ShutDown
    Exit Sub
    
err:
    MsgBox err.Description & ".Error while applying settings", vbInformation
End Sub

Private Sub cmdBrowseFolder_Click()
    Dim uBrowseFolder As BrowseInfo
    Dim sTemp         As String
    Dim lResult       As Long
    
    
    'set the values of ubroswe
    With uBrowseFolder
        .hWndOwner = Me.hwnd
        .lpszTitle = lstrcat("C:\", "")
        .ulFlags = 1  'for directory
        '.pszDisplayName = "Select Folder..."
    End With

    lResult = SHBrowseForFolder(uBrowseFolder)
    
    'if some value is returned then
    'lresult will have no null value
    
    If lResult Then
        sTemp = Space$(255)
        SHGetPathFromIDList lResult, sTemp
        
        'clear memory
        CoTaskMemFree lResult
        
        'now get path
        sTemp = Mid(sTemp, 1, InStr(1, sTemp, Chr$(0)) - 1)
        If Not IsNull(sTemp) Then
            txtFolderPath.Text = sTemp
        End If
    End If

End Sub

Private Sub cmdChangeCompName_Click()
    Dim sCompName As String
    
    sCompName = Trim(txtNewComputerName.Text)
    
    
    On Error GoTo err
    
        If SetComputerName(sCompName) = 0 Then
              MsgBox "Unable to change the name.", vbInformation
        Else
            MsgBox "Machine name changed.Requires Restart", vbInformation
            ShutDown
        End If
    
    Exit Sub
err:
    Debug.Print err.LastDllError
End Sub

Private Sub cmdGrtInformation_Click()
    InterNetNetwork
End Sub

Private Sub cmdListFiles_Click(Index As Integer)
    Select Case Index
        Case 0   'temperory folder
            MsgBox "This Folder is a System Folder." & vbCrLf & "All files in this folder will be listed in the Accompaning list.", vbInformation
            FindFiles lblTempFolderPath.Caption
        Case 1   'Selected folder
            MsgBox "Delete files in a selected folder." & vbCrLf & "To do this just click Browse button at the end" & vbCrLf & "and select the required folder.", vbInformation
            FindFiles txtFolderPath
        Case 2
            MsgBox "Empty the data in clipboard.", vbInformation
            GetClipData
        Case 3 'ie temp folder
            
            FindUrls
        Case 4
            MsgBox "Empty Recycle Bin of all Drives", vbInformation
            
        Case 5
            MsgBox "Deletes file on specified extension from the selected drive", vbInformation
            FindFiles CStr(Drive1.Drive), Trim(txtExtension.Text)
            
    End Select
End Sub

Private Sub cmdShowDesktop_Click()
    Const SW_SHOWNORMAL As Long = 1
    Dim lDeskTopHandle As Long
    
    lDeskTopHandle = GetDesktopWindow
    
    Call ShowWindow(lDeskTopHandle, SW_SHOWNORMAL)
    

End Sub

Private Sub cmdSweep_Click()
    Dim iCnt        As Integer
    Dim bSelection  As Boolean
    
    'check for some selection
    For iCnt = chkSweepOptions.LBound To chkSweepOptions.UBound
        'if some selection exists
        If chkSweepOptions(iCnt).Value = 1 Then
            bSelection = True
            Exit For
        Else
            bSelection = False
        End If
    Next
    
    If Not bSelection Then
        MsgBox "No selection has been done.Please make a selection", vbInformation
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'if selection is done
    If MsgBox("Are you sure to delete all the files?", vbYesNo) = vbNo Then
        Exit Sub
    Else
        bSweep = True
        If chkSweepOptions(0).Value = 1 Then
            FindFiles lblTempFolderPath.Caption
        End If
        If chkSweepOptions(1).Value = 1 Then
            FindFiles txtFolderPath.Text
        End If
        If chkSweepOptions(2).Value = 1 Then
            FindFiles CStr(Drive1.Drive), Trim(txtExtension.Text)
        End If
        If chkSweepOptions(3).Value = 1 Then
            EmptyClip
        End If
        If chkSweepOptions(5).Value = 1 Then
            EmptyReCycleBin
        End If
    End If
    bSweep = False
    
    MsgBox "Removal complete." & vbCrLf & "It is possible that some of the files may not be deleted." & vbCrLf & "Please click the Help button to see remaining files", vbInformation
    
    
End Sub



Private Sub Form_Load()
cmdBrowseFolder.Picture = ImageList1.ListImages(2).Picture
    Dim i As Integer
    For i = 0 To cmdListFiles.Count - 1
        cmdListFiles(i).Picture = ImageList1.ListImages(1).Picture
    Next
    
    'adding Icon to taskbar
    lblMessage.ForeColor = vbRed
    AddIconToTaskar "Register"
    
    If chkSettings(2).Value = 1 Then
        EmptyReCycleBin
    End If
    
End Sub


'this is used when the
'application is in the minimised state
'and mouse is over the icon
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim MSG
    MSG = X / Screen.TwipsPerPixelX
    
    If MSG = WM_LBUTTONDOWN Or MSG = WM_RBUTTONDOWN Then
        
    End If
    
    If MSG = WM_LBUTTONDBLCLK Then
        Me.PopupMenu mnuMain
    End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
    Me.Height = lHeight
    Me.Width = lWidth
End Sub
Private Sub GetClipData()
  Const CF_Text = 1
  Const CF_BITMap = 2
  Dim hStrPtr As Long
  Dim lLength As Long
  Dim sBuffer As String
  
  'check for text format
  'get the content of the clipboard
  'if text is prsent then add it to
  'thstext of the frmclip form
  'and display it to the user'
  
  
  If IsClipboardFormatAvailable(CF_Text) <> 0 Then
        If MsgBox("Clipboard consist of text format." & vbCrLf & "Do you want to see the contents?", vbYesNo) = vbNo Then
            Exit Sub
        Else
            OpenClipboard Me.hwnd
            hStrPtr = GetClipboardData(CF_Text)
            If hStrPtr <> 0 Then
                lLength = lstrlen(hStrPtr)
                If lLength > 0 Then
                    sBuffer = Space$(lLength)
                    CopyMemory ByVal sBuffer, ByVal hStrPtr, lLength
                    Debug.Print sBuffer
                    frmClipBoard.Text1.Text = frmClipBoard.Text1.Text & sBuffer & vbCrLf
                End If
            End If
            CloseClipboard
            frmClipBoard.Show vbModal, frmSweepMemory
            frmClipBoard.Text1.ZOrder 1
            
        End If
  'for bitmaps
  ElseIf IsClipboardFormatAvailable(CF_BITMap) <> 0 Then
        MsgBox "Clip board contains Bitmap or Picture File."
  Else
        MsgBox "Unable to determine Clipboard contents."
  End If
        
    
  


End Sub

'for making clipboard empty

Private Sub EmptyClip()
On Error GoTo err

    'open clipboard exclusively for this
    'application
    
    OpenClipboard Me.hwnd
    EmptyClipboard
    Exit Sub
err:
    MsgBox err.Number & " Dll Error " & err.LastDllError
    
    
End Sub

'for empting recycle bin
'this proc will empty recycle bin
'for all drives

Private Sub EmptyReCycleBin()
    Const SHERB_NOCONFIRMATION As Long = &H1
    
    'passing arguments
    ' the handle to the window
    ' vbnullstring for empting Recycle Bin of all drives
    ' offer no confirmation on delition
    SHEmptyRecycleBin Me.hwnd, vbNullString, SHERB_NOCONFIRMATION
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Timer1_Timer()
    
    
    If lblMessage.ForeColor = vbBlue Then
        lblMessage.ForeColor = vbRed
    Else
        lblMessage.ForeColor = vbBlue
    End If
    
End Sub

'proc to register and unregister a icon to taskbar
Public Sub AddIconToTaskar(sAction As String)




'initialise the Notify structure
    
    Dim NICON As NOTIFYICONDATA
    NICON.cbSize = Len(NICON)
    NICON.hwnd = Me.hwnd
    NICON.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    NICON.uID = 1&
    NICON.uCallbackMessage = WM_LBUTTONDOWN
    NICON.hIcon = Me.Icon

    Select Case sAction
        Case "Register"
            Shell_NotifyIcon NIM_ADD, NICON
        
        
        Case Else
            Shell_NotifyIcon NIM_DELETE, NICON
    End Select
End Sub

'to handle the messages
'when the application is closed
'but still in the notify icon
Public Sub HandleMessages()



End Sub

Private Sub Form_Unload(Cancel As Integer)
    AddIconToTaskar "Unregister"
End Sub

Private Sub mnuMaximise_Click()
    frmSweepMemory.Show
End Sub

Private Sub mnuMinimise_Click()
    frmSweepMemory.Hide
    mnuMaximise.Enabled = True
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

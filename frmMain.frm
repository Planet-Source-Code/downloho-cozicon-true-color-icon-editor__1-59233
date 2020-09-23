VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "untitled - CozIcon"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10890
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   726
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picExtra 
      Height          =   1335
      Left            =   0
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   25
      Top             =   1590
      Width           =   900
      Begin VB.PictureBox picToolSel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   900
         Index           =   10
         Left            =   240
         MouseIcon       =   "frmMain.frx":0CCC
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1996
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   300
         Begin VB.Shape shpToolSel 
            BorderColor     =   &H000000FF&
            Height          =   255
            Index           =   10
            Left            =   0
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.PictureBox picToolSel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   600
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmMain.frx":1A02
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":26CC
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   300
         Begin VB.Shape shpToolSel 
            BorderColor     =   &H000000FF&
            Height          =   255
            Index           =   1
            Left            =   0
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.PictureBox picToolSel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1200
         Index           =   5
         Left            =   240
         MouseIcon       =   "frmMain.frx":2779
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":3443
         ScaleHeight     =   80
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   300
         Begin VB.Shape shpToolSel 
            BorderColor     =   &H000000FF&
            Height          =   255
            Index           =   5
            Left            =   0
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.PictureBox picToolSel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1200
         Index           =   3
         Left            =   120
         MouseIcon       =   "frmMain.frx":361C
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":42E6
         ScaleHeight     =   80
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   26
         Top             =   8
         Visible         =   0   'False
         Width           =   600
         Begin VB.Shape shpToolSel 
            BorderColor     =   &H000000FF&
            Height          =   255
            Index           =   3
            Left            =   0
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
      End
   End
   Begin VB.PictureBox picUndo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   24
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   22
      Top             =   6960
      Visible         =   0   'False
      Width           =   480
   End
   Begin CozIcon.winConnect wc 
      Left            =   0
      Top             =   7320
      _ExtentX        =   1535
      _ExtentY        =   661
   End
   Begin VB.FileListBox flbPlugin 
      Height          =   285
      Left            =   240
      Pattern         =   "*.exe"
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   1080
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":462A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5032
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5536
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6442
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6946
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbTools 
      Height          =   390
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   688
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Select"
            Object.Tag             =   "11"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Text"
            Object.Tag             =   "12"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pen Tool"
            Object.Tag             =   "1"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Line Tool"
            Object.Tag             =   "10"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eye Dropper"
            Object.Tag             =   "9"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fill Color"
            Object.Tag             =   "2"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Box"
            Object.Tag             =   "3"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Circle"
            Object.Tag             =   "5"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1080
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   7
   End
   Begin VB.PictureBox picTrans 
      BackColor       =   &H00808000&
      Height          =   180
      Left            =   120
      MouseIcon       =   "frmMain.frx":6E4A
      MousePointer    =   99  'Custom
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   180
   End
   Begin VB.PictureBox picClr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   60
      MouseIcon       =   "frmMain.frx":7B14
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   6
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox picClr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   300
      MouseIcon       =   "frmMain.frx":87DE
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   7
      Top             =   3240
      Width           =   480
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00808080&
      Height          =   6780
      Left            =   960
      ScaleHeight     =   448
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   656
      TabIndex        =   3
      Top             =   0
      Width           =   9900
      Begin VB.OptionButton optEdit 
         Height          =   375
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5760
         Width           =   375
      End
      Begin VB.VScrollBar vsEdit 
         Height          =   615
         LargeChange     =   10
         Left            =   9600
         Max             =   0
         SmallChange     =   5
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin VB.HScrollBar hsEdit 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   0
         SmallChange     =   5
         TabIndex        =   29
         Top             =   6360
         Width           =   1815
      End
      Begin VB.PictureBox picEdit 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   5760
         Left            =   120
         MousePointer    =   2  'Cross
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   384
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   456
         TabIndex        =   4
         Top             =   120
         Width           =   6840
         Begin VB.Shape shpSel 
            BorderStyle     =   3  'Dot
            Height          =   735
            Left            =   1440
            Top             =   3360
            Visible         =   0   'False
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox picIconsBack 
      BackColor       =   &H00808080&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   0
      Top             =   3840
      Width           =   900
      Begin VB.VScrollBar vsIcons 
         Height          =   495
         Left            =   0
         SmallChange     =   34
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.PictureBox picIconsMove 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   90
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   45
         TabIndex        =   1
         Top             =   120
         Width           =   675
         Begin VB.PictureBox picIcon 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   90
            MouseIcon       =   "frmMain.frx":94A8
            MousePointer    =   99  'Custom
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   2
            Tag             =   "untitled"
            Top             =   0
            Width           =   480
         End
         Begin VB.Label txtIcon 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   420
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   540
            Width           =   675
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   960
      TabIndex        =   5
      Top             =   6840
      Width           =   9900
      Begin VB.PictureBox picClrBack 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   44
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   417
         TabIndex        =   17
         Top             =   160
         Width           =   6255
         Begin VB.PictureBox picClrSel 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   1
            Left            =   2400
            ScaleHeight     =   18
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   255
            TabIndex        =   20
            Top             =   0
            Width           =   3855
         End
         Begin VB.PictureBox picClrSel 
            AutoRedraw      =   -1  'True
            Height          =   615
            Index           =   0
            Left            =   0
            MouseIcon       =   "frmMain.frx":A172
            MousePointer    =   99  'Custom
            ScaleHeight     =   37
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   155
            TabIndex        =   19
            Top             =   0
            Width           =   2385
         End
         Begin VB.PictureBox picClrSel 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   2
            Left            =   2400
            ScaleHeight     =   18
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   255
            TabIndex        =   18
            Top             =   330
            Width           =   3855
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   680
         Left            =   6480
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   10
         Top             =   120
         Width           =   1455
         Begin VB.TextBox txtClr 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "000000"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtClr 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Height          =   285
            Index           =   2
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0"
            Top             =   60
            Width           =   495
         End
         Begin VB.TextBox txtClr 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Height          =   285
            Index           =   1
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0"
            Top             =   60
            Width           =   495
         End
         Begin VB.TextBox txtClr 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0"
            Top             =   60
            Width           =   495
         End
      End
   End
   Begin VB.Label lblSwitch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   540
      MouseIcon       =   "frmMain.frx":AE3C
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2970
      Width           =   300
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Begin VB.Menu mnuFileNew32 
            Caption         =   "32x32 Icon"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuFileNew16 
            Caption         =   "16x16 Icon"
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLinei43i566 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "Recent Files"
         Begin VB.Menu mnuFileRecentFiles 
            Caption         =   "(empty)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuLine39k3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuLine34445 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLinei821 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuLinej4049d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuLinek431 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewZoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "100%"
            Index           =   1
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "500%"
            Index           =   5
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "1000%"
            Index           =   10
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuViewZoomSize 
            Caption         =   "2000%"
            Checked         =   -1  'True
            Index           =   20
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu mnuLine45r4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewGrid 
         Caption         =   "Grid"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Begin VB.Menu mnuImageFlipHo 
         Caption         =   "Flip Horizontal"
      End
      Begin VB.Menu mnuImageFlipVert 
         Caption         =   "Flip Vertical"
      End
      Begin VB.Menu mnuLine33kk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageCopyNew 
         Caption         =   "Copy to New"
      End
      Begin VB.Menu mnuLinerirt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "Plugins"
      Begin VB.Menu mnuPluginsRun 
         Caption         =   "(empty)"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpTest 
         Caption         =   "Test Icon"
      End
      Begin VB.Menu mnuLinei392 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'

'Supports 16x16 and 32x32 Full Transparency True Color Icons. Has most of the basic image editor type settings: pen, fill, eclipse, rect, text, selections, etc.
'But also includes gradient rectangles and gradient eclipses. Also has mild support for custom brushes and my own type of plugins.
'Drag and Drop is supported as well is bmp, gif or jpeg to icon conversion. Can also load language packs for our non-english reading friends.

'Included is the Main Project, some sample Plugins Projects, some custom brush files, a read me file about custom brushes and a read me concerning the language packs(with two example packs: one for english and the other for spanish).
'Some subsets of the Main Project include dynamically loading custom language packs, winconnect(a custom activex for communication through different windows) and hard coding icons (not using savepicture via Visual Basic).


Option Explicit
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Enum UDE_TOOLS 'these are our tools
 tDraw = 1
 tFill = 2
 tBox = 3
 tBoxFill = 4
 tCircle = 5
 tCircleFill = 6
 tBoxGrad = 7
 tCircleGrad = 8
 tEye = 9
 tLine = 10
 tSelect = 11
 tText = 12
 tBoxGradNS = 13
 tBoxFillX = 14
 tCircleFillX = 15
 tcustom = 16
 tBoxGradNESW = 17
 tBoxGradNWSE = 18
 tLineX2 = 19
 tLineX3 = 20
End Enum

Private Enum UDE_GRADDIR 'something quick to refer to when painting Gradient Boxes
 gNS = 0
 gEW = 1
 gNESW = 2
 gNWSE = 3
End Enum

'x1 and y1 are original mousedown position, xz1 and yz1 are used for the selection tool
Dim X1 As Integer, xz1 As Integer
Dim Y1 As Integer, yz1 As Integer
Dim X2 As Integer 'mouseup postion
Dim Y2 As Integer 'mouseup postion
Dim mButton As Integer 'current button down
Dim SelTool As UDE_TOOLS 'current tool
Dim KeyDown As Integer 'current keydown
Dim mColor(4) As Long 'colors I originally planned on using the mouse wheel click so it's available
Dim mZoom As Integer 'current zoom (ie. IconWidth * mZoom)
Dim mDrawGrid As Boolean 'should we draw the grid?
'current icon being edited, Custom Tool Data,    Type of custom tool
Dim mCurrIcon As Integer, mCustomTool As String, tCustomType As Integer
'Are we moving the selection or drawing it?, Last Select position
Dim MovingSel As Boolean, OldSel(3) As Integer
'are we drawing anything?, stores wether an icon was edited.
Dim drawing As Boolean, FileChanged(99) As Boolean
Dim mOpacity As Integer

Private Sub AddRecent(ByVal s As String)
'adds 's' to the recent files menu, only 10 items are allowed in the menu
Dim x As Integer
 For x = 0 To mnuFileRecentFiles.Count - 1
   If s = mnuFileRecentFiles(x).Caption Then Exit Sub
 Next x

If mnuFileRecentFiles.Count >= 10 Then
 Dim arr(9) As String
 For x = mnuFileRecentFiles.Count - 1 To 0 Step -1
  arr(x) = mnuFileRecentFiles(x).Caption
 Next x
 For x = 0 To mnuFileRecentFiles.Count - 2
   mnuFileRecentFiles(x).Caption = arr(x + 1)
 Next x
  mnuFileRecentFiles(mnuFileRecentFiles.Count - 1).Caption = s
Else
 x = mnuFileRecentFiles.Count
 If x = 1 And mnuFileRecentFiles(0).Enabled = False Then
  mnuFileRecentFiles(0).Caption = s
  mnuFileRecentFiles(0).Visible = True
  mnuFileRecentFiles(0).Enabled = True
 Else
  Call Load(mnuFileRecentFiles(x))
  mnuFileRecentFiles(x).Caption = s
  mnuFileRecentFiles(x).Visible = True
  mnuFileRecentFiles(x).Enabled = True
 End If
End If
End Sub

Private Function BlendColor(ByVal Clr1 As Long, ByVal Clr2 As Long, ByVal Amount As Single) As Long
'used for pen opacity, not implemented
'mOpacity = 50
'IIf(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)) = picTrans.BackColor, mColor(mButton), BlendColor(picIcon(mCurrIcon).Point(Int(x1 / mZoom), Int(y1 / mZoom)), mColor(mButton), mOpacity))
If Amount = 100 Then BlendColor = Clr1: Exit Function
If Amount < 1 Then BlendColor = Clr2: Exit Function

Dim r(2) As Single, g(2) As Single, b(2) As Single

Amount = 100 / Amount

r(0) = GetRGB(Clr1).Red
r(1) = GetRGB(Clr2).Red
g(0) = GetRGB(Clr1).Green
g(1) = GetRGB(Clr2).Green
b(0) = GetRGB(Clr1).Blue
b(1) = GetRGB(Clr2).Blue

r(2) = (r(1) - r(0)) / Amount
g(2) = (g(1) - g(0)) / Amount
b(2) = (b(1) - b(0)) / Amount

r(0) = r(0) + r(2)
If r(0) < 0 Then r(0) = 0
If r(0) > 255 Then r(0) = 255

g(0) = g(0) + g(2)
If g(0) < 0 Then g(0) = 0
If g(0) > 255 Then g(0) = 255

b(0) = b(0) + b(2)
If b(0) < 0 Then b(0) = 0
If b(0) > 255 Then b(0) = 255

BlendColor = RGB(r(0), g(0), b(0))
If BlendColor = picTrans.BackColor Then BlendColor = BlendColor + 1
End Function

Private Sub DrawGrid()
'draws a grid on the Edit screen according to the currect zoom
If picEdit.Height = picIcon(mCurrIcon).Width Or mDrawGrid = False Then Exit Sub
Dim x As Integer
Dim y As Integer

For y = 0 To picIcon(mCurrIcon).Height
 For x = 0 To picIcon(mCurrIcon).Width
  picEdit.Line (Int(x * mZoom), Int(y * mZoom))-(Int(x * mZoom), picEdit.ScaleHeight), RGB(190, 190, 190)
  picEdit.Line (Int(x * mZoom), Int(y * mZoom))-(picEdit.ScaleWidth, Int(y * mZoom)), RGB(190, 190, 190)
 Next x
Next y

  picEdit.Refresh
  picEdit.Line (picEdit.ScaleWidth - 1, 0)-(picEdit.ScaleWidth - 1, picEdit.ScaleHeight), RGB(190, 190, 190)
  picEdit.Line (0, picEdit.ScaleHeight - 1)-(picEdit.ScaleWidth - 1, picEdit.ScaleHeight - 1), RGB(190, 190, 190)
  'picedit.Line (Int(X * mzoom), Int(Y * mzoom))-(picedit.ScaleWidth, Int(Y * mzoom)), RGB(190, 190, 190)
End Sub

Private Function ExtractFileName(ByVal File As String) As String
'returns a file's name sans Directory and Extension
Dim i As Integer
i = InStrRev(File, "\")
If i = 0 Then ExtractFileName = File: Exit Function
ExtractFileName = Mid(File, i + 1)
i = InStrRev(ExtractFileName, ".")
ExtractFileName = Left(ExtractFileName, i - 1)
End Function

Private Sub FillRegion(ByVal x As Integer, ByVal y As Integer)
'for paint bucket tool
  Dim a As Integer, b As Integer
  a = picIcon(mCurrIcon).FillStyle
  b = picIcon(mCurrIcon).FillColor
  picIcon(mCurrIcon).FillStyle = 0
  picIcon(mCurrIcon).FillColor = mColor(mButton)
  Call ExtFloodFill(picIcon(mCurrIcon).hdc, Int(x / mZoom), Int(y / mZoom), picIcon(mCurrIcon).Point(Int(x / mZoom), Int(y / mZoom)), 1)
  picIcon(mCurrIcon).FillStyle = a
  picIcon(mCurrIcon).FillColor = b
End Sub

Private Sub GradientClr(ByRef PicBox As PictureBox, ByVal c1 As Long, ByVal c2 As Long)
'paints a gradient to the referred picture box
'mainly used for displaying the color selections
Dim r(2) As Single, g(2) As Single, b(2) As Single
Dim i As Integer, ix As Integer

 r(0) = GetRGB(c1).Red
 g(0) = GetRGB(c1).Green
 b(0) = GetRGB(c1).Blue

 r(1) = GetRGB(c2).Red
 g(1) = GetRGB(c2).Green
 b(1) = GetRGB(c2).Blue

i = PicBox.ScaleWidth
If i > 255 Then i = 255

 r(2) = (r(1) - r(0)) / i
 g(2) = (g(1) - g(0)) / i
 b(2) = (b(1) - b(0)) / i

For ix = 0 To PicBox.ScaleWidth

 If r(0) < 0 Then r(0) = 0
 If r(0) > 255 Then r(0) = 255
 If g(0) < 0 Then g(0) = 0
 If g(0) > 255 Then g(0) = 255
 If b(0) < 0 Then b(0) = 0
 If b(0) > 255 Then b(0) = 255

 PicBox.Line (ix, 0)-(ix, PicBox.ScaleHeight), RGB(r(0), g(0), b(0)), BF
 r(0) = r(0) + r(2)
 g(0) = g(0) + g(2)
 b(0) = b(0) + b(2)
 
Next ix
End Sub

Public Sub IconFlip(Picture1 As PictureBox)
   'flip vertical
   Dim px As Integer, py As Integer
   px = Picture1.ScaleWidth
   py = Picture1.ScaleHeight
   Call StretchBlt(Picture1.hdc, 0, py, px, -py - 1, Picture1.hdc, 0, 0, px, py, vbSrcCopy)
End Sub

Public Sub IconMirror(Picture1 As PictureBox)
   'flip horizontal
   Dim px As Integer, py As Integer
   px = Picture1.ScaleWidth - 1
   py = Picture1.ScaleHeight - 1
   Set picTemp.Picture = picTemp.Image
   Call StretchBlt(Picture1.hdc, px, 0, -px, py, Picture1.hdc, 0, 0, px, py, vbSrcCopy)
End Sub

Private Sub LoadPrefs()
'load preferences
Dim s As String
s = GetFromINI("Main", "Grid", App.Path & "\prefs.ini")
If s = "" Or s = "1" Then mDrawGrid = True Else mDrawGrid = False
s = GetFromINI("Main", "Zoom", App.Path & "\prefs.ini")
If s = "" Or IsNumeric(s) = False Then
 Call mnuViewZoomSize_Click(10)
Else
 Call mnuViewZoomSize_Click(CInt(s))
End If
s = GetFromINI("Main", "RecentCount", App.Path & "\prefs.ini")
If IsNumeric(s) = True Then
 Dim x As Integer, y As Integer
 y = CInt(s)
 If y <> 0 Then
  s = GetFromINI("Main", "Recent" & 0, App.Path & "\prefs.ini")
  If s <> "" Then
   mnuFileRecentFiles(0).Caption = s
   mnuFileRecentFiles(0).Enabled = True
  End If
  For x = 1 To y - 1
   s = GetFromINI("Main", "Recent" & x, App.Path & "\prefs.ini")
   If s <> "" Then
    Call Load(mnuFileRecentFiles(x))
    mnuFileRecentFiles(x).Caption = s
    mnuFileRecentFiles(x).Enabled = True
   End If
  Next x
 End If
End If

s = GetFromINI("Main", "Window", App.Path & "\prefs.ini")
If s = "" Or IsNumeric(s) = False Then
 Me.WindowState = vbNormal
Else
 Me.WindowState = CInt(s)
End If

If FileExist(App.Path & "\lan.lpk") = True Then Call LoadLanguage(App.Path & "\lan.lpk")

End Sub

Private Sub NewIcon(Optional ByVal Size As Integer = 32)
'completes the steps to create a new fresh icon
Dim i As Integer, x As Integer, p As Integer
p = -1

For i = 0 To picIcon.Count - 1
 If picIcon(i).Tag = "" Then
  If p <> -1 Then p = i
 Else
  x = x + ((picIcon(i).Height + txtIcon(i).Height) + 8)
 End If
Next i
If p <> -1 Then GoTo 1
p = picIcon.Count
Call Load(picIcon(p))
Call Load(txtIcon(p))
Call Load(picUndo(p))
1
With picIcon(p)
 .Tag = "untitled"
 Set .Picture = LoadPicture()
 .Top = x
 .Left = 6 + ((32 - Size) / 2)
 .Height = Size
 .Width = Size
 .Visible = True
End With

picUndo(p).Width = Size
picUndo(p).Height = Size

txtIcon(p).Top = x + picIcon(p).Height + 4
txtIcon(p).Visible = True
txtIcon(p).Caption = "untitled" & vbCrLf & picIcon(p).Width & "x" & picIcon(p).Height

picEdit.Height = Size * mZoom
picEdit.Width = Size * mZoom

Set picEdit.Picture = LoadPicture()
picEdit.Visible = True
mCurrIcon = p
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
picIconsMove.Height = x + (((picIcon(p).Height + txtIcon(p).Height) + 8) * 2)
End Sub

Private Sub NewPlugin(ByVal f As String)
'really just creates a menu item for the plugin 'f'
Dim i As Integer
For i = 0 To mnuPluginsRun.Count - 1
 If mnuPluginsRun(i).Tag = "" Then GoTo 1
Next i

Call Load(mnuPluginsRun(i))

1
 mnuPluginsRun(i).Tag = f
 mnuPluginsRun(i).Caption = f
 mnuPluginsRun(i).Visible = True
End Sub

Private Sub SavePrefs()
'saves preferences
Dim s As String
Call WriteToINI("Main", "Grid", IIf(mDrawGrid = True, 1, 0), App.Path & "\prefs.ini")
Call WriteToINI("Main", "Zoom", CStr(mZoom), App.Path & "\prefs.ini")
Call WriteToINI("Main", "Window", CStr(Me.WindowState), App.Path & "\prefs.ini")
Call WriteToINI("Main", "RecentCount", CStr(mnuFileRecentFiles.Count), App.Path & "\prefs.ini")

Dim x As Integer
 For x = 0 To mnuFileRecentFiles.Count - 1
  If mnuFileRecentFiles(x).Caption = "(empty)" Then Exit For
  Call WriteToINI("Main", "Recent" & x, mnuFileRecentFiles(x).Caption, App.Path & "\prefs.ini")
 Next x
End Sub

Public Sub SetFont(ByVal x As Integer, ByVal y As Integer, ByVal FontName As String, ByVal FontSize As Integer, ByVal Text As String)
'accessed by frmText to draw text to the current icon
If FontName = "" Then Exit Sub
With picIcon(mCurrIcon)
 .CurrentX = x - 1
 .CurrentY = y - 4
 .ForeColor = mColor(1)
 .Font = FontName
 .FontSize = FontSize
End With
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
  Dim arr() As String, v As Variant
  arr() = Split(Text, vbCrLf)
  For Each v In arr()
   picIcon(mCurrIcon).Print v
   picIcon(mCurrIcon).CurrentY = picIcon(mCurrIcon).CurrentY + (FontSize / 5) - 5
   picIcon(mCurrIcon).CurrentX = x - 1
  Next v
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  picEdit.Cls
  Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
  Call DrawGrid
  picEdit.Refresh
  Set picEdit.Picture = picEdit.Image

End Sub

Private Sub ShowRGBVal(ByVal Index As Integer)
'displays RGB and Hex Values to the color textboxes
Dim l As Long
l = mColor(Index)
txtClr(0).Text = GetRGB(l).Red
txtClr(1).Text = GetRGB(l).Green
txtClr(2).Text = GetRGB(l).Blue
txtClr(3).Text = IIf(Len(Hex(txtClr(0).Text)) = 1, "0", "") & Hex(txtClr(0).Text) & IIf(Len(Hex(txtClr(1).Text)) = 1, "0", "") & Hex(txtClr(1).Text) & IIf(Len(Hex(txtClr(2).Text)) = 1, "0", "") & Hex(txtClr(2).Text)
End Sub

Private Sub Tool_Box_Move(ByVal x As Integer, ByVal y As Integer, Optional OptLine As Integer, Optional HideBox As Boolean)
'displays a box outline when drawing a box or gradient box
'OptLine is the center line's orientation
On Error Resume Next
If X1 = -1 Then Exit Sub
picEdit.Cls

X2 = Int(x / mZoom) * mZoom + mZoom - 1
Y2 = Int(y / mZoom) * mZoom + mZoom - 1
picEdit.DrawMode = vbInvert
If HideBox = False Then
 If KeyDown = 17 Then
  Y2 = Y1 + (X2 - X1)
  picEdit.Line (X1, Y1)-(X2, Y2), , B
 Else
  picEdit.Line (X1, Y1)-(X2, Y2), , B
 End If
End If
Select Case OptLine
 Case 1
  picEdit.Line (X1, Y1 + ((Y2 - Y1) / 2))-(X2, Y1 + ((Y2 - Y1) / 2)), 0
 Case 2
  picEdit.Line (X1 + ((X2 - X1) / 2), Y1)-(X1 + ((X2 - X1) / 2), Y2), 0
 Case 3
  picEdit.Line (X2, Y1)-(X1, Y2), 0
 Case 4
  picEdit.Line (X1, Y1)-(X2, Y2), 0
End Select
picEdit.DrawMode = vbCopyPen
End Sub

Private Sub Tool_Box_Up(ByVal x As Integer, ByVal y As Integer, Optional ByVal Fill As Boolean = False, Optional OutLine As Boolean)
X1 = Int(X1 / mZoom)
X2 = Int(X2 / mZoom)
Y1 = Int(Y1 / mZoom)
Y2 = Int(Y2 / mZoom)

If Fill = True Then
 picIcon(mCurrIcon).Line (X1, Y1)-(X2, Y2), mColor(mButton), BF
Else
 picIcon(mCurrIcon).Line (X1, Y1)-(X2, Y2), mColor(mButton), B
End If

Dim b As Integer
b = mButton
If b = 1 Then b = 2 Else b = 1
If OutLine = True Then picIcon(mCurrIcon).Line (X1, Y1)-(X2, Y2), mColor(b), B
End Sub

Private Sub Tool_BoxGrad_Up(ByVal x As Single, ByVal y As Single, Optional Direction As UDE_GRADDIR)
On Error Resume Next
Dim r(2) As Integer
Dim g(2) As Integer
Dim b(2) As Integer

Dim c1 As Long, c2 As Long, ix As Long
If mButton = 1 Then
 c1 = mColor(1)
 c2 = mColor(2)
Else
 c2 = mColor(1)
 c1 = mColor(2)
End If

 r(0) = GetRGB(c1).Red
 g(0) = GetRGB(c1).Green
 b(0) = GetRGB(c1).Blue

 r(1) = GetRGB(c2).Red
 g(1) = GetRGB(c2).Green
 b(1) = GetRGB(c2).Blue
 Dim hOff As Integer

Select Case Direction
 Case gEW

    r(2) = (r(1) - r(0)) / Int((x / mZoom) - Int(X1 / mZoom) - 1)
    g(2) = (g(1) - g(0)) / Int((x / mZoom) - Int(X1 / mZoom) - 1)
    b(2) = (b(1) - b(0)) / Int((x / mZoom) - Int(X1 / mZoom) - 1)

    For ix = Int(X1 / mZoom) To Int(X2 / mZoom)
    
     If r(0) < 0 Then r(0) = 0
     If r(0) > 255 Then r(0) = 255
     If g(0) < 0 Then g(0) = 0
     If g(0) > 255 Then g(0) = 255
     If b(0) < 0 Then b(0) = 0
     If b(0) > 255 Then b(0) = 255
    
       picIcon(mCurrIcon).Line (ix, Int(Y1 / mZoom))-(ix, Int(Y2 / mZoom)), RGB(r(0), g(0), b(0)), BF
     
     r(0) = r(0) + r(2)
     g(0) = g(0) + g(2)
     b(0) = b(0) + b(2)
    
    Next ix
 Case gNS
    r(2) = (r(1) - r(0)) / Int((y / mZoom) - Int(Y1 / mZoom) - 1)
    g(2) = (g(1) - g(0)) / Int((y / mZoom) - Int(Y1 / mZoom) - 1)
    b(2) = (b(1) - b(0)) / Int((y / mZoom) - Int(Y1 / mZoom) - 1)
 
    For ix = Int(Y1 / mZoom) To Int(Y2 / mZoom)
    
     If r(0) < 0 Then r(0) = 0
     If r(0) > 255 Then r(0) = 255
     If g(0) < 0 Then g(0) = 0
     If g(0) > 255 Then g(0) = 255
     If b(0) < 0 Then b(0) = 0
     If b(0) > 255 Then b(0) = 255
    
       picIcon(mCurrIcon).Line (Int(X1 / mZoom), ix)-(Int(X2 / mZoom), ix), RGB(r(0), g(0), b(0)), BF
     
     r(0) = r(0) + r(2)
     g(0) = g(0) + g(2)
     b(0) = b(0) + b(2)
    
    Next ix
 Case gNESW

   Set picTemp.Picture = LoadPicture()
    picTemp.Height = Int(Y2 / mZoom) - Int(Y1 / mZoom)
    hOff = picTemp.Height / 2
    picTemp.Width = Int(X2 / mZoom) - Int(X1 / mZoom)
    Debug.Print (picTemp.Width + hOff)
    r(2) = (r(1) - r(0)) / ((picTemp.Width + (hOff * 2))) ' * IIf(mButton = 1, 1, 1.3))
    g(2) = (g(1) - g(0)) / ((picTemp.Width + (hOff * 2))) ' * IIf(mButton = 1, 1, 1.3))
    b(2) = (b(1) - b(0)) / ((picTemp.Width + (hOff * 2))) ' * IIf(mButton = 1, 1, 1.3))
    
    For ix = 0 - hOff To picTemp.Width + hOff
    
     If r(0) < 0 Then r(0) = 0
     If r(0) > 255 Then r(0) = 255
     If g(0) < 0 Then g(0) = 0
     If g(0) > 255 Then g(0) = 255
     If b(0) < 0 Then b(0) = 0
     If b(0) > 255 Then b(0) = 255
    
       picTemp.Line (ix + hOff, 0)-(ix - hOff, picTemp.Height), RGB(r(0), g(0), b(0))
    
     r(0) = r(0) + r(2)
     g(0) = g(0) + g(2)
     b(0) = b(0) + b(2)
    
    Next ix
    Set picTemp.Picture = picTemp.Image
    picIcon(mCurrIcon).PaintPicture picTemp.Picture, Int(X1 / mZoom), Int(Y1 / mZoom), Int(X2 / mZoom) - Int(X1 / mZoom) + 1, Int(Y2 / mZoom) - Int(Y1 / mZoom) + 1, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
 Case gNWSE

   Set picTemp.Picture = LoadPicture()
    picTemp.Height = Int(Y2 / mZoom) - Int(Y1 / mZoom)
    hOff = picTemp.Height / 2
    picTemp.Width = Int(X2 / mZoom) - Int(X1 / mZoom)

    r(2) = (r(1) - r(0)) / ((picTemp.Width + (hOff * 2)))
    g(2) = (g(1) - g(0)) / ((picTemp.Width + (hOff * 2)))
    b(2) = (b(1) - b(0)) / ((picTemp.Width + (hOff * 2)))

    For ix = picTemp.Width + hOff To 0 - hOff Step -1
    
     If r(0) < 0 Then r(0) = 0
     If r(0) > 255 Then r(0) = 255
     If g(0) < 0 Then g(0) = 0
     If g(0) > 255 Then g(0) = 255
     If b(0) < 0 Then b(0) = 0
     If b(0) > 255 Then b(0) = 255
    
       picTemp.Line (ix - hOff, 0)-(ix + hOff, picTemp.Height), RGB(r(0), g(0), b(0))
    
     r(0) = r(0) + r(2)
     g(0) = g(0) + g(2)
     b(0) = b(0) + b(2)
    
    Next ix
    Set picTemp.Picture = picTemp.Image
    picIcon(mCurrIcon).PaintPicture picTemp.Picture, Int(X1 / mZoom), Int(Y1 / mZoom), Int(X2 / mZoom) - Int(X1 / mZoom) + 1, Int(Y2 / mZoom) - Int(Y1 / mZoom) + 1, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy

End Select
End Sub

Private Sub Tool_CircleGrad_Up(ByVal x As Single, ByVal y As Single)
On Error Resume Next
Dim r(2) As Integer
Dim g(2) As Integer
Dim b(2) As Integer
X1 = Int(X1 / mZoom)
X2 = Int(X2 / mZoom)
Y1 = Int(Y1 / mZoom)
Y2 = Int(Y2 / mZoom)
Dim ra As Single, a As Single
Dim c1 As Long, c2 As Long, ix As Single, xc As Single, yc As Single
If Abs(X2 - X1) > Abs(Y2 - Y1) Then ra = Abs(X2 - X1) / 2 Else ra = Abs(Y2 - Y1) / 2
If KeyDown = 17 Then a = 1 Else a = Abs(Abs(Y1 - Y2) / Abs(X2 - X1))

xc = X1 + (X2 - X1) / 2
yc = Y1 + (Y2 - Y1) / 2

'Debug.Print KeyDown

If mButton = 1 Then
 c1 = mColor(1)
 c2 = mColor(2)
Else
 c2 = mColor(1)
 c1 = mColor(2)
End If

 r(0) = GetRGB(c1).Red
 g(0) = GetRGB(c1).Green
 b(0) = GetRGB(c1).Blue

 r(1) = GetRGB(c2).Red
 g(1) = GetRGB(c2).Green
 b(1) = GetRGB(c2).Blue

 r(2) = (r(1) - r(0)) / Int(ra - 1)
 g(2) = (g(1) - g(0)) / Int(ra - 1)
 b(2) = (b(1) - b(0)) / Int(ra - 1)

Dim H As Long, k As Long
H = picIcon(mCurrIcon).FillStyle
k = picIcon(mCurrIcon).FillColor
picIcon(mCurrIcon).FillStyle = 0
For ix = ra To 0 Step -1

 If r(0) < 0 Then r(0) = 0
 If r(0) > 255 Then r(0) = 255
 If g(0) < 0 Then g(0) = 0
 If g(0) > 255 Then g(0) = 255
 If b(0) < 0 Then b(0) = 0
 If b(0) > 255 Then b(0) = 255
   picIcon(mCurrIcon).FillColor = RGB(r(0), g(0), b(0))
   picIcon(mCurrIcon).Circle (xc, yc), ix, RGB(r(0), g(0), b(0)), , , a
 
 r(0) = r(0) + r(2)
 g(0) = g(0) + g(2)
 b(0) = b(0) + b(2)

Next ix
picIcon(mCurrIcon).FillColor = k
picIcon(mCurrIcon).FillStyle = H
End Sub

Private Sub Tool_Circle_Move(ByVal x As Integer, ByVal y As Integer)
On Error Resume Next
If X1 = -1 Then Exit Sub
picEdit.Cls
X2 = Int(x / mZoom) * mZoom + mZoom - 1
Y2 = Int(y / mZoom) * mZoom + mZoom - 1


picEdit.DrawMode = vbInvert
Dim r As Integer, a As Single, xc As Single, yc As Single
If Abs(X2 - X1) > Abs(Y2 - Y1) Then r = Abs(X2 - X1) / 2 Else r = Abs(Y2 - Y1) / 2
If KeyDown = 17 Then a = 1 Else a = Abs(Abs(Y1 - Y2) / Abs(X2 - X1)): picEdit.Line (X1, Y1)-(X2, Y2), , B

xc = X1 + (X2 - X1) / 2
yc = Y1 + (Y2 - Y1) / 2
picEdit.Circle (xc, yc), r, , , , a
picEdit.DrawMode = vbCopyPen

End Sub

Private Sub Tool_Circle_Up(ByVal x As Integer, ByVal y As Integer, Optional ByVal Fill As Boolean = False, Optional OutLine As Boolean)
X1 = Int(X1 / mZoom)
X2 = Int(X2 / mZoom)
Y1 = Int(Y1 / mZoom)
Y2 = Int(Y2 / mZoom)
Dim r As Single, a As Single
Dim xc As Single, yc As Single
If Abs(X2 - X1) > Abs(Y2 - Y1) Then r = Abs(X2 - X1) / 2 Else r = Abs(Y2 - Y1) / 2
If Abs(Y1 - Y2) = Abs(X2 - X1) Then a = 1 Else If KeyDown = 17 Then a = 1 Else a = Abs(Abs(Y1 - Y2) / Abs(X2 - X1))

xc = X1 + (X2 - X1) / 2
yc = Y1 + (Y2 - Y1) / 2

'Debug.Print KeyDown

If Fill = True Then

Dim H As Long, k As Long
H = picIcon(mCurrIcon).FillStyle
k = picIcon(mCurrIcon).FillColor

picIcon(mCurrIcon).FillStyle = 0
picIcon(mCurrIcon).FillColor = mColor(mButton)
 picIcon(mCurrIcon).Circle (xc, yc), r, mColor(mButton), , , a
picIcon(mCurrIcon).FillColor = k
picIcon(mCurrIcon).FillStyle = H

Dim b As Integer
b = mButton
If b = 1 Then b = 2 Else b = 1
If OutLine = True Then picIcon(mCurrIcon).Circle (xc, yc), r, mColor(b), , , a


' Call Fill
Else
 picIcon(mCurrIcon).Circle (xc, yc), r, mColor(mButton), , , a
End If
End Sub

Private Sub Tool_Line_Move(ByVal x As Integer, ByVal y As Integer)
On Error Resume Next

Dim x3 As Integer, y3 As Integer
picEdit.Cls
X2 = Int(x / mZoom) * mZoom + mZoom - 1
Y2 = Int(y / mZoom) * mZoom + mZoom - 1
picEdit.DrawMode = vbInvert
picEdit.Line (X1 + (mZoom / 2), Y1 + (mZoom / 2))-(X2 - (mZoom / 2), Y2 - (mZoom / 2)), 0
picEdit.DrawMode = vbCopyPen
End Sub

Private Function Truncate(ByVal s As String) As String
If Len(s) > 8 Then Truncate = Left(s, 7) & "..." Else Truncate = s
End Function

Private Sub Form_Load()
Call LoadPrefs
txtIcon(0).Caption = "untitled" & vbCrLf & "32x32"

Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Call DrawGrid
SelTool = tDraw
mColor(1) = 0
mColor(2) = vbWhite
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
'draw colors
Dim x As Integer, arrClr(1, 14) As Long

arrClr(0, 0) = vbBlack
arrClr(0, 1) = RGB(128, 128, 128)
arrClr(0, 2) = RGB(128, 0, 0)
arrClr(0, 3) = RGB(128, 128, 0)
arrClr(0, 4) = RGB(0, 128, 0)
arrClr(0, 5) = RGB(0, 128, 128) + 1
arrClr(0, 6) = RGB(0, 0, 128)
arrClr(0, 7) = RGB(128, 0, 128)
arrClr(0, 8) = RGB(128, 128, 64)
arrClr(0, 9) = RGB(0, 64, 64)
arrClr(0, 10) = RGB(0, 128, 255)
arrClr(0, 11) = RGB(0, 64, 128)
arrClr(0, 12) = RGB(64, 0, 255)
arrClr(0, 13) = RGB(128, 64, 0)


arrClr(1, 0) = vbWhite
arrClr(1, 1) = RGB(192, 192, 192)
arrClr(1, 2) = RGB(255, 0, 0)
arrClr(1, 3) = RGB(255, 255, 0)
arrClr(1, 4) = RGB(0, 255, 0)
arrClr(1, 5) = RGB(0, 255, 255)
arrClr(1, 6) = RGB(0, 0, 255)
arrClr(1, 7) = RGB(255, 0, 255)
arrClr(1, 8) = RGB(255, 255, 128)
arrClr(1, 9) = RGB(0, 255, 128)
arrClr(1, 10) = RGB(128, 255, 255)
arrClr(1, 11) = RGB(128, 128, 255)
arrClr(1, 12) = RGB(255, 0, 128)
arrClr(1, 13) = RGB(255, 128, 64)

picClrSel(0).Width = 13 * 12 + 3
picClrSel(0).Line (0, 0)-(200, 38), arrClr(0, Int(x / 11)), BF
For x = 0 To 13 * 11
 picClrSel(0).Line (x, 1)-(x + 11, 18), arrClr(0, Int(x / 11)), BF
 picClrSel(0).Line (x, 21)-(x + 11, 36), arrClr(1, Int(x / 11)), BF
 x = x + 10
Next x

On Error GoTo 1
Call MkDir(App.Path & "\Plugins\")
1
flbPlugin.Path = App.Path & "\Plugins\"
flbPlugin.Refresh

For x = 0 To flbPlugin.ListCount - 1
 Call NewPlugin(Left(flbPlugin.List(x), Len(flbPlugin.List(x)) - 4))
Next x
 
If x = 0 Then mnuPluginsRun(0).Caption = "(empty)": mnuPluginsRun(0).Enabled = False


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SavePrefs
Dim i As Integer, a As Integer
For i = 0 To picIcon.Count - 1
 If picIcon(i).Tag <> "" And FileChanged(i) = True Then
  a = MsgBox(ExtractFileName(picIcon(i).Tag) & " has changed." & vbCrLf & "Do you want to save it?", vbQuestion + vbYesNoCancel, "Save Changes")
  If a = vbYes Then 'save
   If InStr(picIcon(i).Tag, "ico") Then
    Open picIcon(i).Tag For Binary Access Write As #1
     Put #1, 1, GenerateIconForSave$(picIcon(i))
    Close #1
   Else
    Call mnuFileSaveAs_Click
   End If

  ElseIf a = vbCancel Then
   Cancel = -1
   Exit Sub
  End If
 End If
Next i
If FileExist(App.Path & "\temp.bmp") Then Kill App.Path & "\temp.bmp"
End
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
If Me.ScaleWidth < 610 Then Me.Width = 610 * Screen.TwipsPerPixelX
If Me.ScaleHeight < 500 Then Me.Height = 500 * Screen.TwipsPerPixelY
On Error Resume Next
Dim w As Integer, H As Integer
w = Me.ScaleWidth
H = Me.ScaleHeight

picBack.Width = w - picBack.Left
picBack.Height = H - Frame1.Height + 4

Frame1.Top = H - Frame1.Height
Frame1.Width = picBack.Width

picIconsBack.Height = picBack.Height - picBack.Top - picIconsBack.Top
End Sub

Private Sub hsEdit_Change()
picEdit.Left = hsEdit.Value + 8
End Sub

Private Sub lblSwitch_Click()
Dim l As Long
l = picClr(1).BackColor
picClr(1).BackColor = picClr(2).BackColor
picClr(2).BackColor = l

mColor(1) = picClr(1).BackColor
mColor(2) = picClr(2).BackColor
End Sub

Private Sub mnuEdit_Click()
 mnuEditCopy.Enabled = shpSel.Visible
 mnuEditCut.Enabled = shpSel.Visible
 mnuEditDelete.Enabled = shpSel.Visible
 mnuEditPaste.Enabled = Clipboard.GetFormat(2)
End Sub

Private Sub mnuEditCopy_Click()
If shpSel.Visible = True Then
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
 With picTemp
  .Width = Int(shpSel.Width / mZoom) '- 1
  .Height = Int(shpSel.Height / mZoom) ' - 1
  .PaintPicture picIcon(mCurrIcon).Picture, 0, 0, picTemp.Width, picTemp.Height, Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom), Int(shpSel.Width / mZoom), Int(shpSel.Height / mZoom), vbSrcCopy
  Set .Picture = .Image
  .Refresh
  Clipboard.Clear
  Clipboard.SetData picTemp.Picture, 2
 End With
End If
End Sub

Private Sub mnuEditCut_Click()
If shpSel.Visible = True Then
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
 With picTemp
  .Width = Int(shpSel.Width / mZoom) '- 1
  .Height = Int(shpSel.Height / mZoom) ' - 1
  .PaintPicture picIcon(mCurrIcon).Picture, 0, 0, picTemp.Width, picTemp.Height, Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom), Int(shpSel.Width / mZoom), Int(shpSel.Height / mZoom), vbSrcCopy
  Set .Picture = .Image
  .Refresh
  Clipboard.Clear
  Clipboard.SetData picTemp.Picture, 2
 End With
 Call mnuEditDelete_Click
End If
End Sub

Private Sub mnuEditDelete_Click()
If shpSel.Visible = True Then
 FileChanged(mCurrIcon) = True
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
 picIcon(mCurrIcon).Line (Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom))-(Int(shpSel.Left / mZoom) + Int(shpSel.Width / mZoom) - 1, Int(shpSel.Top / mZoom) + Int(shpSel.Height / mZoom) - 1), picTrans.BackColor, BF
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 picEdit.Cls
 Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
 Call DrawGrid
 picEdit.Refresh
 Set picEdit.Picture = picEdit.Image
End If
End Sub

Private Sub mnuEditPaste_Click()
On Error GoTo 1
FileChanged(mCurrIcon) = True
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
Set picTemp.Picture = Clipboard.GetData(2)
If shpSel.Visible = True Then
 picIcon(mCurrIcon).PaintPicture picTemp.Picture, Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom), Int(shpSel.Width / mZoom), Int(shpSel.Height / mZoom), 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
Else
 picIcon(mCurrIcon).PaintPicture picTemp.Picture, 0, 0, picTemp.Width, picTemp.Height, 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
End If
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
1
End Sub

Private Sub mnuEditSelAll_Click()
Dim b As MSComctlLib.Button
Set b = tbTools.Buttons(1)
Call tbTools_ButtonClick(b)
shpSel.Left = 1
shpSel.Top = 1
shpSel.Width = picIcon(mCurrIcon).Width * mZoom - 1
shpSel.Height = picIcon(mCurrIcon).Height * mZoom - 1
shpSel.Visible = True
End Sub

Private Sub mnuEditUndo_Click()
Set picIcon(mCurrIcon).Picture = picUndo(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End Sub

Private Sub mnuFileClose_Click()
Dim a As Integer
 If picIcon(mCurrIcon).Tag <> "" And FileChanged(mCurrIcon) = True Then
  a = MsgBox(ExtractFileName(picIcon(mCurrIcon).Tag) & " has changed." & vbCrLf & "Do you want to save it?", vbQuestion + vbYesNoCancel, "Save Changes")
  If a = vbYes Then 'save
   If InStr(picIcon(mCurrIcon).Tag, "ico") Then
    Open picIcon(mCurrIcon).Tag For Binary Access Write As #1
     Put #1, 1, GenerateIconForSave$(picIcon(mCurrIcon))
    Close #1
   Else
    Call mnuFileSaveAs_Click
   End If
  ElseIf a = vbCancel Then
   Exit Sub
  End If
 End If

picEdit.Tag = ""
picEdit.Visible = False
picIcon(mCurrIcon).Tag = ""
picIcon(mCurrIcon).Visible = False
txtIcon(mCurrIcon).Caption = ""
txtIcon(mCurrIcon).Visible = False

Dim i As Integer, x As Integer, p As Integer
p = -1
For i = 0 To picIcon.Count - 1
 If picIcon(i).Tag <> "" Then
  picIcon(i).Top = x
  txtIcon(i).Top = x + 36
  x = x + 68
 End If
Next i

picIconsMove.Height = x + 68


End Sub

Private Sub mnuFileExit_Click()
Call Unload(Me)
End
End Sub

Private Sub mnuFileNew16_Click()
Call NewIcon(16)
End Sub

Private Sub mnuFileNew32_Click()
Call NewIcon(32)
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo 1
cd.Filter = "Icon Files (*.ico, *.cur)|*.ico;*.cur|Bitmap Files (*.bmp)|*.bmp|JPEG Files (*.jpg)|*.jpg|GIF Files (*.gif)|*.gif"
cd.ShowOpen
Dim i As Integer
i = GetIconSize(cd.FileName)

Call AddRecent(cd.FileName)
If i <> -1 Then Call NewIcon(i) Else Call NewIcon(32)

Set picUndo(mCurrIcon).Picture = LoadPicture(cd.FileName)
Set picIcon(mCurrIcon).Picture = LoadPicture(cd.FileName)
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picIcon(mCurrIcon).Refresh
picIcon(mCurrIcon).Tag = cd.FileName
txtIcon(mCurrIcon).Caption = Truncate(LCase(ExtractFileName(cd.FileName))) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image

1
End Sub

Private Sub mnuFileRecentFiles_Click(Index As Integer)
Dim i As Integer
Dim s As String
s = mnuFileRecentFiles(Index).Caption
If FileExist(s) = False Then
 MsgBox s & vbCrLf & "Does not exist", vbInformation, "File Error"
 Exit Sub
End If
i = GetIconSize(s)

If i <> -1 Then Call NewIcon(i) Else Call NewIcon(32)

Set picUndo(mCurrIcon).Picture = LoadPicture(s)
Set picIcon(mCurrIcon).Picture = LoadPicture(s)
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picIcon(mCurrIcon).Refresh
picIcon(mCurrIcon).Tag = s
txtIcon(mCurrIcon).Caption = Truncate(LCase(ExtractFileName(s))) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image

End Sub

Private Sub mnuFileSave_Click()
If InStr(picIcon(mCurrIcon).Tag, "ico") Then
 Open picIcon(mCurrIcon).Tag For Binary Access Write As #1
  Put #1, 1, GenerateIconForSave$(picIcon(mCurrIcon))
 Close #1
 FileChanged(mCurrIcon) = False
Else
 Call mnuFileSaveAs_Click
End If
End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo 1
cd.Filter = "Icon Files (*.ico, *.cur)|*.ico;*.cur|Bitmap (*.bmp)|*.bmp"
cd.FileName = picIcon(mCurrIcon).Tag
cd.ShowSave
Call AddRecent(cd.FileName)
If LCase(Right(cd.FileName, 4)) = ".bmp" Then
 Call SavePicture(picIcon(mCurrIcon), cd.FileName)
 Exit Sub
Else
 Dim l As Long
 l = FreeFile()
 Open cd.FileName For Binary Access Write As #l
  Put #l, 1, GenerateIconForSave$(picIcon(mCurrIcon))
 Close #l
End If
picIcon(mCurrIcon).Tag = cd.FileName
txtIcon(mCurrIcon).Caption = ExtractFileName(cd.FileName) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
FileChanged(mCurrIcon) = False
1
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuHelpTest_Click()
Set picIcon(mCurrIcon).Picture = Me.Icon
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End Sub

Private Sub mnuImageClear_Click()
FileChanged(mCurrIcon) = True
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
Set picIcon(mCurrIcon).Picture = LoadPicture()
picEdit.Cls
Set picEdit.Picture = LoadPicture()
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End Sub

Private Sub mnuImageCopyNew_Click()
Dim i As Integer
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
i = mCurrIcon
Call NewIcon(picIcon(mCurrIcon).ScaleWidth)

Call picIcon(mCurrIcon).PaintPicture(picIcon(i).Picture, 0, 0, picIcon(i).Width, picIcon(i).Height, 0, 0, picIcon(i).Width, picIcon(i).Height, vbSrcCopy)
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
txtIcon(mCurrIcon).Caption = Truncate("Copy of " & LCase(ExtractFileName(picIcon(i).Tag))) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End Sub

Private Sub mnuImageFlipHo_Click()
FileChanged(mCurrIcon) = True
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
Call IconMirror(picIcon(mCurrIcon))
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image

End Sub

Private Sub mnuImageFlipVert_Click()
FileChanged(mCurrIcon) = True
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
Call IconFlip(picIcon(mCurrIcon))
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End Sub

Private Sub mnuPluginsRun_Click(Index As Integer)
Call wc.Run(App.Path & "\plugins\" & mnuPluginsRun(Index).Tag & ".exe")
End Sub

Private Sub mnuViewGrid_Click()
mDrawGrid = IIf(mDrawGrid = True, False, True)
mnuViewGrid.Checked = mDrawGrid
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image

End Sub

Private Sub mnuViewZoomSize_Click(Index As Integer)
Dim oz As Integer
oz = mZoom
mZoom = Index
mnuViewZoomSize(1).Checked = False
mnuViewZoomSize(5).Checked = False
mnuViewZoomSize(10).Checked = False
mnuViewZoomSize(20).Checked = False
mnuViewZoomSize(Index).Checked = True

 picEdit.Width = picIcon(mCurrIcon).Width * mZoom
 picEdit.Height = picIcon(mCurrIcon).Height * mZoom
 picIcon(mCurrIcon).Refresh
 Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
 picEdit.Cls
 Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
 Call DrawGrid
 picEdit.Refresh
 Set picEdit.Picture = picEdit.Image
 If shpSel.Visible = True Then
  shpSel.Left = Int(shpSel.Left / oz) * mZoom
  shpSel.Top = Int(shpSel.Top / oz) * mZoom
  shpSel.Width = Int(shpSel.Width / oz) * mZoom
  shpSel.Height = Int(shpSel.Height / oz) * mZoom
  OldSel(0) = shpSel.Left: OldSel(1) = shpSel.Top
  OldSel(2) = shpSel.Width: OldSel(3) = shpSel.Height
 End If
Call picBack_Resize
End Sub

Private Sub optEdit_Click()
MsgBox mCustomTool
optEdit.Value = False
End Sub

Private Sub picBack_Resize()
hsEdit.Left = 0
hsEdit.Top = picBack.ScaleHeight - hsEdit.Height
hsEdit.Width = picBack.ScaleWidth - vsEdit.Width
vsEdit.Top = 0
vsEdit.Left = picBack.ScaleWidth - vsEdit.Width
vsEdit.Height = picBack.ScaleHeight - hsEdit.Height
optEdit.Width = vsEdit.Width
optEdit.Height = hsEdit.Height
optEdit.Top = hsEdit.Top
optEdit.Left = vsEdit.Left

If picEdit.Height > picBack.ScaleHeight - hsEdit.Height Then
 vsEdit.Max = picBack.ScaleHeight - picEdit.Height - hsEdit.Height - 16
 vsEdit.Value = 0
Else
 vsEdit.Max = 0
 vsEdit.Value = 0
End If
If picEdit.Width > picBack.ScaleWidth - vsEdit.Width Then
 hsEdit.Max = picBack.ScaleWidth - picEdit.Width - vsEdit.Width - 8
 hsEdit.Value = 0
Else
 hsEdit.Max = 0
 hsEdit.Value = 0
End If
End Sub

Private Sub picClr_Click(Index As Integer)
On Error GoTo 1
cd.Color = picClr(Index).BackColor
cd.ShowColor
Dim l As Long
If cd.Color = picTrans.BackColor Then l = cd.Color + 1 Else l = cd.Color

picClr(Index).BackColor = l
mColor(Index) = l

  Call GradientClr(picClrSel(1), vbBlack, l)
  Call GradientClr(picClrSel(2), l, vbWhite)

Call ShowRGBVal(Index)
1
End Sub

Private Sub picClrSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
 Case 0
  picClr(Button).BackColor = picClrSel(0).Point(x, y)
  mColor(Button) = picClrSel(0).Point(x, y)
 
  Call GradientClr(picClrSel(1), vbBlack, picClrSel(0).Point(x, y))
  Call GradientClr(picClrSel(2), picClrSel(0).Point(x, y), vbWhite)
 Case 1
  If Button <> 0 Then picClr(Button).BackColor = picClrSel(1).Point(x, y): mColor(Button) = picClrSel(1).Point(x, y)
 Case 2
  If Button <> 0 Then picClr(Button).BackColor = picClrSel(2).Point(x, y): mColor(Button) = picClrSel(2).Point(x, y)
End Select

Call ShowRGBVal(Button)
End Sub

Private Sub picClrSel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 0 Then Exit Sub
If x < 0 Or y < 0 Then Exit Sub
If x >= picClrSel(Index).ScaleWidth Or y >= picClrSel(Index).ScaleHeight Then Exit Sub
Select Case Index
 Case 1
  If Button <> 0 Then picClr(Button).BackColor = picClrSel(1).Point(x, y): mColor(Button) = picClrSel(1).Point(x, y)
 Case 2
  If Button <> 0 Then picClr(Button).BackColor = picClrSel(2).Point(x, y): mColor(Button) = picClrSel(2).Point(x, y)
End Select

Call ShowRGBVal(Button)
End Sub

Private Sub picEdit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 68 Then
 mColor(1) = 0
 mColor(2) = vbWhite
 picClr(1).BackColor = 0
 picClr(2).BackColor = vbWhite
 Call ShowRGBVal(1)
 Exit Sub
End If
KeyDown = Shift + KeyCode
Select Case SelTool
 Case tBoxGrad, tBoxGradNS
  Call Tool_Box_Move(X2, Y2, True)
 Case tCircleGrad
  Call Tool_Circle_Move(X2, Y2)
 Case tBox To tBoxFill
  Call Tool_Box_Move(X2, Y2)
 Case tCircle To tCircleFill
  Call Tool_Circle_Move(X2, Y2)
End Select

End Sub

Private Sub picEdit_KeyUp(KeyCode As Integer, Shift As Integer)
KeyDown = 0

Select Case SelTool
 Case tBoxGrad, tBoxGradNS
  Call Tool_Box_Move(X2, Y2, True)
 Case tCircleGrad
  Call Tool_Circle_Move(X2, Y2)
 Case tBox To tBoxFill
  Call Tool_Box_Move(X2, Y2)
 Case tCircle To tCircleFill
  Call Tool_Circle_Move(X2, Y2)
End Select

End Sub

Private Sub picEdit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 And Button <> 2 Then Exit Sub
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
mButton = Button
X1 = Int(x / mZoom) * mZoom
Y1 = Int(y / mZoom) * mZoom
Select Case SelTool
 Case tSelect
  If X1 < shpSel.Left Or X1 > shpSel.Left + shpSel.Width Or _
     Y1 < shpSel.Top Or Y1 > shpSel.Top + shpSel.Height Then
      shpSel.Visible = False
      shpSel.Left = -10000
      shpSel.Top = -10000
      shpSel.Width = 10
      shpSel.Height = 20
      MovingSel = False
  End If
 Case tDraw
  X1 = Int(x / mZoom) * mZoom
  Y1 = Int(y / mZoom) * mZoom
  X2 = Int(x / mZoom) * mZoom + mZoom - 1
  Y2 = Int(y / mZoom) * mZoom + mZoom - 1
  picEdit.Line (X1, Y1)-(X2, Y2), mColor(mButton), BF
  Set picEdit.Picture = picEdit.Image
  picIcon(mCurrIcon).PSet (Int(X1 / mZoom), Int(y / mZoom)), mColor(mButton)
 Case tEye
  mColor(Button) = picIcon(mCurrIcon).Point(Int(x / mZoom), Int(y / mZoom))
  picClr(Button).BackColor = mColor(Button)
  Call ShowRGBVal(Button)
  Call GradientClr(picClrSel(1), vbBlack, mColor(Button))
  Call GradientClr(picClrSel(2), mColor(Button), vbWhite)
 Case tcustom
  If tCustomType = 2 Then
  X2 = Int(x / mZoom) * mZoom
  Y2 = Int(y / mZoom) * mZoom
  Dim s As String, a As String, arr() As String, arrx() As String, v As Variant
  s = mCustomTool
  arrx() = Split(s, vbCrLf)
  For Each v In arrx()
  s = v
  Call Randomize
  s = Replace(s, "wrnd", Int(Int((X2 / mZoom) - Int(X1 / mZoom)) * Rnd))
  Call Randomize
  s = Replace(s, "hrnd", Int(Int((Y2 / mZoom) - Int(Y1 / mZoom)) * Rnd))
  s = Replace(s, "h", Int(Y2 / mZoom) - Int(Y1 / mZoom))
  s = Replace(s, "w", Int(X2 / mZoom) - Int(X1 / mZoom))
  s = Replace(s, "x1", Int(X1 / mZoom))
  s = Replace(s, "x2", Int(X2 / mZoom))
  s = Replace(s, "y1", Int(Y1 / mZoom))
  s = Replace(s, "y2", Int(Y2 / mZoom))
  If InStr(s, " ") <> 0 Then
  a = Left(s, InStr(s, " ") - 1)
  arr() = Split(Mid(s, InStr(s, " ") + 1), ",")
  Select Case LCase(Trim(a))
   Case "line"
    picIcon(mCurrIcon).Line (retMath(arr(0)), retMath(arr(1)))-(retMath(arr(2)), retMath(arr(3))), mColor(Button)
    picIcon(mCurrIcon).PSet (retMath(arr(2)), retMath(arr(3))), mColor(Button)
   Case "dot"
    picIcon(mCurrIcon).PSet (retMath(arr(0)), retMath(arr(1))), mColor(Button)
   Case "box"
    picIcon(mCurrIcon).Line (retMath(arr(0)), retMath(arr(1)))-(retMath(arr(2)), retMath(arr(3))), mColor(Button), B
   Case "filledbox"
    picIcon(mCurrIcon).Line (retMath(arr(0)), retMath(arr(1)))-(retMath(arr(2)), retMath(arr(3))), mColor(Button), BF
   Case "circle"
    picIcon(mCurrIcon).Circle (retMath(arr(0)), retMath(arr(1))), retMath(arr(2)), mColor(Button)
  End Select
  End If
  Next v
  picIcon(mCurrIcon).Refresh
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  picEdit.Cls
  Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
  Call DrawGrid
  picEdit.Refresh
  Set picEdit.Picture = picEdit.Image
  
  End If

End Select
End Sub

Private Sub picEdit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
Caption = ExtractFileName(picIcon(mCurrIcon).Tag) & IIf(FileChanged(mCurrIcon) = True, " *", "") & " - CozIcon [" & Int(x / mZoom) & ", " & Int(y / mZoom) & "]"
If X1 = -1 Then Exit Sub
Select Case SelTool
 Case tDraw
  
  X1 = Int(x / mZoom) * mZoom
  Y1 = Int(y / mZoom) * mZoom
  X2 = Int(x / mZoom) * mZoom + mZoom - 1
  Y2 = Int(y / mZoom) * mZoom + mZoom - 1
  picEdit.Cls
  picEdit.Line (X1, Y1)-(X2, Y2), , B
End Select
'picedit.Cls
If Button <> 1 And Button <> 2 Then Exit Sub

Select Case SelTool
 Case tSelect
  X2 = Int(x / mZoom) * mZoom + mZoom - 1
  Y2 = Int(y / mZoom) * mZoom + mZoom - 1
    If MovingSel = True Then
     shpSel.Left = xz1 + (X2 - X1)
     shpSel.Top = yz1 + (Y2 - Y1)
     Exit Sub
    End If
     
     xz1 = X1 - (mZoom - 1)
     yz1 = Y1 - (mZoom - 1)
    shpSel.Left = X1
    shpSel.Top = Y1
    If X2 > X1 Then shpSel.Width = Int(X2 - X1) + 2
    If shpSel.Width > picIcon(mCurrIcon).Width * mZoom Then shpSel.Width = picIcon(mCurrIcon).Width * mZoom
    If Y2 > Y1 Then shpSel.Height = Int(Y2 - Y1) + 2
    If shpSel.Height > picIcon(mCurrIcon).Height * mZoom Then shpSel.Height = picIcon(mCurrIcon).Height * mZoom
    OldSel(0) = shpSel.Left: OldSel(1) = shpSel.Top
    OldSel(2) = shpSel.Width: OldSel(3) = shpSel.Height
    shpSel.Visible = True
    MovingSel = False
 Case tDraw
 'If Int(x / mZoom) * mZoom = X1 And Int(y / mZoom) * mZoom = Y1 Then Exit Sub
  X2 = Int(x / mZoom) * mZoom
  Y2 = Int(y / mZoom) * mZoom
  
  picEdit.Line (X2, Y2)-(X2 + mZoom, Y2 + mZoom), mColor(mButton), BF
  Set picEdit.Picture = picEdit.Image
  picIcon(mCurrIcon).PSet (Int(X2 / mZoom), Int(Y2 / mZoom)), mColor(mButton)
 Case tEye
  If picIcon(mCurrIcon).Point(Int(x / mZoom), Int(y / mZoom)) < 0 Then Exit Sub
  mColor(Button) = picIcon(mCurrIcon).Point(Int(x / mZoom), Int(y / mZoom))
  picClr(Button).BackColor = mColor(Button)
  Call ShowRGBVal(Button)
  Call GradientClr(picClrSel(1), vbBlack, mColor(Button))
  Call GradientClr(picClrSel(2), mColor(Button), vbWhite)
 Case tLine, tLineX2, tLineX3
  Call Tool_Line_Move(x, y)
 Case tBoxGrad
  drawing = True
  Call Tool_Box_Move(x, y, 1)
 Case tBoxGradNS
  drawing = True
  Call Tool_Box_Move(x, y, 2)
 Case tBoxGradNESW
  drawing = True
  Call Tool_Box_Move(x, y, 4)
 Case tBoxGradNWSE
  drawing = True
  Call Tool_Box_Move(x, y, 3)
 Case tCircleGrad
  drawing = True
  Call Tool_Circle_Move(x, y)
 Case tBox To tBoxFill, tBoxFillX
  drawing = True
  Call Tool_Box_Move(x, y)
 Case tcustom
  If tCustomType = 1 Then
   drawing = True
   Call Tool_Box_Move(x, y)
  ElseIf tCustomType = 2 Then
  X2 = Int(x / mZoom) * mZoom
  Y2 = Int(y / mZoom) * mZoom
  Dim s As String, a As String, arr() As String, arrx() As String, v As Variant
  s = mCustomTool
  arrx() = Split(s, vbCrLf)
  For Each v In arrx()
  s = v
  Call Randomize
  s = Replace(s, "wrnd", Int(Int((X2 / mZoom) - Int(X1 / mZoom)) * Rnd))
  Call Randomize
  s = Replace(s, "hrnd", Int(Int((Y2 / mZoom) - Int(Y1 / mZoom)) * Rnd))
  s = Replace(s, "h", Int(Y2 / mZoom) - Int(Y1 / mZoom))
  s = Replace(s, "w", Int(X2 / mZoom) - Int(X1 / mZoom))
  s = Replace(s, "x1", Int(X1 / mZoom))
  s = Replace(s, "x2", Int(X2 / mZoom))
  s = Replace(s, "y1", Int(Y1 / mZoom))
  s = Replace(s, "y2", Int(Y2 / mZoom))
  If InStr(s, " ") <> 0 Then
  a = Left(s, InStr(s, " ") - 1)
  arr() = Split(Mid(s, InStr(s, " ") + 1), ",")
  Select Case LCase(Trim(a))
   Case "line"
    picIcon(mCurrIcon).Line (retMath(arr(0)), retMath(arr(1)))-(retMath(arr(2)), retMath(arr(3))), mColor(Button)
    picIcon(mCurrIcon).PSet (retMath(arr(2)), retMath(arr(3))), mColor(Button)
   Case "dot"
    picIcon(mCurrIcon).PSet (retMath(arr(0)), retMath(arr(1))), mColor(Button)
   Case "box"
    picIcon(mCurrIcon).Line (retMath(arr(0)), retMath(arr(1)))-(retMath(arr(2)), retMath(arr(3))), mColor(Button), B
   Case "filledbox"
    picIcon(mCurrIcon).Line (retMath(arr(0)), retMath(arr(1)))-(retMath(arr(2)), retMath(arr(3))), mColor(Button), BF
   Case "circle"
    picIcon(mCurrIcon).Circle (retMath(arr(0)), retMath(arr(1))), retMath(arr(2)), mColor(Button)
  End Select
  End If
  Next v
  picIcon(mCurrIcon).Refresh
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  picEdit.Cls
  Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
  Call DrawGrid
  picEdit.Refresh
  Set picEdit.Picture = picEdit.Image
  
  End If
 Case tCircle To tCircleFill, tCircleFillX
  drawing = True
  Call Tool_Circle_Move(x, y)
End Select

picEdit.Refresh
End Sub

Private Sub picEdit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case SelTool
 Case tSelect
  If MovingSel = False Then
   MovingSel = True
  Else
    xz1 = shpSel.Left - mZoom + 1
    yz1 = shpSel.Top - mZoom + 1
    If xz1 <= -9000 Or yz1 <= -9000 Then Exit Sub
    picTemp.Width = Int(shpSel.Width / mZoom)
    picTemp.Height = Int(shpSel.Height / mZoom)

    picTemp.PaintPicture picIcon(mCurrIcon).Picture, 0, 0, picTemp.Width, picTemp.Height, Int(OldSel(0) / mZoom), Int(OldSel(1) / mZoom), Int(OldSel(2) / mZoom), Int(OldSel(3) / mZoom), vbSrcCopy
    Set picTemp.Picture = picTemp.Image
    If KeyDown <> 19 Then picIcon(mCurrIcon).Line (Int(OldSel(0) / mZoom), Int(OldSel(1) / mZoom))-(Int(OldSel(0) / mZoom) + Int(OldSel(2) / mZoom) - 1, Int(OldSel(1) / mZoom) + Int(OldSel(3) / mZoom) - 1), picTrans.BackColor, BF
    picIcon(mCurrIcon).PaintPicture picTemp.Picture, Int(shpSel.Left / mZoom), Int(shpSel.Top / mZoom), Int(shpSel.Width / mZoom), Int(shpSel.Height / mZoom), 0, 0, picTemp.Width, picTemp.Height, vbSrcCopy
    picEdit.Refresh
    
    OldSel(0) = shpSel.Left: OldSel(1) = shpSel.Top
    OldSel(2) = shpSel.Width: OldSel(3) = shpSel.Height
  End If
  'Exit Sub
 Case tLine, tLineX2 To tLineX3
  picIcon(mCurrIcon).Line (Int(X1 / mZoom), Int(Y1 / mZoom))-(Int(X2 / mZoom), Int(Y2 / mZoom)), mColor(Button)
  picIcon(mCurrIcon).PSet (Int(X2 / mZoom), Int(Y2 / mZoom)), mColor(Button)
  
  If SelTool >= tLineX2 Then
   If Abs(X2 - X1) > Abs(Y2 - Y1) Then
    picIcon(mCurrIcon).Line (Int(X1 / mZoom), Int(Y1 / mZoom) + 1)-(Int(X2 / mZoom), Int(Y2 / mZoom) + 1), mColor(Button)
    picIcon(mCurrIcon).PSet (Int(X2 / mZoom), Int(Y2 / mZoom) + 1), mColor(Button)
   Else
    picIcon(mCurrIcon).Line (Int(X1 / mZoom) + 1, Int(Y1 / mZoom))-(Int(X2 / mZoom) + 1, Int(Y2 / mZoom)), mColor(Button)
    picIcon(mCurrIcon).PSet (Int(X2 / mZoom) + 1, Int(Y2 / mZoom)), mColor(Button)
   End If
  End If

  If SelTool >= tLineX3 Then
   If Abs(X2 - X1) > Abs(Y2 - Y1) Then
    picIcon(mCurrIcon).Line (Int(X1 / mZoom), Int(Y1 / mZoom) - 1)-(Int(X2 / mZoom), Int(Y2 / mZoom) - 1), mColor(Button)
    picIcon(mCurrIcon).PSet (Int(X2 / mZoom), Int(Y2 / mZoom) - 1), mColor(Button)
   Else
    picIcon(mCurrIcon).Line (Int(X1 / mZoom) - 1, Int(Y1 / mZoom))-(Int(X2 / mZoom) - 1, Int(Y2 / mZoom)), mColor(Button)
    picIcon(mCurrIcon).PSet (Int(X2 / mZoom) - 1, Int(Y2 / mZoom)), mColor(Button)
   End If
  End If

 Case tFill
   Call FillRegion(x, y)
 Case tBox To tBoxFill
   If drawing = False Then Exit Sub
   Call Tool_Box_Up(x, y, CBool(SelTool - 3))
 Case tBoxFillX
   If drawing = False Then Exit Sub
   Call Tool_Box_Up(x, y, True, True)
 Case tBoxGrad, tBoxGradNS, tBoxGradNESW, tBoxGradNWSE
  If drawing = False Then Exit Sub
  If SelTool = tBoxGrad Then
   Call Tool_BoxGrad_Up(x, y, gEW)
  ElseIf SelTool = tBoxGradNS Then
   Call Tool_BoxGrad_Up(x, y, gNS)
  ElseIf SelTool = tBoxGradNWSE Then
   Call Tool_BoxGrad_Up(x, y, gNWSE)
  ElseIf SelTool = tBoxGradNESW Then
   Call Tool_BoxGrad_Up(x, y, gNESW)
  End If
 Case tCircleGrad
  If drawing = False Then Exit Sub
  Call Tool_CircleGrad_Up(x, y)
 Case tCircle To tCircleFill
  If drawing = False Then Exit Sub
  Call Tool_Circle_Up(x, y, CBool(SelTool - 5))
 Case tCircleFillX
  If drawing = False Then Exit Sub
  Call Tool_Circle_Up(x, y, True, True)
 Case tcustom 'custom tool
  If tCustomType = 1 Then
  X2 = Int(x / mZoom) * mZoom
  Y2 = Int(y / mZoom) * mZoom
  Dim s As String, a As String, arr() As String, arrx() As String, v As Variant
  s = mCustomTool
  arrx() = Split(s, vbCrLf)
  For Each v In arrx()
  s = v
  Call Randomize
  s = Replace(s, "wrnd", Int(Int((X2 / mZoom) - Int(X1 / mZoom)) * Rnd))
  Call Randomize
  s = Replace(s, "hrnd", Int(Int((Y2 / mZoom) - Int(Y1 / mZoom)) * Rnd))
  s = Replace(s, "w", Int(X2 / mZoom) - Int(X1 / mZoom))
  s = Replace(s, "h", Int(Y2 / mZoom) - Int(Y1 / mZoom))
  s = Replace(s, "x1", Int(X1 / mZoom))
  s = Replace(s, "x2", Int(X2 / mZoom))
  s = Replace(s, "y1", Int(Y1 / mZoom))
  s = Replace(s, "y2", Int(Y2 / mZoom))
  'MsgBox s
  If InStr(s, " ") <> 0 Then
  a = Left(s, InStr(s, " ") - 1)
  arr() = Split(Mid(s, InStr(s, " ") + 1), ",")
  Select Case LCase(Trim(a))
   Case "line"
    picIcon(mCurrIcon).Line (retMath(arr(0)), retMath(arr(1)))-(retMath(arr(2)), retMath(arr(3))), mColor(Button)
    picIcon(mCurrIcon).PSet (retMath(arr(2)), retMath(arr(3))), mColor(Button)
   Case "dot"
    picIcon(mCurrIcon).PSet (retMath(arr(0)), retMath(arr(1))), mColor(Button)
   Case "box"
    picIcon(mCurrIcon).Line (retMath(arr(0)), retMath(arr(1)))-(retMath(arr(2)), retMath(arr(3))), mColor(Button), B
   Case "filledbox"
    picIcon(mCurrIcon).Line (retMath(arr(0)), retMath(arr(1)))-(retMath(arr(2)), retMath(arr(3))), mColor(Button), BF
   Case "circle"
    picIcon(mCurrIcon).Circle (retMath(arr(0)), retMath(arr(1))), retMath(arr(2)), mColor(Button)
  End Select
  End If
  Next v
  End If
  End Select
drawing = False
picIcon(mCurrIcon).Refresh
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
X1 = -1: Y1 = -1
FileChanged(mCurrIcon) = True
End Sub

Private Sub picEdit_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo 1
Dim ext As String, FileName As String
FileName = Data.Files(1)
ext = Right(FileName, 4)
If ext = ".ico" Then
Dim i As Integer
i = GetIconSize(FileName)

Call AddRecent(FileName)
If i <> -1 Then Call NewIcon(i) Else Call NewIcon(32)
Set picUndo(mCurrIcon).Picture = LoadPicture(FileName)
Set picIcon(mCurrIcon).Picture = LoadPicture(FileName)
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picIcon(mCurrIcon).Refresh
picIcon(mCurrIcon).Tag = FileName
txtIcon(mCurrIcon).Caption = ExtractFileName(FileName) & vbCrLf & picIcon(mCurrIcon).Width & "x" & picIcon(mCurrIcon).Height
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
End If
1
End Sub

Private Sub picIcon_Click(Index As Integer)
Dim i As Integer
mCurrIcon = Index
picIcon(mCurrIcon).Refresh
Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
picEdit.Cls
picEdit.Width = picIcon(mCurrIcon).Width * mZoom
picEdit.Height = picIcon(mCurrIcon).Height * mZoom
Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
Call DrawGrid
picEdit.Refresh
Set picEdit.Picture = picEdit.Image
picEdit.Visible = True
End Sub

Private Sub picIconsBack_Resize()
vsIcons.Left = 0
vsIcons.Top = picIconsBack.Height - vsIcons.Height - 4
vsIcons.Width = picIconsBack.Width - 4
End Sub

Private Sub picIconsMove_Resize()
vsIcons.Max = picIconsMove.Height - picIconsBack.Height
End Sub

Private Sub picToolSel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer, j As Integer
Dim l As Long
Dim arr() As String
     
shpToolSel(Index).Visible = True
shpToolSel(Index).Top = Int(y / 20) * 20
shpToolSel(Index).Left = Int(x / 20) * 20
  
Select Case Index
 Case 1
 i = Int(y / 20)
  Select Case i
   Case 0
    SelTool = tDraw
   Case 1
   On Error GoTo 1
    cd.Filter = "Custom Pens (*.ccp)|*.ccp"
    If mCustomTool = "" Then cd.FileName = App.Path
    cd.ShowOpen
    SelTool = tcustom
     l = FreeFile
    Open cd.FileName For Input As #l
     mCustomTool = Input(LOF(1), #1)
    Close #l
     arr() = Split(mCustomTool, vbCrLf)
     Select Case LCase(arr(0))
      Case "box"
       tCustomType = 1
      Case "draw"
       tCustomType = 2
      Case Else
       tCustomType = -1
     End Select
     mCustomTool = ""
     For i = 1 To UBound(arr())
      mCustomTool = mCustomTool & arr(i) & vbCrLf
     Next i
1
  End Select
 Case 3
 i = Int(y / 20)
 j = Int(x / 20)
  Select Case i
   Case 0
    If j = 0 Then
     SelTool = tBox
    Else
     SelTool = tBoxFill
    End If
   Case 1
    If j = 0 Then
     SelTool = tBoxFillX
    Else
   On Error GoTo 2
    cd.Filter = "Custom Boxes (*.ccb)|*.ccb"
    If mCustomTool = "" Then cd.FileName = App.Path
    cd.ShowOpen
    SelTool = tcustom
     l = FreeFile
    Open cd.FileName For Input As #l
     mCustomTool = Input(LOF(1), #1)
    Close #l
     arr() = Split(mCustomTool, vbCrLf)
     Select Case LCase(arr(0))
      Case "box"
       tCustomType = 1
      Case "draw"
       tCustomType = 2
      Case Else
       tCustomType = -1
     End Select
     mCustomTool = ""
     For i = 1 To UBound(arr())
      mCustomTool = mCustomTool & arr(i) & vbCrLf
     Next i
2
    End If
   Case 2
    If j = 0 Then
     SelTool = tBoxGrad
    Else
     SelTool = tBoxGradNESW
    End If
   Case 3
    If j = 0 Then
     SelTool = tBoxGradNS
    Else
     SelTool = tBoxGradNWSE
    End If
  End Select
 Case 5
 i = Int(y / 20)
  Select Case i
   Case 0
    SelTool = tCircle
   Case 1
    SelTool = tCircleFill
   Case 2
    SelTool = tCircleFillX
   Case 3
    SelTool = tCircleGrad
  End Select
 Case 10
 i = Int(y / 20)
  Select Case i
   Case 0
    SelTool = tLine
   Case 1
    SelTool = tLineX2
   Case 2
    SelTool = tLineX3
  End Select
End Select
End Sub

Private Sub picToolSel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer, j As Integer
Dim l As Long
Dim arr() As String

 
Select Case Index
 Case 1
 i = Int(y / 20)
  Select Case i
   Case 0
    picToolSel(Index).ToolTipText = "Pen Tool"
   Case 1
    picToolSel(Index).ToolTipText = "Custom Pen Tool"
  End Select
 Case 3
 i = Int(y / 20)
 j = Int(x / 20)
  Select Case i
   Case 0
    If j = 0 Then
     picToolSel(Index).ToolTipText = "Box Tool"
    Else
     picToolSel(Index).ToolTipText = "Filled Box"
    End If
   Case 1
    If j = 0 Then
     picToolSel(Index).ToolTipText = "Filled Box w/ Border"
    Else
     picToolSel(Index).ToolTipText = "Custom Box Tool"
    End If
   Case 2
    If j = 0 Then
     picToolSel(Index).ToolTipText = "Gradient Filled Box EW"
    Else
     picToolSel(Index).ToolTipText = "Gradient Filled Box NESW"
    End If
   Case 3
    If j = 0 Then
     picToolSel(Index).ToolTipText = "Gradient Filled Box NS"
    Else
     picToolSel(Index).ToolTipText = "Gradient Filled Box NWSE"
    End If
  End Select
 Case 5
 i = Int(y / 20)
  Select Case i
   Case 0
    picToolSel(Index).ToolTipText = "Circle Tool"
   Case 1
    picToolSel(Index).ToolTipText = "Filled Circle"
   Case 2
    picToolSel(Index).ToolTipText = "Filled Circle w/ Border"
   Case 3
    picToolSel(Index).ToolTipText = "Gradient Filled Circle"
  End Select
 Case 10
 i = Int(y / 20)
  Select Case i
   Case 0
    picToolSel(Index).ToolTipText = "Line Tool"
   Case 1
    picToolSel(Index).ToolTipText = "Thick Line Tool"
   Case 2
    picToolSel(Index).ToolTipText = "Thicker Line Tool"
  End Select
End Select

End Sub

Private Sub picTrans_Click()
picClr(1).BackColor = picTrans.BackColor
mColor(1) = picClr(1).BackColor
Dim s As String
s = GetRGB(mColor(1)).Red & "," & GetRGB(mColor(1)).Green & "," & GetRGB(mColor(1)).Blue
Clipboard.Clear
Clipboard.SetText s
End Sub

Private Sub tbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
picEdit.Cls
shpSel.Visible = False
shpSel.Left = -1000
shpSel.Top = -1000

If Button.Tag = tText Then
  Call Load(frmText)
  Set frmText.picTemp.Picture = picIcon(mCurrIcon).Picture
  Call frmText.Show(vbModal)
  Exit Sub
End If
SelTool = Button.Tag
Dim b As MSComctlLib.Button
For Each b In tbTools.Buttons
 b.Value = tbrUnpressed
Next b
Button.Value = tbrPressed
tbTools.Refresh
On Error Resume Next
Dim i As Integer
For i = 0 To 10
 picToolSel(i).Visible = False
Next i

shpToolSel(SelTool).Width = 20
shpToolSel(SelTool).Left = 0
shpToolSel(SelTool).Height = 20
shpToolSel(SelTool).Top = 0
shpToolSel(SelTool).Visible = True
picToolSel(SelTool).Visible = True
picToolSel(SelTool).Left = (picExtra.Width / 2) - (picToolSel(SelTool).Width / 2) - 4
End Sub

Private Sub vsEdit_Change()
picEdit.Top = vsEdit.Value + 8
End Sub

Private Sub vsIcons_Change()
If picIconsMove.Height - picIconsBack.Height > 0 Then picIconsMove.Top = Val("-" & vsIcons.Value) + 8 Else picIconsMove.Top = 8
End Sub

Private Sub wc_Got(ByVal Msg As String)
Select Case Left(Msg, 1)
 Case "@"
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Call SavePicture(picIcon(mCurrIcon).Picture, App.Path & "\temp.bmp")
  wc.mHwnd = Mid(Msg, 2)
  Call wc.Send("$" & App.Path & "\temp.bmp")
 Case "!"
  FileChanged(mCurrIcon) = True
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
  Set picUndo(mCurrIcon).Picture = picIcon(mCurrIcon).Picture
  Set picIcon(mCurrIcon).Picture = LoadPicture(App.Path & "\temp.bmp")
  Dim x As Integer, y As Integer
  For y = 0 To 31
   For x = 0 To 31

    If picIcon(mCurrIcon).Point(x, y) = 8420352 Then picIcon(mCurrIcon).PSet (x, y), picTrans.BackColor
    
   Next x
  Next y
  picIcon(mCurrIcon).Refresh
  Set picIcon(mCurrIcon).Picture = picIcon(mCurrIcon).Image
    picEdit.Cls
    Call picEdit.PaintPicture(picIcon(mCurrIcon).Picture, 0, 0, picIcon(mCurrIcon).Width * mZoom, picIcon(mCurrIcon).Height * mZoom, 0, 0, picIcon(mCurrIcon).Width, picIcon(mCurrIcon).Height, vbSrcCopy)
    Call DrawGrid
    picEdit.Refresh
    Set picEdit.Picture = picEdit.Image
End Select
End Sub


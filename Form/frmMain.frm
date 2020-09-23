VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8205
   ClientLeft      =   2175
   ClientTop       =   2040
   ClientWidth     =   11850
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   11850
   Begin AddErrorChecking.Frame3D fraOptions 
      Height          =   7665
      Left            =   6270
      Top             =   30
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   13520
      BorderType      =   8
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   -2147483633
      CornerDiameter  =   7
      FillColor       =   16761247
      FillStyle       =   1
      DrawStyle       =   0
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":1CFA
      Picture         =   "frmMain.frx":1D16
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   ""
      UseMnemonic     =   -1  'True
      Begin VB.TextBox txtMinCodeLines 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2175
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "4"
         Top             =   7110
         Width           =   495
      End
      Begin VB.TextBox txtExitLabel 
         Height          =   285
         Left            =   1545
         TabIndex        =   17
         Text            =   "Exit_Proc"
         Top             =   4590
         Width           =   1545
      End
      Begin VB.TextBox txtErrLbl 
         Height          =   285
         Left            =   1545
         TabIndex        =   16
         Text            =   "Err_Proc"
         Top             =   4245
         Width           =   1545
      End
      Begin VB.TextBox txtTabLength 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4290
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "3"
         Top             =   4710
         Width           =   450
      End
      Begin VB.TextBox txtUpperGap 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4290
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "1"
         Top             =   4080
         Width           =   450
      End
      Begin VB.TextBox txtLowerGap 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4290
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "1"
         Top             =   4395
         Width           =   450
      End
      Begin VB.CheckBox chkMinCodeLines 
         Caption         =   "Ignore functions with "
         Height          =   255
         Left            =   315
         TabIndex        =   27
         Top             =   7140
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkApplyOnProps 
         Caption         =   "Apply to Properties"
         Height          =   255
         Left            =   285
         TabIndex        =   23
         Top             =   5820
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.TextBox txtControlPrefixes 
         Height          =   285
         Left            =   285
         TabIndex        =   26
         Text            =   "chk,lbl,cbo,lst,txt,opt,img,form_unload,form_initialize"
         Top             =   6705
         Width           =   5055
      End
      Begin VB.CheckBox chkIgnoreControlsPrefix 
         Caption         =   "Ignore procedures starting with:"
         Height          =   255
         Left            =   285
         TabIndex        =   25
         Top             =   6435
         UseMaskColor    =   -1  'True
         Width           =   4725
      End
      Begin VB.CheckBox chkIgnoreOnErr 
         Caption         =   "Ignore procedures with ""ON ERROR"" commands"
         Height          =   255
         Left            =   285
         TabIndex        =   24
         Top             =   6120
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   4725
      End
      Begin VB.CheckBox chkApplyOnFunc 
         Caption         =   "Apply to Functions"
         Height          =   255
         Left            =   285
         TabIndex        =   22
         Top             =   5505
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.CheckBox chkApplyOnProc 
         Caption         =   "Apply to Subs"
         Height          =   255
         Left            =   285
         TabIndex        =   21
         Top             =   5205
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.TextBox txtErrHndl 
         Enabled         =   0   'False
         Height          =   1065
         Left            =   405
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Text            =   "frmMain.frx":1D32
         Top             =   360
         Width           =   4890
      End
      Begin VB.OptionButton optUseFreeText 
         Caption         =   "Use free text"
         DownPicture     =   "frmMain.frx":1D5E
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   105
         MaskColor       =   &H00FFC19F&
         TabIndex        =   4
         Top             =   60
         Width           =   1305
      End
      Begin VB.OptionButton optUseErrFunc 
         Caption         =   "Use Error handling procedure"
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   105
         TabIndex        =   6
         Top             =   1530
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   3285
      End
      Begin VB.TextBox txtFuncName 
         Height          =   285
         Left            =   2115
         TabIndex        =   14
         Text            =   "Err_Handler"
         Top             =   3210
         Width           =   2625
      End
      Begin VB.TextBox txtModName 
         Height          =   285
         Left            =   2115
         TabIndex        =   15
         Text            =   "modLogError"
         Top             =   3555
         Width           =   2625
      End
      Begin AddErrorChecking.Frame3D MyFrame2 
         Height          =   675
         Left            =   2850
         Top             =   1785
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1191
         BorderType      =   8
         BevelWidth      =   3
         BevelInner      =   0
         Caption3D       =   0
         CaptionAlignment=   0
         CaptionLocation =   0
         BackColor       =   -2147483633
         CornerDiameter  =   7
         FillColor       =   16761247
         FillStyle       =   1
         DrawStyle       =   0
         FloodPercent    =   0
         FloodShowPct    =   0   'False
         FloodType       =   0
         FloodColor      =   16761247
         FillGradient    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":2068
         Picture         =   "frmMain.frx":2084
         Border3DHighlight=   -2147483628
         Border3DShadow  =   -2147483632
         Enabled         =   -1  'True
         CaptionMAlignment=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   -2147483630
         Caption         =   ""
         UseMnemonic     =   -1  'True
         Begin VB.OptionButton optShowLog 
            Caption         =   "Log Error Only"
            Height          =   210
            Index           =   1
            Left            =   105
            TabIndex        =   11
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   2190
         End
         Begin VB.OptionButton optShowLog 
            Caption         =   "Show and Log Error"
            Height          =   210
            Index           =   0
            Left            =   105
            TabIndex        =   10
            Top             =   90
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   2190
         End
      End
      Begin AddErrorChecking.Frame3D MyFrame1 
         Height          =   1305
         Left            =   405
         Top             =   1785
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2302
         BorderType      =   8
         BevelWidth      =   3
         BevelInner      =   0
         Caption3D       =   0
         CaptionAlignment=   0
         CaptionLocation =   0
         BackColor       =   -2147483633
         CornerDiameter  =   7
         FillColor       =   16761247
         FillStyle       =   1
         DrawStyle       =   0
         FloodPercent    =   0
         FloodShowPct    =   0   'False
         FloodType       =   0
         FloodColor      =   16761247
         FillGradient    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":20A0
         Picture         =   "frmMain.frx":20BC
         Border3DHighlight=   -2147483628
         Border3DShadow  =   -2147483632
         Enabled         =   -1  'True
         CaptionMAlignment=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   -2147483630
         Caption         =   "Include:"
         UseMnemonic     =   -1  'True
         Begin VB.CheckBox chkProcName 
            Caption         =   "Proc name"
            Height          =   255
            Left            =   840
            TabIndex        =   8
            Top             =   510
            UseMaskColor    =   -1  'True
            Value           =   1  'Checked
            Width           =   1185
         End
         Begin VB.CheckBox chkModuleName 
            Caption         =   "Module name"
            Height          =   255
            Left            =   840
            TabIndex        =   7
            Top             =   255
            UseMaskColor    =   -1  'True
            Value           =   1  'Checked
            Width           =   1305
         End
         Begin VB.CheckBox chkErrObj 
            Caption         =   "Err Info"
            Height          =   255
            Left            =   840
            MaskColor       =   &H8000000F&
            TabIndex        =   9
            Top             =   765
            UseMaskColor    =   -1  'True
            Value           =   1  'Checked
            Width           =   1080
         End
      End
      Begin AddErrorChecking.Frame3D MyFrame3 
         Height          =   660
         Left            =   2850
         Top             =   2430
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1164
         BorderType      =   8
         BevelWidth      =   3
         BevelInner      =   0
         Caption3D       =   0
         CaptionAlignment=   0
         CaptionLocation =   0
         BackColor       =   -2147483633
         CornerDiameter  =   7
         FillColor       =   16761247
         FillStyle       =   1
         DrawStyle       =   0
         FloodPercent    =   0
         FloodShowPct    =   0   'False
         FloodType       =   0
         FloodColor      =   16761247
         FillGradient    =   0
         MousePointer    =   0
         MouseIcon       =   "frmMain.frx":20D8
         Picture         =   "frmMain.frx":20F4
         Border3DHighlight=   -2147483628
         Border3DShadow  =   -2147483632
         Enabled         =   -1  'True
         CaptionMAlignment=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   -2147483630
         Caption         =   ""
         UseMnemonic     =   -1  'True
         Begin VB.OptionButton optResume 
            Caption         =   "Resume Err Label"
            Height          =   210
            Index           =   0
            Left            =   105
            TabIndex        =   12
            Top             =   90
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   2205
         End
         Begin VB.OptionButton optResume 
            Caption         =   "Resume Next"
            Height          =   210
            Index           =   1
            Left            =   105
            TabIndex        =   13
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   2205
         End
      End
      Begin VB.Label lblMiscInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exit goto label:"
         Height          =   255
         Index           =   8
         Left            =   315
         TabIndex        =   3
         Top             =   4620
         Width           =   1185
      End
      Begin VB.Label lblMiscInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Error goto label:"
         Height          =   255
         Index           =   4
         Left            =   315
         TabIndex        =   29
         Top             =   4275
         Width           =   1185
      End
      Begin VB.Label lblMiscInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Tab length: "
         Height          =   195
         Index           =   6
         Left            =   3390
         TabIndex        =   30
         Top             =   4740
         Width           =   900
      End
      Begin VB.Label lblMiscInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Lower gap: "
         Height          =   195
         Index           =   5
         Left            =   3405
         TabIndex        =   31
         Top             =   4410
         Width           =   885
      End
      Begin VB.Label lblMiscInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Upper gap: "
         Height          =   195
         Index           =   3
         Left            =   3420
         TabIndex        =   35
         Top             =   4095
         Width           =   885
      End
      Begin VB.Label lblMiscInfo 
         BackStyle       =   0  'Transparent
         Caption         =   " or less lines of code."
         Height          =   210
         Index           =   1
         Left            =   2715
         TabIndex        =   34
         Top             =   7155
         Width           =   2010
      End
      Begin VB.Line linOptions 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   15
         X2              =   5520
         Y1              =   5100
         Y2              =   5100
      End
      Begin VB.Line linOptions 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   15
         X2              =   5520
         Y1              =   5085
         Y2              =   5085
      End
      Begin VB.Line linOptions 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   15
         X2              =   5520
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line linOptions 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   15
         X2              =   5520
         Y1              =   3975
         Y2              =   3975
      End
      Begin VB.Label lblMiscInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Error Sub Name: "
         Height          =   195
         Index           =   7
         Left            =   855
         TabIndex        =   33
         Top             =   3240
         Width           =   1260
      End
      Begin VB.Label lblMiscInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Error Module Name: "
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   32
         Top             =   3585
         Width           =   1500
      End
   End
   Begin MSComctlLib.TreeView trvModules 
      Height          =   7335
      Left            =   45
      TabIndex        =   0
      Top             =   195
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   12938
      _Version        =   393217
      Indentation     =   882
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList imglst1 
      Left            =   12000
      Top             =   2910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2110
            Key             =   "FOLDER"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2562
            Key             =   "Designer"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28B4
            Key             =   "Classes"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C06
            Key             =   "Forms"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F58
            Key             =   "Modules"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32AA
            Key             =   "User controls"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35FC
            Key             =   "User documents"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":370E
            Key             =   "ROOT"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A60
            Key             =   "SUB"
         EndProperty
      EndProperty
   End
   Begin AddErrorChecking.Frame3D picToolBar 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      Top             =   7710
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   873
      BorderType      =   0
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   -2147483633
      CornerDiameter  =   7
      FillColor       =   -2147483632
      FillStyle       =   1
      DrawStyle       =   0
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   2
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":3DB2
      Picture         =   "frmMain.frx":3DCE
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   ""
      UseMnemonic     =   0   'False
      Begin VB.CommandButton cmdCollapse 
         Caption         =   "Collapse Tree"
         Height          =   375
         Left            =   4080
         TabIndex        =   44
         Top             =   60
         Width           =   1410
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   375
         Left            =   10575
         TabIndex        =   43
         ToolTipText     =   "End Program"
         Top             =   60
         Width           =   1230
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   375
         Left            =   9675
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   60
         Width           =   870
      End
      Begin VB.CommandButton cmdBrows 
         Caption         =   "Open"
         Default         =   -1  'True
         Height          =   375
         Left            =   90
         TabIndex        =   41
         ToolTipText     =   "Add vb project to the list"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Transfer"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1710
         TabIndex        =   40
         ToolTipText     =   "Replace the original files"
         Top             =   60
         Width           =   1125
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Close"
         Height          =   375
         Left            =   2880
         TabIndex        =   39
         ToolTipText     =   "Clear list"
         Top             =   60
         Width           =   795
      End
      Begin VB.CommandButton cmdCommit 
         Caption         =   "Commit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   870
         TabIndex        =   38
         ToolTipText     =   "Add error handling to the selected files and functions"
         Top             =   60
         Width           =   795
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View Code"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   37
         ToolTipText     =   "Compare original file with the new file"
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdShowInterface 
         Caption         =   "View Summary"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6765
         TabIndex        =   36
         ToolTipText     =   "Show selected file interface"
         Top             =   60
         Width           =   1425
      End
   End
   Begin VB.Label lblProcedureCount 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      ForeColor       =   &H80000014&
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   7515
      UseMnemonic     =   0   'False
      Width           =   6180
   End
   Begin VB.Label lblMiscInfo 
      BackColor       =   &H80000010&
      Caption         =   "Code modules"
      ForeColor       =   &H80000014&
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   6165
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//**********************************************/
'//     Authors: Morgan Haueisen and Adi barda  */
'//             morganh@hartcom.net             */
'//     Copyright (c) 2003-2004                 */
'//**********************************************/

'Legal:
'        This is intended for and was uploaded to www.planetsourcecode.com
'
'        Redistribution of this code, whole or in part, as source code or in binary form, alone or
'        as part of a larger distribution or product, is forbidden for any commercial or for-profit
'        use without the author's explicit written permission.
'
'        Redistribution of this code, as source code or in binary form, with or without
'        modification, is permitted provided that the following conditions are met:
'
'        Redistributions of source code must include this list of conditions, and the following
'        acknowledgment:
'
'        This code was developed by Adi barda and Morgan Haueisen.  <morganh@hartcom.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'        Redistributions in binary form, as part of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.

Option Explicit

Private Const C_MODULE_NAME As Integer = 0
Private Const C_PROC_NAME   As Integer = 1

Private mstrAControlsPrefix() As String
Private mstrProjectName       As String
Private mblnAlwaysIgnored()   As Boolean
Private mlngModuleCount       As Long
Private mstrChangeSummary()   As String
Private mstrErrorModulePath   As String

Private mcFile      As clsFileUtilities
Private mcTextBx    As clsTextBox

Private Function AddErrHandling(ByVal vstrFilePath As String, _
                                Optional ByVal vblnCreateTempFiles As Boolean = False) As Boolean

   '// Purpose: Add error handling to the temporary file
   '//          if vblnCreateTempFiles = true

  Const C_PROCESS_REMARK As String = "'"

  Dim lngSourceFN           As Long    '// source file number
  Dim lngTempFN             As Long    '// dest file number

  Dim strLine               As String  '// source read line
  Dim strLineTrimmed        As String  '// source read line trimmed
  Dim strDest               As String  '// temp line write

  Dim strModuleName         As String  '// current module's name
  Dim strModuleFileName     As String  '// current module's file name
  Dim strProcName           As String  '// current procedure name
  Dim lngLineCount          As Long    '// Code line count in a prodedure

  Dim blnStartSub           As Boolean '// recognize sub start
  Dim strEndSub             As String  '// recognize end of sub or function
  Dim blnClassFile          As Boolean '// Class file, added private error sub

  Dim blnAddOnErr           As Boolean '// flag saying need to add on error statement
  Dim blnModuleOnErrAdded   As Boolean '// global flag saying added on error statement in this module
  Dim blnOnErrorAdded       As Boolean '// flag indicated whether on error added
  Dim blnHasOnError         As Boolean '// the function already has on error
  Dim blnHasOnErrorSub      As Boolean '// the module already has on error sub
  Dim blnFoundModuleName    As Boolean '// flag-found the module name
  Dim blnIgnoreSub          As Boolean '// Ignore sub based on parameters choosen
  Dim lngTVIndex            As Long    '// Current node index
  Dim blnHasExitLabel       As Boolean '// Procedure already has exit label
  Dim blnHasErrLabel        As Boolean '// Procedure already has error label
  Dim blnLineContinue       As Boolean '// is the line a continuation of the previous line?

   On Error GoTo Err_Proc

   '// init algorithm flags
   blnStartSub = False
   strEndSub = vbNullString
   blnAddOnErr = False
   blnModuleOnErrAdded = False
   blnHasOnErrorSub = False
   blnOnErrorAdded = False
   blnFoundModuleName = False

   '// init vars
   strModuleName = vbNullString
   strProcName = vbNullString
   lngLineCount = -1
   strModuleFileName = mcFile.RetOnlyFilename(vstrFilePath)

   If vblnCreateTempFiles Then

      '// Store Error module as private inside module?
      blnClassFile = (mcFile.GetExtensionName(strModuleFileName) = "cls" _
         Or mcFile.GetExtensionName(strModuleFileName) = "dob")

      '// ensures that the temp files folder exists

      If Not mcFile.FolderExists(App.Path & "\DestTmp") Then
         mcFile.CreateDir App.Path & "\DestTmp"
      End If

      frmStatus.lblStatusText.Caption = App.Path & "\DestTmp\" & strModuleFileName
      frmStatus.lblStatusText.Refresh

      '// open temp dest file
      lngTempFN = FreeFile
      Open App.Path & "\DestTmp\" & strModuleFileName For Output As #lngTempFN

   Else
      frmStatus.lblStatusText.Caption = vstrFilePath
      frmStatus.lblStatusText.Refresh
   End If

   '// Open Source File for reading
   lngSourceFN = FreeFile
   Open vstrFilePath For Input As #lngSourceFN

   '// Main scanning loop

   Do Until EOF(lngSourceFN)

      '// read the current line from the file
      Line Input #lngSourceFN, strLine

      '// init dest line
      strDest = vbNullString
      strLineTrimmed = Trim$(strLine)

      '//  Check for the module name

      If Not blnFoundModuleName Then
         strModuleName = GetModuleName(strLine)
         blnFoundModuleName = (LenB(strModuleName) <> 0)
      End If

      '//  Check if it is the Begining or Ending of a procedure
      Call CheckForStartEnd(strLine, _
            strModuleFileName, _
            strProcName, _
            blnStartSub, _
            strEndSub, _
            blnHasOnError, _
            blnHasOnErrorSub, _
            blnHasExitLabel, _
            blnHasErrLabel, _
            lngLineCount, _
            vblnCreateTempFiles)

      '// Count lines of code

      If LenB(strLineTrimmed) And Not blnLineContinue Then
         If blnStartSub And LenB(strEndSub) = 0 _
               And Left$(strLineTrimmed, 1) <> C_PROCESS_REMARK _
               And Left$(strLineTrimmed, 4) <> "Dim " _
               And Left$(strLineTrimmed, 6) <> "Const " _
               And Left$(strLineTrimmed, 7) <> "Static " Then

            lngLineCount = lngLineCount + 1
         End If

      End If
      blnLineContinue = CBool(Right$(strLineTrimmed, 1) = "_")

      '//  Check if the function already has on error statement

      If Not blnHasOnError Then
         blnHasOnError = (InStrB(LCase$(strLine), "on error ") > 0)
      End If

      If Not blnHasOnError Then
         blnHasOnError = (InStrB(LCase$(strLine), "on local error ") > 0)
      End If

      If Not blnHasExitLabel Then
         blnHasExitLabel = (InStrB(LCase$(strLine), LCase$(txtExitLabel.Text)) > 0)
      End If

      If Not blnHasErrLabel Then
         blnHasErrLabel = (InStrB(LCase$(strLine), LCase$(txtErrLbl.Text)) > 0)
      End If

      If vblnCreateTempFiles Then

         If blnStartSub And blnAddOnErr And Not blnOnErrorAdded Then
            If LenB(strLineTrimmed) Then
               If LenB(strEndSub) = 0 _
                     And Left$(strLineTrimmed, 1) <> C_PROCESS_REMARK _
                     And Left$(strLineTrimmed, 4) <> "Dim " _
                     And Left$(strLineTrimmed, 6) <> "Const " _
                     And Left$(strLineTrimmed, 7) <> "Static " Then
                  
                  If Right$(strProcName, 7) = "_Unload" Then
                     strDest = strDest & Space$(CLng(txtTabLength.Text)) & "On Error Resume Next" & vbNewLine
                     strDest = strDest & String$(CLng(txtUpperGap.Text), vbNewLine)
                     strDest = strDest & strLine
                  
                  Else
                     '// If its the error handling proc don't insert any error handling code
                     '// Add on error goto...
                     strDest = strDest & Space$(CLng(txtTabLength.Text)) & "On Error GoTo " & txtErrLbl.Text & vbNewLine
                     strDest = strDest & String$(CLng(txtUpperGap.Text), vbNewLine)
                     strDest = strDest & strLine
                  End If
                  
                  blnOnErrorAdded = True
                  blnModuleOnErrAdded = True

                  If LenB(strProcName) Then
                     mstrChangeSummary(C_MODULE_NAME, UBound(mstrChangeSummary, 2)) = strModuleName
                     mstrChangeSummary(C_PROC_NAME, UBound(mstrChangeSummary, 2)) = strProcName
                     ReDim Preserve mstrChangeSummary(1, UBound(mstrChangeSummary, 2) + 1) As String
                  End If

               End If
            End If
         End If

         If blnStartSub Then
            If Not blnAddOnErr Then
               '// Purpose: Is it Ok to add "On Error Goto" after this line?
               blnAddOnErr = CBool((InStr(strLine, ")") > 0) And (Right$(strLine, 1) <> "_"))
            End If

         End If

      End If

      '//  Check if its end of sub, function, or property

      If LenB(strEndSub) Then
         If vblnCreateTempFiles Then

            If Right$(strProcName, 7) = "_Unload" Then
               strDest = strLine
            
            Else
               '// does not have exit label
               If Not blnHasExitLabel Then
                  strDest = vbNewLine & txtExitLabel.Text & ":" & vbNewLine
                  strDest = strDest & Space$(CLng(txtTabLength.Text)) & strEndSub & vbNewLine
               End If
   
               If Not blnHasErrLabel Then
                  '// Add label text
                  strDest = strDest & vbNewLine
                  strDest = strDest & txtErrLbl.Text & ":" & vbNewLine
   
                  '// Add reference to err handling
                  If optUseFreeText.Value Then
                     strDest = strDest & Space$(CLng(txtTabLength.Text)) & txtErrHndl.Text & vbNewLine
                  Else
                     strDest = strDest & GetErrFunctionConst(strModuleName, strProcName) & vbNewLine
                  End If
   
                  '// Resume to exit point or next
                  If optResume(0).Value Then
                     strDest = strDest & Space$(CLng(txtTabLength.Text)) & "Resume " & txtExitLabel.Text & vbNewLine
                  Else
                     strDest = strDest & Space$(CLng(txtTabLength.Text)) & "Resume Next" & vbNewLine
                  End If
   
                  '// insert lower gap space
                  strDest = strDest & String$(CLng(txtLowerGap.Text), vbNewLine)
               End If
   
               '// Add orignal line
               strDest = strDest & strLine
            End If
            
         Else

            '// Update functions array: Decide if the procedure gets error code or not
            '// ignore this function if it begins with ??? or has on error
            blnIgnoreSub = (((blnHasOnError) And (chkIgnoreOnErr.Value = vbChecked)) Or (HasControlPrefix(strProcName)))
            If chkMinCodeLines.Value = vbChecked Then '// ignore if there are less then X lines of code
               blnIgnoreSub = (blnIgnoreSub Or (lngLineCount <= Val(txtMinCodeLines.Text)))
            End If

            lngTVIndex = modTreeView.AddNode(mcFile.RetOnlyFilename(vstrFilePath), GetNextKey(), strProcName, , "SUB", Not blnIgnoreSub)
            ReDim Preserve mblnAlwaysIgnored(lngTVIndex) As Boolean
            mblnAlwaysIgnored(lngTVIndex) = blnIgnoreSub

         End If

         '//  Clear variables:
         blnStartSub = False
         strEndSub = vbNullString
         blnAddOnErr = False
         blnOnErrorAdded = False
         strProcName = vbNullString

      End If

      If vblnCreateTempFiles Then
         '// if nessesary insert default value

         If LenB(strDest) = 0 Then
            strDest = strLine
         End If

         '// prints to destination temp file
         Print #lngTempFN, strDest
      End If

   Loop

   '//  If nessesary, insert private error handling function into class modules

   If vblnCreateTempFiles Then
      If blnClassFile And blnModuleOnErrAdded And Not blnHasOnErrorSub Then
         Print #lngTempFN, vbNewLine & WriteErrFunction()
      End If

   End If

Exit_Proc:
   '// Close file ports
   Close #lngSourceFN

   If vblnCreateTempFiles Then
      Close #lngTempFN
   End If

   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "AddErrHandling"
   Err.Clear
   Resume Exit_Proc

End Function

Private Sub AddFileToList(ByVal vstrFilePath As String, _
                          ByVal vstrFileName As String)

  Dim strFileType     As String
  Dim strFileNameOnly As String

   On Error GoTo Err_Proc

   If Dir$(vstrFilePath & vstrFileName) > vbNullString Then

      With frmStatus
         .lblStatusText.Caption = vstrFileName
         .lblStatusText.Refresh
      End With

      strFileNameOnly = mcFile.RetOnlyFilename(vstrFilePath & vstrFileName)

      '// Add file to list
      strFileType = mcFile.GetExtensionName(strFileNameOnly)

      Select Case strFileType
      Case "frm"
         strFileType = "Forms"

      Case "cls"
         strFileType = "Classes"

      Case "bas"
         strFileType = "Modules"

      Case "ctl"
         strFileType = "User controls"

      Case "dob"
         strFileType = "User documents"

      Case "dsr"
         strFileType = "Designer"
      End Select

      If LenB(vstrFilePath & vstrFileName) > 0 Then
         If modTreeView.FindTag(vstrFilePath & vstrFileName) = 0 Then
            modTreeView.AddNode strFileType, strFileNameOnly, strFileNameOnly, _
               vstrFilePath & vstrFileName, strFileType
         End If

      End If

   Else
      frmMsgBox.SMessageModal vstrFileName & " file is missing.", vbInformation + vbOkButton
   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "AddFileToList"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub AddProject(ByVal vstrProjectFileName As String)

   '// Purpose: adds the selected project (all its files) to the system manager

  Dim lngFN         As Long
  Dim lngI          As Long
  Dim strLine       As String
  Dim strFileName   As String
  Dim strProjPath   As String

   On Error GoTo Err_Proc

   '// Add project name to list
   trvModules.Nodes(1).Text = mcFile.RetOnlyFilename(vstrProjectFileName)
   trvModules.Nodes(1).Checked = True

   '// ensures backslash is exists
   strProjPath = mcFile.RetOnlyPath(vstrProjectFileName)

   '// open file port
   lngFN = FreeFile
   Open vstrProjectFileName For Input As #lngFN

   '// scan vb project file

   Do Until EOF(lngFN)
      Line Input #lngFN, strLine '// read next line in the project file

      '// check for the next object:

      If InStrB(1, LCase$(strLine), "form=") Then
         lngI = InStr(1, strLine, "=") + 1
         strFileName = Mid$(strLine, lngI, Len(strLine) - lngI + 1) '// find object name

         '// check that there is no (") in the object name

         If InStrB(1, strFileName, Chr$(34)) = 0 Then
            Call AddFileToList(strProjPath, strFileName)
         End If

      End If

      If InStrB(1, LCase$(strLine), "class=") Then
         lngI = InStr(1, strLine, ";") + 2
         strFileName = Mid$(strLine, lngI, Len(strLine) - lngI + 1)
         Call AddFileToList(strProjPath, strFileName)
      End If

      If InStrB(1, LCase$(strLine), "module=") Then
         lngI = InStr(1, strLine, ";") + 2
         strFileName = Mid$(strLine, lngI, Len(strLine) - lngI + 1)
         Call AddFileToList(strProjPath, strFileName)
      End If

      If InStrB(1, LCase$(strLine), "usercontrol=") Then
         lngI = InStr(1, strLine, "=") + 1
         strFileName = Mid$(strLine, lngI, Len(strLine) - lngI + 1)
         Call AddFileToList(strProjPath, strFileName)
      End If

      If InStrB(1, LCase$(strLine), "userdocument=") Then
         lngI = InStr(1, strLine, "=") + 1
         strFileName = Mid$(strLine, lngI, Len(strLine) - lngI + 1)
         Call AddFileToList(strProjPath, strFileName)
      End If

      If InStrB(1, LCase$(strLine), "designer=") Then
         lngI = InStr(1, strLine, "=") + 1
         strFileName = Mid$(strLine, lngI, Len(strLine) - lngI + 1)
         Call AddFileToList(strProjPath, strFileName)
      End If

   Loop

Exit_Proc:
   '// close project file port
   On Error Resume Next
   Close #lngFN
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "AddProject"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub CheckErrHandling()

   On Error GoTo Err_Proc

   '// Enable / disable objects attached to err handling frame
   chkErrObj.Enabled = optUseErrFunc.Value
   chkModuleName.Enabled = optUseErrFunc.Value
   chkProcName.Enabled = optUseErrFunc.Value

   txtErrHndl.Enabled = optUseFreeText.Value

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "CheckErrHandling"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub CheckForStartEnd(ByVal vstrLine As String, _
                             ByVal vstrModuleName As String, _
                             ByRef rstrProcName As String, _
                             ByRef rblnStartSub As Boolean, _
                             ByRef rblnEndSub As String, _
                             ByRef rblnHasOnError As Boolean, _
                             ByRef rblnHasOnErrorSub As Boolean, _
                             ByRef rblnHasExitLabel As Boolean, _
                             ByRef rblnHasErrLabel As Boolean, _
                             ByRef rlngLineCount As Long, _
                             ByRef rblnCreateTempFiles As Boolean)

   '// vstrLine         current line code
   '// rblnStartSub     start of SUB flag
   '// rblnEndSub       end of SUB flag
   '// rblnHasOnError   on error exist flag
   '// rblnCreateTempFiles

  Dim chkObjChk As VB.CheckBox

   If InStrB(vstrLine, "Sub") Then
      Set chkObjChk = chkApplyOnProc
   ElseIf InStrB(vstrLine, "Function") Then
      Set chkObjChk = chkApplyOnFunc
   ElseIf InStrB(vstrLine, "Property") Then
      Set chkObjChk = chkApplyOnProps
   End If

   vstrLine = Trim$(vstrLine)

   If Not rblnStartSub Then '// Start of SUB?

      If Left$(vstrLine, 12) = "Private Sub " Or _
            Left$(vstrLine, 17) = "Private Function " Or _
            Left$(vstrLine, 11) = "Public Sub " Or _
            Left$(vstrLine, 16) = "Public Function " Or _
            Left$(vstrLine, 14) = "Public Static " Or _
            Left$(vstrLine, 15) = "Private Static " Or _
            Left$(vstrLine, 11) = "Friend Sub " Or _
            Left$(vstrLine, 11) = "Static Sub " Or _
            Left$(vstrLine, 4) = "Sub " Or _
            Left$(vstrLine, 9) = "Function " Or _
            Left$(vstrLine, 21) Like "Private Property [LGS]et " Or _
            Left$(vstrLine, 20) Like "Public Property [LGS]et " Or _
            Left$(vstrLine, 20) Like "Static Property [LGS]et " Or _
            Left$(vstrLine, 20) Like "Friend Property [LGS]et " Or _
            Left$(vstrLine, 13) Like "Property [LGS]et " Then

         rstrProcName = GetProcName(vstrLine, InStr(vstrLine, "("))

         If rblnCreateTempFiles Then
            rblnStartSub = ((chkObjChk.Value = vbChecked) And (IsProcedureSelected(rstrProcName, vstrModuleName, rblnCreateTempFiles)))
         Else
            rblnStartSub = (IsProcedureSelected(rstrProcName, vstrModuleName, rblnCreateTempFiles))
         End If

         If rblnStartSub Then
            rblnHasOnError = False
            rblnHasExitLabel = False
            rblnHasErrLabel = False
            rlngLineCount = -1
         End If

         '// dont't try to add error handling to the Error handling function

         If UCase$(rstrProcName) = UCase$(txtFuncName.Text) Then
            rblnHasOnError = True
         End If

      End If

      '//  Check if this module already has the Error handling function

      If UCase$(rstrProcName) = UCase$(txtFuncName.Text) Then
         rblnHasOnErrorSub = True
      End If

   Else
      '// End of Procedure?

      If Left$(vstrLine, 7) = "End Sub" Then
         rblnEndSub = "Exit Sub"
      ElseIf Left$(vstrLine, 12) = "End Function" Then
         rblnEndSub = "Exit Function"
      ElseIf Left$(vstrLine, 12) = "End Property" Then
         rblnEndSub = "Exit Property"
      Else
         rblnEndSub = vbNullString
      End If

   End If

End Sub

Private Function CheckValidation() As Boolean

   '// Purpose: check that all the nesesary fields has data

  Dim ctrObj         As Control
  Dim blnMissingInfo As Boolean

   On Error GoTo Err_Proc

   For Each ctrObj In frmMain

      If TypeOf ctrObj Is TextBox Then
         If Trim$(ctrObj.Text) = vbNullString Then
            blnMissingInfo = True
            Exit For
         End If

      End If
   Next ctrObj

Exit_Proc:
   CheckValidation = Not blnMissingInfo

   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "CheckValidation"
   Err.Clear
   Resume Exit_Proc

End Function

Private Sub CloseOpenedProject()

   On Error Resume Next

   modTreeView.SetDefaultTree

   Erase mstrChangeSummary
   Erase mblnAlwaysIgnored

   cmdTransfer.Enabled = False
   cmdCommit.Enabled = False
   cmdView.Enabled = False
   cmdShowInterface.Enabled = False
   cmdBrows.Enabled = True
   txtControlPrefixes.Enabled = True
   chkIgnoreControlsPrefix.Enabled = True
   chkMinCodeLines.Enabled = True
   txtMinCodeLines.Enabled = True
   lblProcedureCount.Caption = vbNullString

   If mcFile.FolderExists(App.Path & "\DestTmp") Then

      With frmStatus
         .ProgressBar1.Visible = False
         .lblStatusText.Caption = "Deleting Temporary Files"
         .lblStatusText.Refresh
         .Show , Me
      End With

      DoEvents
      mcFile.DeleteDir App.Path & "\DestTmp"
      Unload frmStatus
   End If

   cmdBrows.SetFocus

End Sub

Private Sub cmdAbout_Click()

   frmAbout.Show , Me

End Sub

Private Sub cmdBrows_Click()

  Dim strFileName As String

   On Error GoTo Err_Proc

   With mcFile

      If .FolderExists(App.Path & "\DestTmp") Then
         .DeleteDir App.Path & "\DestTmp"
         .CreateDir App.Path & "\DestTmp"
      End If

      '// open dialog box
      .VBGetOpenFileName strFileName, , , , , , "VB Project (*.vbp)|*.vbp|All files (*.*)|*.*", , , "Open a VB Project", "vbp", Me.hWnd
   End With

   If Len(strFileName) = 0 Then GoTo Exit_Proc

   Call OpenProjectFile(strFileName)

Exit_Proc:
   Me.Show
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdBrows_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub cmdClear_Click()

   '// clear all the files from the list

   On Error Resume Next

   If trvModules.Nodes.Count = 5 Then Exit Sub

   If frmMsgBox.SMessageModal("Are you sure you want to clear the list ?", vbOKCancel Or vbQuestion) = vbOK Then

      Call CloseOpenedProject

   End If

End Sub

Private Sub cmdCollapse_Click()

   modTreeView.CollapseAllNodes (2)

End Sub

Private Sub cmdCommit_Click()

  Dim lngI        As Long
  Dim lngFN       As Long
  Dim strLastName As String

   '// Purpose:Make the temporary files (generate error handling code)

   On Error GoTo Err_Proc

   With frmStatus
      .lblStatusTitle.Caption = "Adding Error Handling to Temp File:"
      .lblStatusTitle.Refresh
      .Show , Me
   End With

   Me.Enabled = False
   DoEvents

   ReDim mstrChangeSummary(1, 0) As String

   '// parse and make new files using predefine rules like
   '// which function needs err handling
   Call ProcessFiles(True)

   '// Create Change Summary File

   If UBound(mstrChangeSummary, 2) > 0 Then
      lngFN = FreeFile
      Open App.Path & "\DestTmp\AddErrorCheckSummary.txt" For Output As #lngFN
      Print #lngFN, mcFile.RetOnlyPath(mstrProjectName) & "AddErrorCheckSummary.txt"
      Print #lngFN, "Change Summary - created: " & Now

      For lngI = 0 To UBound(mstrChangeSummary, 2) - 1

         If strLastName <> mstrChangeSummary(0, lngI) Then
            Print #lngFN, vbNullString
            Print #lngFN, mstrChangeSummary(0, lngI)
         End If

         Print #lngFN, "... " & mstrChangeSummary(1, lngI)
         strLastName = mstrChangeSummary(0, lngI)
      Next lngI

      Close #lngFN
   End If

   cmdView.Enabled = True
   cmdShowInterface.Enabled = True
   lblProcedureCount.Caption = "Total Procedures Changed: " & CStr(UBound(mstrChangeSummary, 2))

   Me.Enabled = True
   Unload frmStatus

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdCommit_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub cmdExit_Click()

   Unload Me

End Sub

Private Sub cmdShowInterface_Click()

   On Error GoTo Err_Proc

   If Dir$(App.Path & "\DestTmp\AddErrorCheckSummary.txt") > vbNullString Then

      With frmView
         .mblnSingleScreen = True
         .LoadTextView
         .Icon = Me.Icon
         .Show , Me
      End With

   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdShowInterface_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub cmdTransfer_Click()

   '// Purpose: Replace the original files with the generated files
   '//          the generated files has err handling code in every function

  Dim lngI As Long

   On Error GoTo Err_Proc

   lngI = frmMsgBox.SMessageModal("Are you sure you want to replace the original files with the " & _
         "error handled files ?", vbOKCancel + vbQuestion)

   If lngI = vbOK Then
      Me.Enabled = False
      frmStatus.Show , Me
      DoEvents

      If optUseErrFunc.Value = True Then
         If CreateErrorModule Then
            Call ReplaceFiles '// the final step
            frmMsgBox.SMessageModal "The transfer completed successfully", vbInformation + vbOkButton
         End If

      Else
         Call ReplaceFiles '// the final step
         frmMsgBox.SMessageModal "The transfer completed successfully", vbInformation + vbOkButton
      End If

   End If

Exit_Proc:
   Unload frmStatus
   Me.Enabled = True
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdTransfer_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub cmdView_Click()

   '// view source code vs generated code

   On Error GoTo Err_Proc

   If Not trvModules.SelectedItem Is Nothing Then

      With trvModules.SelectedItem

         If LenB(.Tag) > 0 Then
            frmView.Icon = Me.Icon
            frmView.Show , Me
            frmView.ShowEX .Tag
         End If

      End With
   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdView_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Function CreateErrorModule() As Boolean

  Dim lngRFN             As Long
  Dim lngFN              As Long
  Dim strAppPath         As String
  Dim strAppName         As String
  Dim strTextString      As String
  Dim blnErrModuleExists As Boolean

   On Error GoTo Err_Proc

   txtFuncName.Text = Trim$(txtFuncName.Text)
   txtModName.Text = Trim$(txtModName.Text)
   If txtFuncName.Text = vbNullString Or txtModName.Text = vbNullString Then Exit Function

   strAppPath = mcFile.RetOnlyPath(mstrProjectName)
   strAppName = mcFile.RetOnlyFilename(mstrProjectName)

   If Len(mstrErrorModulePath) Then
      If LenB(Dir$(mstrErrorModulePath)) Then
         If frmMsgBox.SMessageModal(txtModName.Text & ".bas   already exists in " & vbNewLine & _
               mcFile.RetOnlyPath(mstrErrorModulePath) & vbNewLine & vbNewLine & _
               "Do you wish to replace it?.", vbQuestion + vbYesNo, , , , , Me.Width) = vbYes Then

            mcFile.DeleteFile mstrErrorModulePath
         Else
            blnErrModuleExists = True
         End If

      End If
   End If

   If Not blnErrModuleExists Then
      '// Create module
      lngFN = FreeFile

      If Len(mstrErrorModulePath) Then
         Open mstrErrorModulePath For Output As #lngFN
      Else
         Open strAppPath & "\" & txtModName.Text & ".bas" For Output As #lngFN
      End If

      strTextString = "Attribute VB_Name = " & Chr$(34) & txtModName.Text & ".bas" & Chr$(34) & vbNewLine
      strTextString = strTextString & "Option Explicit" & vbNewLine
      strTextString = strTextString & WriteErrFunction(False)
      Print #lngFN, strTextString
      Close #lngFN
   End If

   '// Create backup directory
   mcFile.CreateDir strAppPath & "BeforeErrorCheck"

   With frmStatus
      .lblStatusText.Caption = "Creating " & txtModName.Text & ".bas"
      .lblStatusText.Refresh
   End With

   '// Add module to project
   mcFile.MoveFile mstrProjectName, strAppPath & "BeforeErrorCheck\" & strAppName, True, True

   blnErrModuleExists = False
   lngRFN = FreeFile
   Open strAppPath & "BeforeErrorCheck\" & strAppName For Input As #lngRFN
   lngFN = FreeFile
   Open mstrProjectName For Output As #lngFN
   strTextString = vbNullString

   Do Until EOF(lngRFN)
      Line Input #lngRFN, strTextString

      If InStrB(LCase$(strTextString), LCase$(txtModName.Text) & ".bas") Then
         blnErrModuleExists = True
      End If

      If Left$(strTextString, 8) = "Startup=" And Not blnErrModuleExists Then
         Print #lngFN, "Module=" & txtModName.Text & "; " & txtModName.Text & ".bas"
      End If

      Print #lngFN, strTextString
   Loop

   Close #lngRFN
   Close #lngFN

   CreateErrorModule = True

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "CreateErrorModule"
   Err.Clear
   Resume Exit_Proc

End Function

Private Sub Form_Load()

  Dim strCommandLine As String
  Dim cScreen        As clsScreenSize

   On Error GoTo Err_Proc

   Call ManifestWrite

   Set cScreen = New clsScreenSize
   cScreen.CenterForm Me
   Set cScreen = Nothing

   Me.Caption = "Add Error Handling - v" & App.Major & "." & App.Minor & "." & App.Revision & "   " & App.LegalCopyright

   Set mcFile = New clsFileUtilities
   Set mcTextBx = New clsTextBox
   mcTextBx.pAllowColorChange = True

   With trvModules
      .ImageList = imglst1
      .Checkboxes = True
      .LabelEdit = tvwManual
      .Indentation = 500
      .Style = tvwTreelinesPlusMinusPictureText
      .HideSelection = False
   End With

   '// set tree view in the global module
   Call InitializeTreeView(trvModules)
   Call SetDefaultTree

   '// Load Project From Command Line
   strCommandLine = Trim$(Command)

   If LenB(strCommandLine) > 2 Then
      If Left(strCommandLine, 1) = Chr$(34) Then
         strCommandLine = Mid$(strCommandLine, 2, Len(strCommandLine) - 2)
      End If

      If LenB(Dir$(strCommandLine)) Then
         If mcFile.GetExtensionName(strCommandLine) = "vbp" Then
            Call OpenProjectFile(strCommandLine)
         End If

      End If
   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "Form_Load"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub Form_Resize()

   On Error Resume Next

   fraOptions.Left = (Me.ScaleWidth \ 2) + 330
   trvModules.Left = fraOptions.Left - trvModules.Width - 60
   trvModules.Height = Me.ScaleHeight - trvModules.Top - picToolBar.Height - lblProcedureCount.Height
   lblProcedureCount.Move trvModules.Left, trvModules.Top + trvModules.Height, trvModules.Width
   lblMiscInfo(2).Left = trvModules.Left
   lblMiscInfo(2).Width = trvModules.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error Resume Next

   Me.Hide

   If mcFile.FolderExists(App.Path & "\DestTmp") Then

      With frmStatus
         .ProgressBar1.Visible = False
         .lblStatusTitle.Caption = "Deleting Temporary Files.."
         .lblStatusTitle.Refresh
         .Show , Me
      End With

      mcFile.DeleteDir App.Path & "\DestTmp"

      Unload frmStatus
   End If

   '// Clean-up
   Set mcFile = Nothing
   Set mcTextBx = Nothing

   Call EndApp(Me)
   Set frmMain = Nothing
   '// End Application

End Sub

Private Function GetErrFunctionConst(ByVal vstrModuleName As String, _
                                     ByVal vstrProcName As String) As String

   '// Purpose: gets the code for referencing to global error handling function

  Dim strTemp As String

   On Error GoTo Err_Proc

   '// insert tab
   strTemp = Space$(CLng(txtTabLength.Text))

   If optShowLog(0).Value Then '// Show and Log
      strTemp = strTemp & txtFuncName.Text & " True,"
   Else '// Log Only
      strTemp = strTemp & txtFuncName.Text & " False,"
   End If

   '// insert function params

   If chkErrObj.Value = vbChecked Then
      strTemp = strTemp & " Err.Number, Err.Description, "
   Else
      strTemp = strTemp & ", , "
   End If

   If chkModuleName.Value = vbChecked Then
      strTemp = strTemp & Chr$(34) & Trim$(vstrModuleName) & Chr$(34) & ", "
   Else
      strTemp = strTemp & ", "
   End If

   If chkProcName.Value = vbChecked Then
      strTemp = strTemp & Chr$(34) & Trim$(vstrProcName) & Chr$(34)
   End If

   If Len(strTemp) > 0 Then
      strTemp = strTemp & vbNewLine & Space$(CLng(txtTabLength.Text)) & "Err.Clear"
   End If

   GetErrFunctionConst = strTemp

Exit_Proc:

   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "GetErrFunctionConst"
   Err.Clear
   Resume Exit_Proc

End Function

Private Function GetModuleName(ByVal vstrLine As String) As String

   '// Purpose: parse the module name from the initializing line

   On Error GoTo Err_Proc

   If InStrB(vstrLine, "Attribute VB_Name") Then
      GetModuleName = Trim$(Mid$(vstrLine, 22))
   End If

   If Right$(GetModuleName, 1) = Chr$(34) Then
      GetModuleName = Left$(GetModuleName, Len(GetModuleName) - 1)
   End If

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "MGlobal", "GetModuleName"
   Err.Clear
   Resume Exit_Proc

End Function

Private Function GetProcName(ByVal vstrLine As String, _
                             ByVal vlngStartPoint As Long) As String

  Dim lngStartBr        As Long
  Dim strPropertyPrefix As String

   '// Purpose: parse the procedure name from the initializing line

   On Error GoTo Err_Proc

   '// If a property then add the prefix (Let or Get)

   If Left$(vstrLine, 21) Like "Private Property [LGS]et " Then
      strPropertyPrefix = Mid$(vstrLine, 18, 4)

   ElseIf Left$(vstrLine, 20) Like "Public Property [LGS]et " Or _
         Left$(vstrLine, 20) Like "Static Property [LGS]et " Or _
         Left$(vstrLine, 20) Like "Friend Property [LGS]et " Then

      strPropertyPrefix = Mid$(vstrLine, 17, 4)

   ElseIf Left$(vstrLine, 13) Like "Property [LGS]et " Then
      strPropertyPrefix = Mid$(vstrLine, 10, 4)
   Else
      strPropertyPrefix = vbNullString
   End If

   vlngStartPoint = vlngStartPoint - 1

   Do Until Len(GetProcName) > 0
      lngStartBr = InStrRev(vstrLine, " ", vlngStartPoint, vbTextCompare)

      If vlngStartPoint - lngStartBr > 0 Then
         GetProcName = strPropertyPrefix & Trim$(Mid$(vstrLine, lngStartBr + 1, (vlngStartPoint - lngStartBr)))
      End If

      If LenB(GetProcName) = 0 Then
         vlngStartPoint = lngStartBr - 1
      End If

   Loop

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "MGlobal", "GetProcName"
   Err.Clear
   Resume Exit_Proc

End Function

Private Function HasControlPrefix(ByVal vstrProcName As String) As Boolean

   '// Purpose: check if the function's name begins with a control prefix
   '//         that has been selected to ignore, if so then don't add error handling

  Dim lngI As Long

   On Error GoTo Err_Proc

   HasControlPrefix = False '// has no prefix by default

   If chkIgnoreControlsPrefix.Value = vbChecked Then
      vstrProcName = LCase$(Trim$(vstrProcName))

      For lngI = 0 To UBound(mstrAControlsPrefix)
         HasControlPrefix = CBool(InStr(vstrProcName, mstrAControlsPrefix(lngI)))
         If HasControlPrefix Then Exit For
      Next lngI

   End If

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "HasControlPrefix"
   Err.Clear
   Resume Exit_Proc

End Function

Private Function IsProcedureSelected(ByVal vstrProcName As String, _
                                     Optional ByVal vstrModuleName As String = vbNullString, _
                                     Optional ByVal vblnCreateTempFiles As Boolean = False) As Boolean

   '// Purpose: checks if the function was selected and not ignored
   '//          function is ignored when user unmark its checkbox

  Dim blnProcIndex As Long

   On Error GoTo Err_Proc

   IsProcedureSelected = True

   If Not vblnCreateTempFiles Then Exit Function

   IsProcedureSelected = modTreeView.IsNodeChecked(vstrProcName, vstrModuleName, blnProcIndex)

   If blnProcIndex > 0 Then
      If mblnAlwaysIgnored(blnProcIndex) Then
         trvModules.Nodes(blnProcIndex).Checked = False
         IsProcedureSelected = False
      End If

   End If

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "FunctionSelected"
   Err.Clear
   Resume Exit_Proc

End Function

Private Sub OpenProjectFile(ByVal strFileName As String)

   '// Purpose:Brows for a vb project or just one more free file
   '//         if vb project found than i load all its relevant code files

   On Error GoTo Err_Proc

   With mcFile

      If .FolderExists(App.Path & "\DestTmp") Then
         .DeleteDir App.Path & "\DestTmp"
         .CreateDir App.Path & "\DestTmp"
      End If

   End With

   If strFileName = vbNullString Then GoTo Err_Proc

   Screen.MousePointer = vbHourglass

   With frmStatus
      .lblStatusTitle.Caption = "Opening Project File:"
      .lblStatusTitle.Refresh
      .Show , Me
   End With

   Me.Enabled = False
   DoEvents

   '// Checks for file type and load TreeView control

   If mcFile.GetExtensionName(strFileName) = "vbp" Then
      mstrProjectName = strFileName
      Call AddProject(strFileName)  '// vb project - add all relevant files
   Else
      GoTo Exit_Proc
   End If

   '// Allow defining the selected files:
   mlngModuleCount = trvModules.Nodes.Count

   If trvModules.Nodes.Count > 0 Then

      With frmStatus
         .lblStatusTitle.Caption = "Processing Project File:"
         .lblStatusTitle.Refresh
      End With

      '// Parse the selected files
      'ReDim mblnAlwaysIgnored(trvModules.Nodes.Count) As Boolean
      Call ProcessFiles(False)  '// don't use the previous definitions
      cmdCommit.Enabled = True
   End If

   cmdBrows.Enabled = False
   txtControlPrefixes.Enabled = False
   chkIgnoreControlsPrefix.Enabled = False
   chkMinCodeLines.Enabled = False
   txtMinCodeLines.Enabled = False

Exit_Proc:
   Screen.MousePointer = vbDefault
   Unload frmStatus
   Me.Enabled = True
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "OpenProjectFile"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub optUseErrFunc_Click()

   Call CheckErrHandling

End Sub

Private Sub optUseFreeText_Click()

   Call CheckErrHandling

End Sub

Private Sub ProcessFiles(Optional ByVal vblnCreateTempFiles As Boolean = False)

   '// Purpose: parse all the selected files in the files list and generate
   '//          err handling code for all the selected functions

  Dim lngI As Long

   On Error GoTo Err_Proc

   If CheckValidation Then '// Check for missing data entries

      '// Prepare controls prefix array
      If chkIgnoreControlsPrefix.Value = vbChecked Then
         txtControlPrefixes.Text = TrimEX(txtControlPrefixes.Text, True) '// remove all spaces and make lower case
         mstrAControlsPrefix = Split(txtControlPrefixes.Text, ",")
      Else
         Erase mstrAControlsPrefix
      End If

      frmStatus.mcProg.Max = mlngModuleCount
      mstrErrorModulePath = vbNullString

      '// scan the files list
      With trvModules

         For lngI = 1 To .Nodes.Count

            If InStrB(LCase$(.Nodes(lngI).Text), LCase$(txtModName.Text) & ".bas") Then
               mstrErrorModulePath = .Nodes(lngI).Tag
            End If

            If .Nodes(lngI).Checked Then
               If LenB(.Nodes(lngI).Tag) > 0 Then
                  '// add err handling to the destination temp file
                  Call AddErrHandling(.Nodes(lngI).Tag, vblnCreateTempFiles)
               End If

            End If

            If lngI > mlngModuleCount Then Exit For
            frmStatus.mcProg.Value = lngI
         Next lngI

      End With

      cmdTransfer.Enabled = True
   Else
      frmMsgBox.SMessageModal "Cannot commit because one or more parameters are missing", vbOkButton + vbExclamation
   End If

Exit_Proc:

   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "ProcessFiles"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub ReplaceFiles()

   '//  This is the final step: replacing the old files with the new files

  Dim lngI          As Long
  Dim strFilePath   As String
  Dim strFileName   As String

   On Error GoTo Err_Proc

   '// Backup

   With frmStatus
      .mcProg.Max = trvModules.Nodes.Count
      .lblStatusTitle.Caption = "Backing-Up the Orignal Project File:"
      .lblStatusTitle.Refresh
      .Show , Me
   End With

   With trvModules
      strFilePath = mcFile.RetOnlyPath(mstrProjectName)
      mcFile.CopyFile App.Path & "\DestTmp\AddErrorCheckSummary.txt", strFilePath & "AddErrorCheckSummary.txt", True, True

      For lngI = 1 To .Nodes.Count

         If .Nodes.Item(lngI).Checked Then
            If .Nodes.Item(lngI).Tag <> vbNullString Then

               strFileName = mcFile.RetOnlyFilename(.Nodes.Item(lngI).Tag)
               strFilePath = mcFile.RetOnlyPath(.Nodes.Item(lngI).Tag) & "BeforeErrorCheck"

               If Not mcFile.FolderExists(strFilePath) Then
                  mcFile.CreateDir strFilePath
               End If

               frmStatus.lblStatusText.Caption = .Nodes.Item(lngI).Tag & vbNewLine & vbNewLine & _
                     " - to - " & vbNewLine & vbNewLine & strFilePath & strFileName
               frmStatus.lblStatusText.Refresh

               mcFile.MoveFile .Nodes.Item(lngI).Tag, strFilePath & "\" & strFileName, , True
            End If

         End If
         frmStatus.mcProg.Value = lngI
      Next lngI

   End With

   '// Replace

   With frmStatus
      .lblStatusTitle.Caption = "Replacing the Orignal Project File with Error Checking"
      .lblStatusTitle.Refresh
      .Show , Me
   End With

   With trvModules

      For lngI = 1 To .Nodes.Count

         If .Nodes.Item(lngI).Checked Then
            If .Nodes.Item(lngI).Tag <> vbNullString Then

               strFileName = mcFile.RetOnlyFilename(.Nodes.Item(lngI).Tag)
               frmStatus.lblStatusText.Caption = App.Path & "\DestTmp\" & strFileName & vbNewLine & vbNewLine & _
                  " - to -" & vbNewLine & vbNewLine & .Nodes.Item(lngI).Tag
               frmStatus.lblStatusText.Refresh

               mcFile.CopyFile App.Path & "\DestTmp\" & strFileName, .Nodes.Item(lngI).Tag, True, True

            End If
         End If
         frmStatus.mcProg.Value = lngI
      Next lngI

   End With

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "ReplaceFiles"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Function TrimEX(ByVal vstrValue As String, _
                        Optional ByVal blnLowerCase As Boolean = False) As String

  Dim strTemp As String
  Dim lngI    As Long

   '// remove all spaces
   On Error GoTo Err_Proc
   
   vstrValue = Trim$(vstrValue)
   strTemp = vbNullString

   For lngI = 1 To Len(vstrValue)
      strTemp = strTemp & IIf(Mid$(vstrValue, lngI, 1) <> " ", Mid$(vstrValue, lngI, 1), vbNullString)
   Next lngI

   If blnLowerCase Then
      TrimEX = LCase$(strTemp)
   Else
      TrimEX = strTemp
   End If

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "TrimEX"
   Err.Clear
   Resume Exit_Proc

End Function

Private Sub trvModules_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

   Call DisableCheck

End Sub

Private Sub trvModules_NodeCheck(ByVal vNode As MSComctlLib.Node)

   Call NodeCheckedEvent(vNode)

End Sub

Private Sub trvModules_OLEDragDrop(ByRef Data As MSComctlLib.DataObject, _
                                   ByRef Effect As Long, _
                                   ByRef Button As Integer, _
                                   ByRef Shift As Integer, _
                                   ByRef X As Single, _
                                   ByRef y As Single)

  Dim lngI        As Long
  Dim strFileName As String

   On Error GoTo Err_Proc

   Call CloseOpenedProject

   For lngI = 1 To Data.Files.Count
      '// File or directory?

      If (GetAttr(Data.Files(lngI)) And vbDirectory) = vbDirectory Then
         '// do nothing
      Else
         strFileName = Data.Files(lngI)
         '// Checks for file type and load TreeView control

         If mcFile.GetExtensionName(strFileName) = "vbp" Then
            mstrProjectName = strFileName
            Call AddProject(strFileName)  '// vb project - add all relevant files

            '// Allow defining the selected files:
            mlngModuleCount = trvModules.Nodes.Count

            If trvModules.Nodes.Count > 0 Then

               With frmStatus
                  .lblStatusTitle.Caption = "Processing Project File:"
                  .lblStatusTitle.Refresh
               End With

               '// Parse the selected files
               Call ProcessFiles(False)  '// don't use the previous definitions
               cmdCommit.Enabled = True
            End If

            cmdBrows.Enabled = False
            txtControlPrefixes.Enabled = False
            chkIgnoreControlsPrefix.Enabled = False
            chkMinCodeLines.Enabled = False
            txtMinCodeLines.Enabled = False
            Exit For
         End If

      End If
   Next lngI

Exit_Proc:

   Screen.MousePointer = vbDefault
   Unload frmStatus
   Me.Enabled = True
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "trvModules_OLEDragDrop"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub txtControlPrefixes_Change()

   mcTextBx.CaseLower txtControlPrefixes

End Sub

Private Sub txtControlPrefixes_GotFocus()

   mcTextBx.GotFocus txtControlPrefixes

End Sub

Private Sub txtControlPrefixes_KeyPress(KeyAscii As Integer)

   mcTextBx.KeyPress KeyAscii

End Sub

Private Sub txtControlPrefixes_KeyUp(KeyCode As Integer, Shift As Integer)

   mcTextBx.KeyUp txtControlPrefixes, KeyCode

End Sub

Private Sub txtControlPrefixes_LostFocus()

   mcTextBx.LostFocus txtControlPrefixes

End Sub

Private Sub txtErrLbl_Change()

   mcTextBx.Alpha_NoSymbols txtErrLbl

End Sub

Private Sub txtErrLbl_GotFocus()

   mcTextBx.GotFocus txtErrLbl

End Sub

Private Sub txtErrLbl_KeyPress(KeyAscii As Integer)

   mcTextBx.KeyPress KeyAscii

End Sub

Private Sub txtErrLbl_KeyUp(KeyCode As Integer, Shift As Integer)

   mcTextBx.KeyUp txtErrLbl, KeyCode

End Sub

Private Sub txtErrLbl_LostFocus()

   mcTextBx.LostFocus txtErrLbl

End Sub

Private Sub txtExitLabel_Change()

   mcTextBx.Alpha_NoSymbols txtExitLabel

End Sub

Private Sub txtExitLabel_GotFocus()

   mcTextBx.GotFocus txtExitLabel

End Sub

Private Sub txtExitLabel_KeyPress(KeyAscii As Integer)

   mcTextBx.KeyPress KeyAscii

End Sub

Private Sub txtExitLabel_KeyUp(KeyCode As Integer, Shift As Integer)

   mcTextBx.KeyUp txtExitLabel, KeyCode

End Sub

Private Sub txtExitLabel_LostFocus()

   mcTextBx.LostFocus txtExitLabel

End Sub

Private Sub txtFuncName_Change()

   mcTextBx.AlphaNumeric_NoSymbols txtFuncName

End Sub

Private Sub txtFuncName_GotFocus()

   mcTextBx.GotFocus txtFuncName

End Sub

Private Sub txtFuncName_KeyPress(KeyAscii As Integer)

   mcTextBx.KeyPress KeyAscii

End Sub

Private Sub txtFuncName_KeyUp(KeyCode As Integer, Shift As Integer)

   mcTextBx.KeyUp txtFuncName, KeyCode

End Sub

Private Sub txtFuncName_LostFocus()

   mcTextBx.LostFocus txtFuncName

End Sub

Private Sub txtLowerGap_Change()

   mcTextBx.Integer_Pos txtLowerGap

End Sub

Private Sub txtLowerGap_GotFocus()

   mcTextBx.GotFocus txtLowerGap

End Sub

Private Sub txtLowerGap_KeyPress(KeyAscii As Integer)

   mcTextBx.KeyPress KeyAscii

End Sub

Private Sub txtLowerGap_KeyUp(KeyCode As Integer, Shift As Integer)

   mcTextBx.KeyUp txtLowerGap, KeyCode

End Sub

Private Sub txtLowerGap_LostFocus()

   mcTextBx.LostFocus txtLowerGap

End Sub

Private Sub txtMinCodeLines_Change()

   mcTextBx.Integer_Pos txtMinCodeLines

End Sub

Private Sub txtMinCodeLines_GotFocus()

   mcTextBx.GotFocus txtMinCodeLines

End Sub

Private Sub txtMinCodeLines_KeyPress(KeyAscii As Integer)

   mcTextBx.KeyPress KeyAscii

End Sub

Private Sub txtMinCodeLines_KeyUp(KeyCode As Integer, Shift As Integer)

   mcTextBx.KeyUp txtMinCodeLines, KeyCode

End Sub

Private Sub txtMinCodeLines_LostFocus()

   mcTextBx.LostFocus txtMinCodeLines

End Sub

Private Sub txtModName_Change()

   mcTextBx.Alpha_NoSymbols txtModName

End Sub

Private Sub txtModName_GotFocus()

   mcTextBx.GotFocus txtModName

End Sub

Private Sub txtModName_KeyPress(KeyAscii As Integer)

   mcTextBx.KeyPress KeyAscii

End Sub

Private Sub txtModName_KeyUp(KeyCode As Integer, Shift As Integer)

   mcTextBx.KeyUp txtModName, KeyCode

End Sub

Private Sub txtModName_LostFocus()

   mcTextBx.LostFocus txtModName

End Sub

Private Sub txtTabLength_Change()

   mcTextBx.Integer_Pos txtTabLength

End Sub

Private Sub txtTabLength_GotFocus()

   mcTextBx.GotFocus txtTabLength

End Sub

Private Sub txtTabLength_KeyPress(KeyAscii As Integer)

   mcTextBx.KeyPress KeyAscii

End Sub

Private Sub txtTabLength_KeyUp(KeyCode As Integer, Shift As Integer)

   mcTextBx.KeyUp txtTabLength, KeyCode

End Sub

Private Sub txtTabLength_LostFocus()

   mcTextBx.LostFocus txtTabLength

End Sub

Private Sub txtUpperGap_Change()

   mcTextBx.Integer_Pos txtUpperGap

End Sub

Private Sub txtUpperGap_GotFocus()

   mcTextBx.GotFocus txtUpperGap

End Sub

Private Sub txtUpperGap_KeyPress(KeyAscii As Integer)

   mcTextBx.KeyPress KeyAscii

End Sub

Private Sub txtUpperGap_KeyUp(KeyCode As Integer, Shift As Integer)

   mcTextBx.KeyUp txtUpperGap, KeyCode

End Sub

Private Sub txtUpperGap_LostFocus()

   mcTextBx.LostFocus txtUpperGap

End Sub

Private Function WriteErrFunction(Optional ByVal vblnMakePrivate As Boolean = True) As String

   '// Returns the error function format which should be inseted to the module

  Dim strTemp As String
  Dim strTab1 As String
  Dim strTab2 As String

   On Error GoTo Err_Proc

   strTab1 = String$(txtTabLength.Text, " ")
   strTab2 = String$(txtTabLength.Text * 2, " ")

   If vblnMakePrivate Then
      strTemp = "Private Sub "
   Else
      strTemp = vbNewLine & "Public Sub "
   End If

   strTemp = strTemp & txtFuncName.Text & "(Optional ByVal vblnDisplayError As Boolean = True, _" & vbNewLine
   strTemp = strTemp & "                       Optional ByVal vstrErrNumber As String = vbNullString, _" & vbNewLine
   strTemp = strTemp & "                       Optional ByVal vstrErrDescription As String = vbNullString, _" & vbNewLine
   strTemp = strTemp & "                       Optional ByVal vstrModuleName As String = vbNullString, _" & vbNewLine
   strTemp = strTemp & "                       Optional ByVal vstrProcName As String = vbNullString)" & vbNewLine

   strTemp = strTemp & vbNewLine
   strTemp = strTemp & "  Dim strTemp As String" & vbNewLine
   strTemp = strTemp & "  Dim lngFN   As Long" & vbNewLine

   strTemp = strTemp & vbNewLine
   strTemp = strTemp & strTab1 & "On Error Resume Next" & vbNewLine
   strTemp = strTemp & strTab1 & "'// Purpose: Error handling - On Error " & vbNewLine
   strTemp = strTemp & vbNewLine

   strTemp = strTemp & strTab1 & "'// Show Error Message" & vbNewLine
   strTemp = strTemp & strTab1 & "If vblnDisplayError Then" & vbNewLine
   strTemp = strTemp & strTab2 & "strTemp = " & Chr$(34) & "Error occured: " & Chr$(34) & vbNewLine
   strTemp = strTemp & strTab2 & "If Lenb(vstrErrNumber) Then strTemp = strTemp & vstrErrNumber & vbNewLine else strTemp = strTemp & vbNewLine" & vbNewLine
   strTemp = strTemp & strTab2 & "If Lenb(vstrErrDescription) Then strTemp = strTemp & " & Chr$(34) & "Description: " & Chr$(34) & " & vstrErrDescription & vbNewLine" & vbNewLine
   strTemp = strTemp & strTab2 & "If Lenb(vstrModuleName) Then strTemp = strTemp & " & Chr$(34) & "Module: " & Chr$(34) & " & vstrModuleName & vbNewLine" & vbNewLine
   strTemp = strTemp & strTab2 & "If Lenb(vstrProcName) Then strTemp = strTemp & " & Chr$(34) & "Function: " & Chr$(34) & " & vstrProcName" & vbNewLine
   strTemp = strTemp & strTab2 & "MsgBox strTemp, vbCritical, App.Title & " & Chr$(34) & " - ERROR" & Chr$(34) & vbNewLine
   strTemp = strTemp & strTab1 & "End If" & vbNewLine

   strTemp = strTemp & vbNewLine
   strTemp = strTemp & strTab1 & "'// Write error log" & vbNewLine
   strTemp = strTemp & strTab1 & "lngFN = FreeFile" & vbNewLine
   strTemp = strTemp & strTab1 & "Open App.Path & " & Chr$(34) & "\ErrorLog.txt" & Chr$(34) & " For Append As #lngFN" & vbNewLine
   strTemp = strTemp & strTab1 & "Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, _" & vbNewLine
   strTemp = strTemp & strTab2 & "App.Title & " & Chr$(34) & " v" & Chr$(34) & " & App.Major & " & Chr$(34) & "." & Chr$(34) & " & App.Minor & " & Chr$(34) & "." & Chr$(34) & " & App.Revision, _" & vbNewLine
   strTemp = strTemp & strTab2 & "Environ(" & Chr$(34) & "username" & Chr$(34) & "), Environ(" & Chr$(34) & "computername" & Chr$(34) & ")" & vbNewLine
   strTemp = strTemp & strTab1 & "Close #lngFN" & vbNewLine & vbNewLine
   strTemp = strTemp & "End Sub"

   WriteErrFunction = strTemp

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmMain", "WriteErrFunction"
   Err.Clear
   Resume Exit_Proc

End Function


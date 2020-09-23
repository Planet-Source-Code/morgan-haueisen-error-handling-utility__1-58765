VERSION 5.00
Begin VB.Form frmStatus 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3870
   ClientLeft      =   3375
   ClientTop       =   2550
   ClientWidth     =   7305
   ControlBox      =   0   'False
   ForeColor       =   &H80000015&
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7305
   Begin VB.PictureBox ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   7245
      TabIndex        =   2
      Top             =   3570
      Visible         =   0   'False
      Width           =   7305
   End
   Begin AddErrorChecking.Frame3D txtStatusWorking 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   873
      BorderType      =   0
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlignment=   4
      CaptionLocation =   0
      BackColor       =   -2147483628
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
      MouseIcon       =   "frmStatus.frx":000C
      Picture         =   "frmStatus.frx":0028
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Times New Roman"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   "Working... Please Wait"
      UseMnemonic     =   0   'False
   End
   Begin VB.Label lblStatusText 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   2565
      Left            =   210
      TabIndex        =   1
      Top             =   855
      Width           =   6900
   End
   Begin VB.Label lblStatusTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   7305
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//************************************/
'// Author: Morgan Haueisen
'//         morganh@hartcom.net
'// Copyright (c) 2003-2004
'//************************************/
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
'        This code was developed by Morgan Haueisen.  <morganh@hartcom.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'        Redistributions in binary form, as part of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.

Option Explicit

Public mcProg As clsProgressBar

Private Sub Form_Initialize()

   On Error GoTo Err_Proc

   Set mcProg = New clsProgressBar

   With mcProg
      .ForeColor2 = .ForeColor
      .PicBox = ProgressBar1
      .Style = pbStepped
      .ShowCounts = False
      .ShowStatus = False
   End With

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmStatus", "Form_Initialize"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub Form_Load()

  Dim cScreen As clsScreenSize

   Set cScreen = New clsScreenSize
   cScreen.CenterForm Me
   Set cScreen = Nothing

   Me.Show , frmMain
   DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error Resume Next
   Set mcProg = Nothing
   Set frmStatus = Nothing

End Sub


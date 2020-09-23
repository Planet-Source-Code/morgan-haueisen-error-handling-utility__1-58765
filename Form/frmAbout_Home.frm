VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2205
   ClientLeft      =   2895
   ClientTop       =   3015
   ClientWidth     =   7125
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAbout_Home.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout_Home.frx":000C
   ScaleHeight     =   2205
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblWebSite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WWW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   555
   End
   Begin VB.Label lblCompanyName 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCompanyName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   390
      UseMnemonic     =   0   'False
      Width           =   5130
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout_Home.frx":1E6E
      ForeColor       =   &H00C0C0FF&
      Height          =   1110
      Left            =   2040
      TabIndex        =   3
      Top             =   1380
      UseMnemonic     =   0   'False
      Width           =   4965
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   2040
      TabIndex        =   2
      Top             =   645
      UseMnemonic     =   0   'False
      Width           =   5100
   End
   Begin VB.Label lblProdDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   30
      UseMnemonic     =   0   'False
      Width           =   5085
   End
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail ME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   510
      TabIndex        =   0
      ToolTipText     =   "morganh@hartcom.net"
      Top             =   1875
      Width           =   855
   End
End
Attribute VB_Name = "frmAbout"
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

Private Declare Function SystemParametersInfo Lib "user32" _
      Alias "SystemParametersInfoA" ( _
      ByVal uAction As Long, _
      ByVal uParam As Long, _
      ByRef lpvParam As Any, _
      ByVal fuWinIni As Long) As Long
Private Const C_SPI_GETWORKAREA As Long = 48&

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cX As Long, _
      ByVal cY As Long, _
      ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" ( _
      ByVal hWnd As Long, _
      ByVal lpOperation As String, _
      ByVal lpFile As String, _
      ByVal lpParameters As String, _
      ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long

Public PreventClose As Boolean
Public AlwaysOnTop  As Boolean
Public SleepTime    As Long

Private Sub CenterForm()

  Dim udtRc As RECT
  Dim lngT  As Long
  Dim lngB  As Long
  Dim lngL  As Long
  Dim lngR  As Long
  Dim lngmT As Long
  Dim lngmL As Long

   On Error GoTo Err_Proc

   Call SystemParametersInfo(C_SPI_GETWORKAREA, 0&, udtRc, 0&)

   lngT = udtRc.Top * Screen.TwipsPerPixelY
   lngB = udtRc.Bottom * Screen.TwipsPerPixelY
   lngL = udtRc.Left * Screen.TwipsPerPixelX
   lngR = udtRc.Right * Screen.TwipsPerPixelX

   lngmT = Abs((lngB / 2.8) - (Me.Height / 2))
   lngmL = Abs((lngR / 2) - (Me.Width / 2))

   If lngmT < lngT Then lngmT = lngT
   If lngmT > lngB - Me.Height Then lngmT = lngB - Me.Height
   If lngmL < lngL Then lngmL = lngL

   Me.Move lngmL, lngmT

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmAbout", "CenterForm"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub Form_Click()

   If Not PreventClose Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()

   Call CenterForm

   On Error Resume Next
   lblProdDesc.Caption = App.ProductName
   lblCompanyName.Caption = "MorganWare™" 'App.CompanyName

   lblVersion.Caption = "By: Morgan Haueisen and Adi barda" & vbCrLf & _
         "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
         App.LegalCopyright

   Me.Show
   DoEvents

   If AlwaysOnTop Then Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
   If SleepTime > 0 Then Sleep SleepTime

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   lblEMail.Font.Underline = False
   lblWebSite.Font.Underline = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error Resume Next
   Set frmAbout = Nothing

End Sub

Private Sub lblDisclaimer_Click()

   If Not PreventClose Then Unload Me

End Sub

Private Sub lblEMail_Click()

   ShellExecute Me.hWnd, "open", "mailto:" & lblEMail.ToolTipText & _
         "?subject=" & App.ProductName, vbNullString, "C:\", 5

End Sub

Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   lblEMail.Font.Underline = True

End Sub

Private Sub lblProdDesc_Click()

   If Not PreventClose Then Unload Me

End Sub

Private Sub lblVersion_Click()

   If Not PreventClose Then Unload Me

End Sub

Private Sub lblWebSite_Click()

   ShellExecute Me.hWnd, "open", _
         "http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=885253927&strAuthorName=Morgan%20Haueisen&txtMaxNumberOfEntriesPerPage=25", _
         vbNullString, "C:\", 5

End Sub

Private Sub lblWebSite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   lblWebSite.Font.Underline = True

End Sub


VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View File"
   ClientHeight    =   6930
   ClientLeft      =   1395
   ClientTop       =   1740
   ClientWidth     =   8625
   Icon            =   "frmViewS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   4470
      ScaleHeight     =   4785
      ScaleWidth      =   3075
      TabIndex        =   5
      Top             =   225
      Width           =   3105
      Begin VB.CommandButton cmdDFullScreen 
         Caption         =   "Full Screen"
         Height          =   360
         Left            =   15
         TabIndex        =   7
         Top             =   0
         Width           =   1230
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   360
         Left            =   1485
         TabIndex        =   6
         Top             =   0
         Width           =   1230
      End
      Begin RichTextLib.RichTextBox txtDest 
         Height          =   3945
         Left            =   0
         TabIndex        =   8
         Top             =   585
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   6959
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmViewS.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMiscLabel 
         AutoSize        =   -1  'True
         Caption         =   " Dest file: "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   9
         Top             =   375
         Width           =   705
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   0
      ScaleHeight     =   4875
      ScaleWidth      =   3345
      TabIndex        =   1
      Top             =   105
      Width           =   3375
      Begin VB.CommandButton cmdSFullScreen 
         Caption         =   "Full Screen"
         Height          =   360
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1230
      End
      Begin RichTextLib.RichTextBox txtSource 
         Height          =   4110
         Left            =   0
         TabIndex        =   3
         Top             =   585
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   7250
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmViewS.frx":00A2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMiscLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Source file: "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   375
         Width           =   885
      End
   End
   Begin VB.PictureBox spltVertical 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H80000010&
      FillStyle       =   0  'Solid
      ForeColor       =   &H8000000F&
      Height          =   4935
      Left            =   3870
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4905
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   210
      Width           =   105
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_SPLT_WDTH As Long = 50&  'width of the spltter bar
Private Const C_MIN_WINDOW As Long = 10& 'Minimum size For any frame created by splitter bars

Public mblnSingleScreen As Boolean

Private mcObjEditor  As clsEditor '// main editor object
Private mcObjEditor2 As clsEditor '// main editor object

Private mblnDFullScreen As Boolean
Private mblnSFullScreen As Boolean

Private Sub cmdDFullScreen_Click()

   On Error GoTo Err_Proc

   If Not mblnDFullScreen Then
      spltVertical.Left = C_MIN_WINDOW
      Call ResizeLeftRightWindows
      cmdDFullScreen.Caption = "Normal View"
      txtDest.ZOrder
      lblMiscLabel(0).Visible = False
      mblnDFullScreen = True
   Else
      spltVertical.Left = Me.ScaleWidth \ 2
      Call ResizeLeftRightWindows
      cmdDFullScreen.Caption = "Full Screen"
      lblMiscLabel(0).Visible = True
      mblnDFullScreen = False
   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "cmdDFullScreen_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub cmdPrint_Click()

   On Error GoTo Err_Proc

   If frmMsgBox.SMessageModal("Ok to print the Destination file?", vbQuestion + vbYesNo) = vbYes Then
      Printer.ColorMode = vbPRCMColor
      Printer.Print vbNullString;
      Printer.Print Me.txtDest.Text
      Printer.EndDoc
   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "cmdPrint_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub cmdSFullScreen_Click()

   On Error GoTo Err_Proc

   If Not mblnSFullScreen Then
      spltVertical.Left = Me.ScaleWidth
      Call ResizeLeftRightWindows
      cmdSFullScreen.Caption = "Normal View"
      txtSource.ZOrder
      lblMiscLabel(1).Visible = False
      mblnSFullScreen = True
   Else
      spltVertical.Left = Me.ScaleWidth \ 2
      Call ResizeLeftRightWindows
      cmdSFullScreen.Caption = "Full Screen"
      lblMiscLabel(1).Visible = True
      mblnSFullScreen = False
   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "cmdSFullScreen_Click"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub Form_Load()

  Dim cScreen As clsScreenSize

   On Error GoTo Err_Proc

   Set cScreen = New clsScreenSize
   cScreen.FitScreen Me
   Set cScreen = Nothing

   Screen.MousePointer = vbHourglass
   DoEvents

   If Not mblnSingleScreen Then
      Set mcObjEditor = New clsEditor
      Set mcObjEditor2 = New clsEditor

      '// set editor objects
      mcObjEditor.SetEditorObjects Me.txtSource, Nothing, Nothing, Nothing, Nothing, Me.txtDest
      InitWords mcObjEditor

      mcObjEditor2.SetEditorObjects Me.txtDest, Nothing, Nothing, Nothing, Nothing, Me.txtSource
      InitWords mcObjEditor2
   End If

   '// Set defaults for vertical splitter bar
   spltVertical.BorderStyle = 0
   picLeft.BorderStyle = 0
   picRight.BorderStyle = 0
   spltVertical.ZOrder
   spltVertical.Left = Me.ScaleWidth \ 2
   Call ResizeLeftRightWindows

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "Form_Load"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error Resume Next
   Set mcObjEditor = Nothing
   Set mcObjEditor2 = Nothing
   Set frmView = Nothing
   frmMain.trvModules.SetFocus

End Sub

Private Sub InitWords(ByRef fcObj As clsEditor)

   '// hard code init the basic vb script words -
   '// you can init any words you want with any colors you like

   On Error GoTo Err_Proc

   With fcObj
      .AddEditorWord "On Local Error", vbMagenta
      .AddEditorWord "On Error", vbMagenta
      .AddEditorWord "GoTo " & frmMain.txtErrLbl, vbRed
      .AddEditorWord frmMain.txtErrLbl & ":", vbRed
      .AddEditorWord frmMain.txtExitLabel & ":", vbRed
      .AddEditorWord frmMain.txtFuncName, vbRed
      .AddEditorWord "Public", vbBlue
      .AddEditorWord "Private", vbBlue
      .AddEditorWord "Function", vbBlue
      .AddEditorWord "Sub", vbBlue
      .AddEditorWord "End", vbBlue
   End With

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "InitWords"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub LoadTextView()

  Dim lngFN   As Long
  Dim strTemp As String
  Dim strView As String

   On Error GoTo Err_Proc

   spltVertical.Left = C_MIN_WINDOW
   Call ResizeLeftRightWindows
   spltVertical.Enabled = False
   txtSource.Visible = False
   cmdSFullScreen.Visible = False
   cmdDFullScreen.Visible = False
   lblMiscLabel(0).Visible = False
   lblMiscLabel(1).Left = 15
   lblMiscLabel(1).Caption = "This file will be copied to the projects folder on transfer."

   '//  Open Source file:
   lngFN = FreeFile
   Open App.Path & "\DestTmp\AddErrorCheckSummary.txt" For Input As #lngFN
   strView = vbNullString

   Do Until EOF(lngFN)
      Line Input #lngFN, strTemp
      strView = strView & strTemp & vbNewLine
   Loop

   txtDest.Text = strView
   Close #lngFN

Exit_Proc:
   Screen.MousePointer = vbDefault
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "LoadTextView"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub ResizeLeftRightWindows()

   On Error Resume Next

   picLeft.Move 0, 0, spltVertical.Left - 1, Me.ScaleHeight
   spltVertical.Move spltVertical.Left, 0, C_SPLT_WDTH, Me.ScaleHeight
   picRight.Move spltVertical.Left + C_SPLT_WDTH + 1, 0, _
         Me.ScaleWidth - spltVertical.Left + C_SPLT_WDTH + 1, Me.ScaleHeight

   If Not mblnSingleScreen Then
      txtSource.Width = picLeft.Width
      txtSource.Height = picLeft.Height - txtSource.Top

      txtDest.Width = picRight.Width
      txtDest.Height = picRight.Height - txtDest.Top

      lblMiscLabel(0).Visible = True
      lblMiscLabel(1).Visible = True

   Else
      txtDest.Move 0, txtDest.Top, Me.ScaleWidth, Me.ScaleHeight - txtDest.Top - 50
   End If

End Sub

Public Sub ShowEX(ByVal vstrFilePath As String, Optional ByVal vblnShowInterface As Boolean = False)

   On Error GoTo Err_Proc

   Me.txtSource.Text = vbNullString
   Me.txtDest.Text = vbNullString

   Call UpdateTextView(vstrFilePath, vblnShowInterface)

   If vblnShowInterface Then

      With Me.txtDest
         .Width = .Left + .Width
         .Left = Me.txtSource.Left
         .ZOrder 0
         lblMiscLabel(0).Visible = False
         lblMiscLabel(1).Left = lblMiscLabel(0).Left
      End With

   End If

   mcObjEditor.PaintText False, True
   mcObjEditor2.PaintText False, True

Exit_Proc:
   Screen.MousePointer = vbDefault
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "ShowEX"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub spltVertical_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then
      spltVertical.Move (spltVertical.Left - (C_SPLT_WDTH \ 2)) + X, 0, C_SPLT_WDTH, Me.ScaleHeight
      spltVertical.BackColor = vbButtonShadow 'change the splitter colour
   End If

End Sub

Private Sub spltVertical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If spltVertical.BackColor = vbButtonShadow Then
      spltVertical.Move (spltVertical.Left - (C_SPLT_WDTH \ 2)) + X, 0, C_SPLT_WDTH, Me.ScaleHeight
   End If

End Sub

Private Sub spltVertical_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim lngAbsLeft As Long
  Dim lngAbsRight As Long

   On Error GoTo Err_Proc

   If spltVertical.BackColor = vbButtonShadow Then
      spltVertical.BackColor = vbHighlight 'vbButtonFace 'restore splitter colour
      spltVertical.Move (spltVertical.Left - (C_SPLT_WDTH \ 2)) + X, 0, C_SPLT_WDTH, Me.ScaleHeight

      'Set the absolute Boundaries
      lngAbsLeft = C_MIN_WINDOW
      lngAbsRight = Me.ScaleWidth - (C_SPLT_WDTH + C_MIN_WINDOW)

      Select Case spltVertical.Left
      Case Is < lngAbsLeft 'the pane is too thin
         spltVertical.Move lngAbsLeft, 0, C_SPLT_WDTH, Me.ScaleHeight

      Case Is > lngAbsRight 'the pane is too wide
         spltVertical.Move lngAbsRight, 0, C_SPLT_WDTH, Me.ScaleHeight
      End Select

      'reposition both frames, and the spltVertical bar
      Call ResizeLeftRightWindows
   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "spltVertical_MouseUp"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub UpdateTextView(ByVal vstrFilePath As String, Optional ByVal vblnShowInterface As Boolean = False)

  Dim lngFN    As Long
  Dim strTemp  As String
  Dim strView  As String
  Dim blnFlag  As Boolean
  Dim blnFlag2 As Boolean
  Dim cFile    As clsFileUtilities

   On Error GoTo Err_Proc

   '//  Open Source file:
   lngFN = FreeFile
   Open vstrFilePath For Input As #lngFN
   strView = vbNullString

   Do Until EOF(lngFN)
      Line Input #lngFN, strTemp
      blnFlag2 = InStr(Trim$(strTemp), "Attribute VB_") = 1

      If Not blnFlag Or blnFlag2 Then
         strTemp = vbNullString
         blnFlag = blnFlag2
      Else
         strView = strView & strTemp & vbNewLine
      End If

   Loop
   txtSource.Text = strView
   Close #lngFN

   Set cFile = New clsFileUtilities
   vstrFilePath = cFile.RetOnlyFilename(vstrFilePath)
   Set cFile = Nothing

   strView = vbNullString
   blnFlag = False

   If LenB(Dir$(App.Path & "\DestTmp\" & vstrFilePath)) Then
      '//  Open Dest file:
      lngFN = FreeFile
      Open App.Path & "\DestTmp\" & vstrFilePath For Input As #lngFN

      Do Until EOF(lngFN)
         Line Input #lngFN, strTemp
         blnFlag2 = InStr(Trim$(strTemp), "Attribute VB_") = 1

         If Not blnFlag Or blnFlag2 Then
            strTemp = vbNullString
            blnFlag = blnFlag2
         Else
            strView = strView & strTemp & vbNewLine
         End If

      Loop
      Me.txtDest.Text = strView
      Close #lngFN
   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "frmView", "UpdateTextView"
   Err.Clear
   Resume Exit_Proc

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//************************************/
'// Author: Morgan Haueisen
'// Copyright (c) 1997-2004
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

'// Windows API/Global Declarations for : Determine the Available Screen Real Estate

Private Type RECT
   Left      As Long
   Top       As Long
   Right     As Long
   Bottom    As Long
End Type

Private Const SPI_GETWORKAREA As Long = 48&
Private Declare Function SystemParametersInfo Lib "user32" _
      Alias "SystemParametersInfoA" ( _
      ByVal uAction As Long, _
      ByVal uParam As Long, _
      ByRef lpvParam As Any, _
      ByVal fuWinIni As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal y As Long, _
      ByVal cX As Long, _
      ByVal cY As Long, _
      ByVal wFlags As Long) As Long

'// Windows API/Global Declarations for IsScreenFontSmall
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" ( _
      ByVal hdc As Long, _
      ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" ( _
      ByVal hWnd As Long, _
      ByVal hdc As Long) As Long

'// Windows API/Global Declarations for Resizing
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

Private Const C_lngSSTAB_DIALOG_OFFSET            As Long = 75000
Private Const C_strSSTAB_TYPE_NAME                As String = "SSTAB"
Private Const C_strResizeRepositionCommand        As String = "@"
Private Const C_strResizeRepositionLeft           As String = "lngLeft"
Private Const C_strResizeRepositionTop            As String = "lngTop"
Private Const C_strResizeRepositionWidth          As String = "W"
Private Const C_strResizeRepositionHeight         As String = "H"

Private Type ControlPosition
   ControlInstance          As Control      'Reference to the control instance.
   OriginalLeft             As Long         'Original Left position of the Control.
   OriginalTop              As Long         'Original Top position of the Control.
   OriginalWidth            As Long         'Original Width of the Control.
   OriginalHeight           As Long         'Original Height of the Control.
End Type

Private mfrm_SourceForm      As Form       'The form to be resized.
Private mlng_FormWidth       As Long       'Original form width.
Private mlng_FormHeight      As Long       'Original form height.
Private mudt_Controls()      As ControlPosition 'Array for storing control information.
Private mbln_IsFirstResize   As Boolean    'Flag indicating first resize.

Public Enum enuOnTopPar
   ON_TOP_YES = False
   ON_TOP_NO = True
End Enum

Public Sub CenterForm(ByRef rfrmForm As Form, _
                      Optional ByVal vsngMoveUpDn As Single = 1)

  Dim RC        As RECT
  Dim lngTop    As Long
  Dim lngBottom As Long
  Dim lngLeft   As Long
  Dim lngRight  As Long
  Dim lngTopT   As Long
  Dim lngLeftT  As Long

   On Error Resume Next

   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, RC, 0&)

   lngTop = RC.Top * Screen.TwipsPerPixelY
   lngBottom = RC.Bottom * Screen.TwipsPerPixelY
   lngLeft = RC.Left * Screen.TwipsPerPixelX
   lngRight = RC.Right * Screen.TwipsPerPixelX

   lngTopT = Abs((lngBottom / 2 * vsngMoveUpDn) - (rfrmForm.Height / 2))
   lngLeftT = Abs((lngRight / 2) - (rfrmForm.Width / 2))

   If lngTopT < lngTop Then lngTopT = lngTop
   If lngTopT > lngBottom - rfrmForm.Height Then lngTopT = lngBottom - rfrmForm.Height
   If lngLeftT < lngLeft Then lngLeftT = lngLeft

   rfrmForm.Move lngLeftT, lngTopT

End Sub

Public Sub CenterObject(ByRef ParentObject As Object, _
                        ByRef ChildObject As Object)

   On Error Resume Next
   ChildObject.Top = (ParentObject.Height - ChildObject.Height) / 2
   ChildObject.Left = (ParentObject.Width - ChildObject.Width) / 2

End Sub

Private Sub Class_Initialize()

   Set mfrm_SourceForm = Nothing
   mbln_IsFirstResize = True

End Sub

Public Sub FitScreen(ByRef rfrmForm As Form, _
                     Optional ByVal vsngPercent As Single = 100)

  Dim RC As RECT

   On Error Resume Next

   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, RC, 0&)
   'left - " & rc.Left
   'top - " & rc.Top
   'rght - " & rc.Right
   'bttm - " & rc.Bottom

   If vsngPercent > 100 Then vsngPercent = 100
   vsngPercent = vsngPercent / 100

   rfrmForm.Move RC.Left * Screen.TwipsPerPixelX * vsngPercent, _
         RC.Top * Screen.TwipsPerPixelY * vsngPercent, _
         RC.Right * Screen.TwipsPerPixelX * vsngPercent, _
         RC.Bottom * Screen.TwipsPerPixelY * vsngPercent

   If vsngPercent < 1 Then Call CenterForm(rfrmForm)

End Sub

Public Sub HCenterObject(ByRef rObj As Object)

  Dim RC       As RECT
  Dim lngLeft  As Long

   On Error Resume Next
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, RC, 0&)
   lngLeft = RC.Right * Screen.TwipsPerPixelX
   lngLeft = lngLeft / 2 - rObj.Width / 2
   rObj.Left = lngLeft

End Sub

Public Property Get Height() As Variant

   Height = GetSystemMetrics(1&)

End Property

Public Sub hFitScreen(ByRef rfrmForm As Form, _
                      Optional vblnCenter As Boolean = True)

  Dim RC As RECT

   On Error Resume Next
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, RC, 0&)
   rfrmForm.Left = RC.Left * Screen.TwipsPerPixelX
   rfrmForm.Width = (RC.Right - RC.Left) * Screen.TwipsPerPixelX
   If vblnCenter Then CenterForm rfrmForm

End Sub

Public Sub OnTop(ByRef rfrmForm As Form, _
                 Optional vblnTurnOff As enuOnTopPar = ON_TOP_YES)

  Const HWND_TOPMOST    As Long = -1
  Const HWND_NOTTOPMOST As Long = -2

   On Error Resume Next

   If vblnTurnOff Then
      Call SetWindowPos(rfrmForm.hWnd, HWND_NOTTOPMOST, 0, 0, 0, 0, 3)
   Else
      Call SetWindowPos(rfrmForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
   End If

End Sub

Public Function scrnBottom() As Long

  Dim RC As RECT

   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, RC, 0&)
   scrnBottom = RC.Bottom * Screen.TwipsPerPixelY

End Function

Public Function scrnLeft() As Long

  Dim RC As RECT

   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, RC, 0&)
   scrnLeft = RC.Left * Screen.TwipsPerPixelX

End Function

Public Function scrnRight() As Long

  Dim RC As RECT

   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, RC, 0&)
   scrnRight = RC.Right * Screen.TwipsPerPixelX

End Function

Public Function scrnTop() As Long

  Dim RC As RECT

   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, RC, 0&)
   scrnTop = RC.Top * Screen.TwipsPerPixelY

End Function

Public Sub vFitScreen(ByRef rfrmForm As Form, _
                      Optional vblnCenter As Boolean = True, _
                      Optional vblnUseFullScreen As Boolean = False)

  Dim RC          As RECT
  Dim lngScTop    As Long
  Dim lngScBottom As Long

   On Error Resume Next

   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, RC, 0&)

   lngScTop = RC.Top * Screen.TwipsPerPixelY
   lngScBottom = RC.Bottom * Screen.TwipsPerPixelY

   If lngScTop = 0 And Not vblnUseFullScreen Then lngScTop = 250
   rfrmForm.Height = lngScBottom - lngScTop
   If vblnCenter Then CenterForm rfrmForm
   rfrmForm.Top = lngScTop

End Sub

Public Property Get Width() As Variant

   Width = GetSystemMetrics(0&)

End Property


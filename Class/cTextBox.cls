VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// DISCLAIMER (for public posting):
'// This software is provided on an "as is, where is" basis, with no warranty whatsoever.

'// Class: cTextBox
'// Purpose: To provide enhanced textbox control
'// Author: Morgan Haueisen (Copyright (c) 2003) morganh@hartcom.net
'// Date: Sept 22, 2003 (Last Updated: 11/5/03)

Private mstrUndoText                As String
Private mlngBackColor_Normal        As Long
Private mlngBackColor_GotFocus      As Long
Private mlngForeColor_Normal        As Long
Private mlngForeColor_GotFocus      As Long
Private mblnAllowColorChange        As Boolean
Private mblnRemoveExtraSpaces       As Boolean
Private mblnWorking                 As Boolean
Private mbytLastChar                As Byte

Public Sub AlphaNumeric_NoSymbols(ByRef rTextObj As TextBox)

  Dim lngI    As Long
  Dim lngN    As Long
  Dim strTemp As String
  Dim bytChar As Byte

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      For lngN = 1 To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If bytChar = 32 Or bytChar = 95 Then
            strTemp = strTemp & "_"
         Else

            If bytChar < 48 Or (bytChar > 57 And bytChar < 65) Or (bytChar > 90 And bytChar < 97) Or bytChar > 122 Then
               lngI = lngI - 1
            Else
               strTemp = strTemp & Mid$(.Text, lngN, 1)
            End If

         End If
      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "AlphaNumeric_NoSymbols"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub AlphaNumeric_PathValidOnly(ByRef rTextObj As TextBox, _
                                      Optional vblnAllowDrivePathSymbols As Boolean = False)

  Dim lngI    As Long
  Dim lngN    As Long
  Dim strTemp As String
  Dim bytChar As Byte

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      If vblnAllowDrivePathSymbols Then '// allow ":" and "\"

         For lngN = 1 To Len(.Text)
            bytChar = Asc(Mid$(.Text, lngN, 1))

            Select Case bytChar
            Case Is < 32, Is > 126, 34, 42, 47, 60, 62, 63, 124
               lngI = lngI - 1

            Case Else
               strTemp = strTemp & Mid$(.Text, lngN, 1)
            End Select

         Next lngN
      Else

         For lngN = 1 To Len(.Text)
            bytChar = Asc(Mid$(.Text, lngN, 1))

            Select Case bytChar '// don't allow ",*,/,:,<,>,?,\,|
            Case Is < 32, Is > 126, 34, 42, 47, 58, 60, 62, 63, 92, 124
               lngI = lngI - 1

            Case Else
               strTemp = strTemp & Mid$(.Text, lngN, 1)
            End Select

         Next lngN
      End If

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "AlphaNumeric_PathValidOnly"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub Alpha_All(ByRef rTextObj As TextBox)

  Dim lngI        As Long
  Dim lngN        As Long
  Dim strTemp     As String
  Dim bytChar     As Byte

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True
   mbytLastChar = 0

   With rTextObj
      lngI = .SelStart

      For lngN = 1 To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If mbytLastChar = vbKeySpace And bytChar = vbKeySpace And mblnRemoveExtraSpaces Then
            lngI = lngI - 1
         Else
            mbytLastChar = bytChar

            If bytChar >= 48 And bytChar <= 57 Then
               lngI = lngI - 1
            Else
               strTemp = strTemp & Mid$(.Text, lngN, 1)
            End If

         End If
      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "Alpha_All"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub Alpha_NoSymbols(ByRef rTextObj As TextBox)

  Dim lngI    As Long
  Dim lngN    As Long
  Dim strTemp As String
  Dim bytChar As Byte

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      For lngN = 1 To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If bytChar = 32 Or bytChar = 95 Then
            strTemp = strTemp & "_"
         Else

            If bytChar < 65 Or (bytChar > 90 And bytChar < 97) Or bytChar > 122 Then
               lngI = lngI - 1
            Else
               strTemp = strTemp & Mid$(.Text, lngN, 1)
            End If

         End If
      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "Alpha_NoSymbols"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub CaseLower(ByRef rTextObj As Object)

  Dim lngI        As Long
  Dim lngN        As Long
  Dim strTemp     As String
  Dim bytChar     As Byte

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True
   mbytLastChar = 0

   With rTextObj
      lngI = .SelStart

      If mblnRemoveExtraSpaces Then

         For lngN = 1 To Len(.Text)
            bytChar = Asc(Mid$(.Text, lngN, 1))

            If mbytLastChar = vbKeySpace And bytChar = vbKeySpace Then
               lngI = lngI - 1
            Else
               mbytLastChar = bytChar
               strTemp = strTemp & Mid$(.Text, lngN, 1)
            End If

         Next lngN

         If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)
         .Text = LCase$(strTemp)
      Else
         .Text = LCase$(.Text)
      End If

      .SelStart = lngI
   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "CaseLower"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub CaseProper(ByRef rTextObj As Object)

  Dim lngI As Long

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart
      .Text = ProperCaseString(.Text)
      .SelStart = lngI
   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "CaseProper"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub CaseUpper(ByRef rTextObj As Object)

  Dim lngI        As Long
  Dim lngN        As Long
  Dim strTemp     As String
  Dim bytChar     As Byte

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True
   mbytLastChar = 0

   With rTextObj
      lngI = .SelStart

      If mblnRemoveExtraSpaces Then

         For lngN = 1 To Len(.Text)
            bytChar = Asc(Mid$(.Text, lngN, 1))

            If mbytLastChar = vbKeySpace And bytChar = vbKeySpace Then
               lngI = lngI - 1
            Else
               mbytLastChar = bytChar
               strTemp = strTemp & Mid$(.Text, lngN, 1)
            End If

         Next lngN

         If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)
         .Text = UCase$(strTemp)
      Else
         .Text = UCase$(.Text)
      End If

      .SelStart = lngI
   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "CaseUpper"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub Class_Initialize()

   mlngBackColor_Normal = &H80000005
   mlngBackColor_GotFocus = &HEEFFEA   '&HEAFFFE
   mlngForeColor_Normal = &H80000008
   mlngForeColor_GotFocus = &H0&
   mblnAllowColorChange = False

End Sub

Private Sub Err_Handler(Optional ByVal vblnDisplayError As Boolean = True, _
                        Optional ByVal vstrErrNumber As String = vbNullString, _
                        Optional ByVal vstrErrDescription As String = vbNullString, _
                        Optional ByVal vstrModuleName As String = vbNullString, _
                        Optional ByVal vstrProcName As String = vbNullString)

  Dim strTemp As String
  Dim lngFN   As Long

   On Error Resume Next
   '// Purpose: Error handling - On Error

   '// Show Error Message
   If vblnDisplayError Then
      strTemp = "Error occured: "
      If LenB(vstrErrNumber) Then strTemp = strTemp & vstrErrNumber & vbNewLine Else strTemp = strTemp & vbNewLine
      If LenB(vstrErrDescription) Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
      If LenB(vstrModuleName) Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
      If LenB(vstrProcName) Then strTemp = strTemp & "Function: " & vstrProcName
      MsgBox strTemp, vbCritical, App.Title & " - ERROR"
   End If

   '// Write error log
   lngFN = FreeFile
   Open App.Path & "\ErrorLog.txt" For Append As #lngFN
   Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, _
         App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, _
         Environ("username"), Environ("computername")
   Close #lngFN

End Sub

Public Sub GotFocus(ByRef rTextObj As Object)

   On Error GoTo Err_Proc

   With rTextObj
      .SelStart = 0
      .SelLength = Len(rTextObj.Text)
      mstrUndoText = .Text

      If mblnAllowColorChange Then
         .BackColor = mlngBackColor_GotFocus
         .ForeColor = mlngForeColor_GotFocus
      End If

   End With

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "GotFocus"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub Integer_All(ByRef rTextObj As TextBox, _
                       Optional ByVal vblnLimitValue As Boolean = False)

  Dim lngI    As Long
  Dim lngN    As Long
  Dim strTemp As String
  Dim bytChar As Byte

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      For lngN = 1 To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If ((bytChar >= 48 And bytChar <= 57) Or (bytChar = 45 And lngN = 1)) Or (bytChar >= 48 And bytChar <= 57 And lngN > 1) Then
            strTemp = strTemp & Mid$(.Text, lngN, 1)
         Else
            lngI = lngI - 1
         End If

      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      If vblnLimitValue Then
         If Val(strTemp) > 32767 Then strTemp = "32767"
         If Val(strTemp) < -32767 Then strTemp = "-32767"
      End If

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "Integer_All"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub Integer_Neg(ByRef rTextObj As TextBox, _
                       Optional ByVal vblnLimitValue As Boolean = False)

  Dim lngI    As Long
  Dim lngN    As Long
  Dim bytX    As Byte
  Dim strTemp As String
  Dim bytChar As Byte

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      strTemp = "-"

      If Left(.Text, 1) <> "-" Then
         lngI = lngI + 1

         If Len(.Text) = 1 Then
            bytX = 2
            strTemp = strTemp & .Text
         Else
            bytX = 1
         End If

      Else
         bytX = 2
      End If

      For lngN = bytX To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If bytChar >= 48 And bytChar <= 57 Then
            strTemp = strTemp & Mid$(.Text, lngN, 1)
         Else
            lngI = lngI - 1
         End If

      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      If vblnLimitValue And Val(strTemp) < -32767 Then strTemp = "-32767"

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "Integer_Neg"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub Integer_Pos(ByRef rTextObj As TextBox, _
                       Optional ByVal vblnLimitValue As Boolean = False)

  Dim lngI    As Long
  Dim lngN    As Long
  Dim strTemp As String

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      For lngN = 1 To Len(.Text)

         Select Case Asc(Mid$(.Text, lngN, 1))
         Case 48 To 57
            strTemp = strTemp & Mid$(.Text, lngN, 1)

         Case Else
            lngI = lngI - 1
         End Select

      Next lngN

      If vblnLimitValue And Val(strTemp) > 32767 Then strTemp = "32767"

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "Integer_Pos"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub KeyPress(ByRef vintKeyAscii As Integer)

   '// No beep

   If vintKeyAscii = vbKeyReturn Then vintKeyAscii = 0

End Sub

Public Sub KeyUp(ByRef rTextObj As Object, _
                 ByRef rintKeyCode As Integer)

   '// vbKeyLeft 37 LEFT ARROW key
   '// vbKeyUp 38 UP ARROW key
   '// vbKeyRight 39 RIGHT ARROW key
   '// vbKeyDown 40 DOWN ARROW key
   '// vbKeyEscape 27 ESC key
   '// vbKeyReturn 13 ENTER key

   On Error GoTo Err_Proc

   Select Case rintKeyCode
   Case vbKeyReturn, vbKeyDown ', vbKeyRight
      rintKeyCode = 0
      SendKeys "{TAB}"

   Case vbKeyUp ', vbKeyLeft
      rintKeyCode = 0
      SendKeys "+{TAB}"

   Case vbKeyEscape  '// Undo command

      With rTextObj

         If Len(mstrUndoText) > 0 And .Text <> mstrUndoText Then
            .Text = mstrUndoText
            .SelStart = Len(.Text)
         End If

      End With
      rintKeyCode = 0
   End Select

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "KeyUp"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub LostFocus(ByRef rTextObj As Object)

   On Error GoTo Err_Proc

   mstrUndoText = vbNullString

   With rTextObj

      If mblnAllowColorChange Then
         .BackColor = mlngBackColor_Normal
         .ForeColor = mlngForeColor_Normal
      End If

   End With

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "LostFocus"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Function MaskedText(ByRef rTextObj As TextBox, _
                           ByVal vstrProfile As String) As Boolean

   '// Place in rTextObj_Change for immediate check */
   '// Place in rTextObj_KeyPress for 1 bytChar delay check */

  Dim lngCursorPlace     As Long
  Dim lngLen             As Long
  Dim lngPlace           As Long
  Dim strBuffer          As String
  Dim strText            As String
  Dim strChar            As String * 1
  Dim lngProfilePlace    As Long
  Dim lngGoodMatchColor  As Long

   If mblnWorking Then Exit Function
   mblnWorking = True

   On Local Error Resume Next

   If mblnAllowColorChange Then
      lngGoodMatchColor = mlngForeColor_GotFocus
   Else
      lngGoodMatchColor = mlngForeColor_Normal
   End If

   lngPlace = 1
   lngProfilePlace = 1

   With rTextObj
      '// If there are more characters than allowed then remove them

      If Len(.Text) > Len(vstrProfile) Then
         .Text = Left(.Text, Len(vstrProfile))
         .SelStart = Len(.Text)
         Beep
      End If

      strText = .Text
      lngLen = Len(.Text)
      lngCursorPlace = .SelStart
   End With

   '// The parser takes the pattern as the transition map. Starting
   '// at the beginning of the map, it compares the current character
   '// with the state of the parser

   Do While lngPlace <= lngLen
      strChar = Mid$(strText, lngPlace, 1)

      Select Case Mid$(vstrProfile, lngProfilePlace, 1)
         '// Check if vstrProfile calls for a numeric input */
      Case "#"

         If IsNumeric(strChar) Then
            '// Add the character to the buffer
            strBuffer = strBuffer & strChar
            '// Move to the next character
            lngPlace = lngPlace + 1
            '// Move to the next valid parser state
            lngProfilePlace = lngProfilePlace + 1
            '// Make sure we are indicating a valid transition state
            rTextObj.ForeColor = lngGoodMatchColor
         Else
            '// The character does not match the parser'strTemp state so
            '// indicate an invalid state and exit the parser
            rTextObj.ForeColor = vbRed
            Beep
            GoTo ExitCheckProfile
         End If

      Case "@"

         If Not IsNumeric(strChar) Then
            '// Add the character to the buffer
            strBuffer = strBuffer & strChar
            '// Move to the next character
            lngPlace = lngPlace + 1
            '// Move to the next valid parser state
            lngProfilePlace = lngProfilePlace + 1
            '// Make sure we are indicating a valid transition state
            rTextObj.ForeColor = lngGoodMatchColor
         Else
            '// The character does not match the parser'strTemp state so
            '// indicate an invalid state and exit the parser
            rTextObj.ForeColor = vbRed
            Beep
            GoTo ExitCheckProfile
         End If

      Case Else

         If strChar = Mid$(vstrProfile, lngProfilePlace, 1) Then
            '// If it Is here Then add the character to the buffer
            strBuffer = strBuffer & strChar
            '// Move to next character position
            lngPlace = lngPlace + 1
            '// Move to next parser state */
            lngProfilePlace = lngProfilePlace + 1
            '// Indicate a valid transition state to the user
            rTextObj.ForeColor = lngGoodMatchColor
         Else
            '// The required character is not present and in this case we
            '// insert it meeting the requirements of the parser state
            strBuffer = strBuffer & Mid$(vstrProfile, lngProfilePlace, 1)
            '// Shift the parser to the next state but stay with the current
            '// character to see if it matches the next state transition
            lngProfilePlace = lngProfilePlace + 1
            '// Move the input cursor over one to account for the added character
            lngCursorPlace = lngCursorPlace + 1
         End If

      End Select
   Loop

   '// If this is a function */

   If Len(strBuffer) = Len(vstrProfile) Then
      MaskedText = True
   Else
      MaskedText = False
   End If

   rTextObj.Text = strBuffer
   rTextObj.SelStart = lngCursorPlace

ExitCheckProfile:
   mblnWorking = False

End Function

Public Function mTrimString(ByVal vstrText As String)

   '//*************************************/
   '// Remove double spaces between words */
   '//*************************************/

  Dim strLastChar    As String
  Dim lngI           As Long
  Const C_strSpaceChar As String = " "

   On Error GoTo Err_Proc

   vstrText = Trim$(vstrText)
   If Len(vstrText) < 2 Then Exit Function
   lngI = 1
   strLastChar = Left(vstrText, lngI)

   Do

      If strLastChar = C_strSpaceChar And Mid$(vstrText, lngI, 1) = C_strSpaceChar Then
         vstrText = Left(vstrText, lngI - 1) & Mid$(vstrText, lngI + 1)
      Else
         strLastChar = Mid$(vstrText, lngI, 1)
         lngI = lngI + 1
      End If

   Loop Until lngI >= Len(vstrText)

   mTrimString = vstrText

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "mTrimString"
   Err.Clear
   Resume Exit_Proc

End Function

Public Sub Numeric_All(ByRef rTextObj As TextBox)

  Dim lngI       As Long
  Dim lngN       As Long
  Dim strTemp    As String
  Dim bytChar    As Byte
  Dim blnNoMatch As Boolean

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      For lngN = 1 To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If ((bytChar >= 48 And bytChar <= 57) Or (bytChar = 45 And lngN = 1) Or (bytChar = 46 And blnNoMatch = False)) Then
            If bytChar = 46 Then blnNoMatch = True
            strTemp = strTemp & Mid$(.Text, lngN, 1)
         Else
            lngI = lngI - 1
         End If

      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "Numeric_All"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub Numeric_FixedDec(ByRef rTextObj As Object, _
                            Optional ByVal vintDecPlaces As Integer = 2, _
                            Optional ByVal vblnShowTrailingZero As Boolean = True, _
                            Optional ByVal vblnAllowNegitive As Boolean = False)

  Dim lngCurrentPosition As Long
  Dim lngI               As Long
  Dim lngJ               As Long
  Dim strTemp            As String
  Dim strLeftString      As String
  Dim strRightString     As String
  Dim strTextString      As String

   On Local Error Resume Next

   If Len(rTextObj.Text) > 0 Then
      '// Save current cursor position
      lngCurrentPosition = rTextObj.SelStart
      strTextString = rTextObj.Text

      '// Make sure it is a number
      strTemp = vbNullString

      For lngI = 1 To Len(strTextString)

         Select Case Mid$(strTextString, lngI, 1)
         Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            strTemp = strTemp & Mid$(strTextString, lngI, 1)

         Case "."

            If vintDecPlaces > 0 Then
               strTemp = strTemp & Mid$(strTextString, lngI, 1)
            End If

         End Select
      Next lngI

      '// Check for negative number

      If vblnAllowNegitive Then
         If Left(strTextString, 1) = "-" Then strTemp = "-" & strTemp
      End If

      '// Remove double .. */
      lngI = InStr(strTemp, ".")
      lngJ = InStrRev(strTemp, ".")

      If lngJ <> lngI Then
         strTemp = Mid$(strTemp, 1, lngJ - 1) & Mid$(strTemp, lngJ + 1)
      End If

      '// Left side of decimal place */

      If lngI > 0 Then
         strLeftString = Left(strTemp, lngI - 1)
      Else
         strLeftString = strTemp
      End If

      '// Right side of decimal place */

      If lngI > 0 Then
         If vintDecPlaces > 0 Then
            strRightString = Mid$(strTemp, lngI, vintDecPlaces + 1)
            If vblnShowTrailingZero Then strRightString = Left(strRightString & String$(vintDecPlaces, "0"), vintDecPlaces + 1)
         Else
            strRightString = vbNullString
         End If

      Else

         If Right(strTemp, 1) = "." Then
            strRightString = "."
         Else
            strRightString = vbNullString
         End If

      End If

      '// Remove leading zeros

      If Val(strLeftString) > 0 Then
         lngI = Len(strLeftString)
         strLeftString = CStr(Val(strLeftString))

         If lngI <> Len(strLeftString) Then
            lngCurrentPosition = lngCurrentPosition - (lngI - Len(strLeftString))
         End If

      End If

      '// Combine left and right */
      rTextObj.Text = strLeftString & strRightString
      '// Restore cursor position to current
      rTextObj.SelStart = lngCurrentPosition

   End If

End Sub

Public Sub Numeric_Neg(ByRef rTextObj As TextBox)

  Dim lngI       As Long
  Dim lngN       As Long
  Dim bytX       As Byte
  Dim strTemp    As String
  Dim bytChar    As Byte
  Dim blnNoMatch As Boolean

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      strTemp = "-"

      If Left(.Text, 1) <> "-" Then
         lngI = lngI + 1

         If Len(.Text) = 1 Then
            bytX = 2
            strTemp = strTemp & .Text
         Else
            bytX = 1
         End If

      Else
         bytX = 2
      End If

      For lngN = bytX To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If ((bytChar >= 48 And bytChar <= 57) Or (bytChar = 46 And blnNoMatch = False)) Then
            If bytChar = 46 Then blnNoMatch = True
            strTemp = strTemp & Mid$(.Text, lngN, 1)
         Else
            lngI = lngI - 1
         End If

      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "Numeric_Neg"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub Numeric_Pos(ByRef rTextObj As TextBox)

  Dim lngI       As Long
  Dim lngN       As Long
  Dim strTemp    As String
  Dim bytChar    As Byte
  Dim blnNoMatch As Boolean

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      For lngN = 1 To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If ((bytChar >= 48 And bytChar <= 57) Or (bytChar = 46 And blnNoMatch = False)) Then
            If bytChar = 46 Then blnNoMatch = True
            strTemp = strTemp & Mid$(.Text, lngN, 1)
         Else
            lngI = lngI - 1
         End If

      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "Numeric_Pos"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub Numeric_WithSymbols(ByRef rTextObj As TextBox)

  Dim lngI       As Long
  Dim lngN       As Long
  Dim strTemp    As String
  Dim bytChar    As Byte

   On Error GoTo Err_Proc

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      For lngN = 1 To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If (bytChar >= 65 And bytChar <= 90) Or (bytChar >= 97 And bytChar <= 122) Then
            lngI = lngI - 1
         Else
            strTemp = strTemp & Mid$(.Text, lngN, 1)
         End If

      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "clsTextBox", "Numeric_WithSymbols"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Property Get pAllowColorChange() As Boolean

   pAllowColorChange = mblnAllowColorChange

End Property

Public Property Let pAllowColorChange(ByVal vNewValue As Boolean)

   mblnAllowColorChange = vNewValue

End Property

Public Property Get pBackColor_GotFocus() As Long

   pBackColor_GotFocus = mlngBackColor_GotFocus

End Property

Public Property Let pBackColor_GotFocus(ByVal oleValue As Long)

   mlngBackColor_GotFocus = oleValue

End Property

Public Property Get pBackColor_Normal() As Long

   pBackColor_Normal = mlngBackColor_Normal

End Property

Public Property Let pBackColor_Normal(ByVal oleValue As Long)

   mlngBackColor_Normal = oleValue

End Property

Public Property Get pForeColor_GotFocus() As Long

   pForeColor_GotFocus = mlngForeColor_GotFocus

End Property

Public Property Let pForeColor_GotFocus(ByVal oleValue As Long)

   mlngForeColor_GotFocus = oleValue

End Property

Public Property Get pForeColor_Normal() As Long

   pForeColor_Normal = mlngForeColor_Normal

End Property

Public Property Let pForeColor_Normal(ByVal oleValue As Long)

   mlngForeColor_Normal = oleValue

End Property

Public Property Get pRemoveExtraSpaces() As Boolean

   pRemoveExtraSpaces = mblnRemoveExtraSpaces

End Property

Public Property Let pRemoveExtraSpaces(ByVal vNewValue As Boolean)

   mblnRemoveExtraSpaces = vNewValue

End Property

Public Function ProperCaseString(ByVal vstrTemp As String) As String

  Dim lngX As Long

   '// Convert any string to the Proper Case

   On Local Error Resume Next

   If vstrTemp = vbNullString Then
      ProperCaseString = vbNullString
      Exit Function
   End If

   vstrTemp = UCase$(Left(vstrTemp, 1)) & LCase$(Mid$(vstrTemp, 2))

   If mblnRemoveExtraSpaces Then
      '// Look for and remove double spaces */
      lngX = InStr(vstrTemp, "  ")

      If lngX Then

         Do
            vstrTemp = Left(vstrTemp, lngX) & Mid$(vstrTemp, lngX + 2)
            lngX = InStr(lngX, vstrTemp, "  ")
            If lngX = 0 Or lngX + 1 > Len(vstrTemp) Then Exit Do
         Loop

      End If
   End If

   '// Look for space */
   lngX = InStr(vstrTemp, " ")

   If lngX Then

      Do
         Mid$(vstrTemp, lngX + 1, 1) = UCase$(Mid$(vstrTemp, lngX + 1, 1))
         lngX = lngX + 1
         lngX = InStr(lngX, vstrTemp, " ")
         If lngX = 0 Or lngX + 1 > Len(vstrTemp) Then Exit Do
      Loop

   End If
   '// Look for . */
   lngX = InStr(vstrTemp, ".")

   If lngX Then

      Do
         Mid$(vstrTemp, lngX + 1, 1) = UCase$(Mid$(vstrTemp, lngX + 1, 1))
         lngX = lngX + 1
         lngX = InStr(lngX, vstrTemp, ".")
         If lngX = 0 Or lngX + 1 > Len(vstrTemp) Then Exit Do
      Loop

   End If
   '// Look for / */
   lngX = InStr(vstrTemp, "/")

   If lngX Then

      Do
         Mid$(vstrTemp, lngX + 1, 1) = UCase$(Mid$(vstrTemp, lngX + 1, 1))
         lngX = lngX + 1
         lngX = InStr(lngX, vstrTemp, "/")
         If lngX = 0 Or lngX + 1 > Len(vstrTemp) Then Exit Do
      Loop

   End If
   '// Look for & */
   lngX = InStr(vstrTemp, "&")

   If lngX Then

      Do
         Mid$(vstrTemp, lngX + 1, 1) = UCase$(Mid$(vstrTemp, lngX + 1, 1))
         lngX = lngX + 1
         lngX = InStr(lngX, vstrTemp, "&")
         If lngX = 0 Or lngX + 1 > Len(vstrTemp) Then Exit Do
      Loop

   End If
   '// Look for Mc */
   lngX = InStr(vstrTemp, "Mc")

   If lngX Then

      Do
         Mid$(vstrTemp, lngX + 2, 1) = UCase$(Mid$(vstrTemp, lngX + 2, 1))
         lngX = lngX + 2
         lngX = InStr(lngX, vstrTemp, "Mc")
         If lngX = 0 Or lngX + 2 > Len(vstrTemp) Then Exit Do
      Loop

   End If
   '// Look for O' */
   lngX = InStr(vstrTemp, "O'")

   If lngX Then

      Do
         Mid$(vstrTemp, lngX + 2, 1) = UCase$(Mid$(vstrTemp, lngX + 2, 1))
         lngX = lngX + 2
         lngX = InStr(lngX, vstrTemp, "O'")
         If lngX = 0 Or lngX + 2 > Len(vstrTemp) Then Exit Do
      Loop

   End If
   '// Look for - */
   lngX = InStr(vstrTemp, "-")

   If lngX Then

      Do
         Mid$(vstrTemp, lngX + 1, 1) = UCase$(Mid$(vstrTemp, lngX + 1, 1))
         lngX = lngX + 1
         lngX = InStr(lngX, vstrTemp, "-")
         If lngX = 0 Or lngX + 1 > Len(vstrTemp) Then Exit Do
      Loop

   End If

   ProperCaseString = vstrTemp

End Function


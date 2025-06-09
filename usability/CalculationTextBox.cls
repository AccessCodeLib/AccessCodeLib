VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalculationTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: usability.CalculationTextBox
'---------------------------------------------------------------------------------------
'
' Allow calculation expression in a TextBox
'
' Author:
'     Josef Poetzl
'
' Example:
'
'   Code in Form class:
'
'      Private m_CalcTextBox As CalculationTextBox
'
'      Private Sub YourTextBoxControl_GotFocus()
'         Set m_CalcTextBox = CalculationTextBox(Me.YourTextBoxControl)
'      End Sub
'
'      Private Sub YourTextBoxControl_LostFocus()
'         Set m_CalcTextBox = Nothing
'      End Sub
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>usability/CalculationTextBox.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const EventProcedurePropertyTag As String = "[Event Procedure]"

Private WithEvents m_TextBox As TextBox
Attribute m_TextBox.VB_VarHelpID = -1
Private m_Format As String
Private m_ForeColor As Long
Private m_TextAlign As Byte
Private m_IsBoundControl As Boolean
Private m_InitialValue As Variant

Private m_CalcModeIsOn As Boolean
Private m_DisableRaiseErrorEvent As Boolean
Private m_DisableRaiseErrorEventResponse As CalculationErrorResponseMode

Public Enum CalculationErrorResponseMode
   [_CallErr_Undefined] = -1
   CalErr_Cancel = 1
   CalErr_Undo = 2
   CalErr_ShowMessage = 256
End Enum

Public Event Calculated(ByVal Expression As String, ByVal Value As Variant, ByRef Cancel As Boolean)
Public Event CalculationError(ByVal Expression As String, ByVal ErrMsg As String, Response As CalculationErrorResponseMode)

Public Function Init(ByVal TextBoxRef As TextBox) As CalculationTextBox
Attribute Init.VB_UserMemId = 0
'Attribute Init.VB_UserMemId = 0

   If Me Is CalculationTextBox Then ' return new instance
      
      With New CalculationTextBox
         Set Init = .Init(TextBoxRef)
      End With
      
      Exit Function
      
   End If
   
   Set m_TextBox = TextBoxRef
   SaveInitialProperties
   CheckEventProcedureProperties

   Set Init = Me
   
End Function

Private Sub SaveInitialProperties()
   m_Format = m_TextBox.Format
   m_ForeColor = m_TextBox.ForeColor
   m_TextAlign = m_TextBox.TextAlign
   m_IsBoundControl = Len(m_TextBox.ControlSource) > 0
   m_InitialValue = m_TextBox.Value
End Sub

Private Sub CheckEventProcedureProperties()
   
   If Len(m_TextBox.OnKeyPress) = 0 Then
      m_TextBox.OnKeyPress = EventProcedurePropertyTag
   End If
   
   If Len(m_TextBox.OnChange) = 0 Then
      m_TextBox.OnChange = EventProcedurePropertyTag
   End If
   
   If m_IsBoundControl Then
   
      If Len(m_TextBox.OnKeyDown) = 0 Then
         m_TextBox.OnKeyDown = EventProcedurePropertyTag
      End If
   
   Else
   
      If Len(m_TextBox.BeforeUpdate) = 0 Then
         m_TextBox.BeforeUpdate = EventProcedurePropertyTag
      End If
   
      If Len(m_TextBox.AfterUpdate) = 0 Then
         m_TextBox.AfterUpdate = EventProcedurePropertyTag
      End If
      
   End If
   
End Sub

Public Property Get TextBox() As TextBox
   Set TextBox = m_TextBox
End Property

Private Sub m_TextBox_KeyPress(KeyAscii As Integer)

   Const KeyAsciiEquals As Integer = 61 ' Asc("=") = 61
   Const KeyAsciiEscape As Integer = 27 ' Asc("=") = 61
   
   Dim TbText As String
   
   If KeyAscii = KeyAsciiEquals Then
      If Len(Trim(m_TextBox.Text)) = 0 Then
         ActivateCalcMode
      ElseIf m_TextBox.SelLength = Len(m_TextBox.Text) Then
         ActivateCalcMode
      ElseIf m_TextBox.SelStart = 0 And m_TextBox.SelLength = 0 Then
         ActivateCalcMode
         m_TextBox.SelStart = 0
      End If
   ElseIf KeyAscii = KeyAsciiEscape Then
      DeactivateCalcMode
   End If
End Sub

Private Sub m_TextBox_Change()
   If m_CalcModeIsOn Then
      If Left(Trim(m_TextBox.Text), 1) <> "=" Then
         DeactivateCalcMode
      End If
   End If
End Sub

Private Sub ActivateCalcMode()
   m_TextBox.Format = vbNullString
   m_TextBox.ForeColor = RGB(0, 0, 255)
   m_TextBox.TextAlign = 1
   m_CalcModeIsOn = True
End Sub

Private Sub DeactivateCalcMode()
   m_CalcModeIsOn = False
   m_TextBox.Format = m_Format
   m_TextBox.ForeColor = m_ForeColor
   m_TextBox.TextAlign = m_TextAlign
End Sub

Private Sub m_TextBox_KeyDown(KeyCode As Integer, Shift As Integer)
   
   Dim Result As Variant
   Dim Response As CalculationErrorResponseMode
   
   If m_CalcModeIsOn Then
   
      If Not m_IsBoundControl Then
         Exit Sub ' => see AfterUpdate
      End If
   
      If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
         Result = Calculate(GetCalcStatement(m_TextBox.Text), Response)
         If (Response And CalErr_Cancel) = CalErr_Cancel Then
            KeyCode = 0
         ElseIf (Response And CalErr_Undo) = CalErr_Undo Then
            m_TextBox.Text = m_InitialValue
            DeactivateCalcMode
         Else
            m_TextBox.Text = Result
            DeactivateCalcMode
         End If
      End If
      
   End If
   
End Sub

Private Sub m_TextBox_BeforeUpdate(Cancel As Integer)

   Dim Response As CalculationErrorResponseMode

   If m_CalcModeIsOn Then
   
      If m_IsBoundControl Then
         Exit Sub ' => see KeyDown
      End If
   
      Calculate GetCalcStatement(m_TextBox.Value), Response, True
      If (Response And CalErr_Cancel) = CalErr_Cancel Then
         Cancel = True
      ElseIf (Response And CalErr_Undo) = CalErr_Undo Then
         m_DisableRaiseErrorEvent = True
         m_DisableRaiseErrorEventResponse = Response
      End If
  
   End If
   
End Sub

Private Sub m_TextBox_AfterUpdate()

   Dim Result As Variant
   Dim Response As CalculationErrorResponseMode
   
   If m_CalcModeIsOn Then
   
      If m_IsBoundControl Then
         Exit Sub ' => see KeyDown
      End If
   
      Result = Calculate(GetCalcStatement(m_TextBox.Value), Response)
      If (Response And CalErr_Cancel) = CalErr_Cancel Then
         m_TextBox.SetFocus
         m_TextBox.SelStart = Len(m_TextBox.Text)
      ElseIf (Response And CalErr_Undo) = CalErr_Undo Then
         m_TextBox.Value = m_InitialValue
         DeactivateCalcMode
      Else
         m_TextBox.Value = Result
         DeactivateCalcMode
      End If
      
   End If

End Sub

Private Function GetCalcStatement(ByVal InputValue As Variant) As String

   InputValue = Trim(InputValue)
   InputValue = Mid(InputValue, 2)
   GetCalcStatement = InputValue

End Function

Private Function Calculate(ByVal CalcStatement As String, ByRef Response As CalculationErrorResponseMode, Optional ByVal IgnoreShowMessageIfUndo As Boolean = False) As Variant

   Dim Result As Variant
   Dim ErrMsg As String
   Dim Cancel As Boolean

   CalcStatement = Replace(CalcStatement, DecimalSeparator, ".")
   If Not TryEval(CalcStatement, Result, ErrMsg) Then
   
      Response = [_CallErr_Undefined] '
      
      If m_DisableRaiseErrorEvent Then
         m_DisableRaiseErrorEvent = False
         Response = m_DisableRaiseErrorEventResponse
         m_DisableRaiseErrorEventResponse = 0
      Else
         RaiseEvent CalculationError(CalcStatement, ErrMsg, Response)
      End If
      If Response = [_CallErr_Undefined] Then
         Response = CalErr_Cancel + CalErr_ShowMessage
      End If
      
      If (Response And CalErr_Undo) = CalErr_Undo Then
         If IgnoreShowMessageIfUndo Then
            Response = Response Or CalErr_ShowMessage Xor CalErr_ShowMessage
         End If
      End If
      
      If (Response And CalErr_ShowMessage) = CalErr_ShowMessage Then
         MsgBox ErrMsg, vbExclamation, "Error in expression: " & CalcStatement
         Response = Response - CalErr_ShowMessage
      End If
      
      If Response > 0 Then
         Exit Function
      End If
      
   End If
   
   RaiseEvent Calculated(CalcStatement, Result, Cancel)
   If Cancel Then
      Response = CalErr_Undo
   End If
   
   Calculate = Result
   
End Function

Private Function TryEval(ByVal CalcStatement As String, ByRef Result As Variant, ByRef ErrMsg As String) As Boolean

   On Error GoTo HandleErr
   
   Result = Eval(CalcStatement)
   
   TryEval = True
   
ExitHere:
   Exit Function
   
HandleErr:
   TryEval = False
   ErrMsg = Err.Description
   Resume ExitHere

End Function

Private Property Get DecimalSeparator() As String

   Static DecChar As String
   Dim CheckString As String
   
   If Len(DecChar) = 0 Then
      CheckString = Trim(CStr(1.2))
      DecChar = Mid(CheckString, 2, 1)
   End If
   
   DecimalSeparator = DecChar

End Property

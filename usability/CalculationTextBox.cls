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

Private m_CalcModeIsOn As Boolean

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
   
      If Len(m_TextBox.AfterUpdate) = 0 Then
         m_TextBox.AfterUpdate = EventProcedurePropertyTag
      End If
      
   End If
   
End Sub

Private Sub m_TextBox_KeyPress(KeyAscii As Integer)
   Const KeyAsciiEquals As Integer = 61 ' Asc("=") = 61
   Const KeyAsciiEscape As Integer = 27 ' Asc("=") = 61
   If KeyAscii = KeyAsciiEquals Then
      If Len(Trim(m_TextBox.Text)) = 0 Then
         ActivateCalcMode
      ElseIf m_TextBox.SelLength = Len(m_TextBox.Text) Then
         ActivateCalcMode
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
   
   If m_CalcModeIsOn Then
   
      If Not m_IsBoundControl Then
         Exit Sub ' => see AfterUpdate
      End If
   
      If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
         m_TextBox.Text = Calculate(GetCalcStatement(m_TextBox.Text))
         DeactivateCalcMode
      End If
      
   End If
   
End Sub

Private Sub m_TextBox_AfterUpdate()

   If m_CalcModeIsOn Then
   
      If m_IsBoundControl Then
         Exit Sub ' => see KeyDown
      End If
   
      m_TextBox.Value = Calculate(GetCalcStatement(m_TextBox.Value))
      DeactivateCalcMode
      
   End If

End Sub

Private Function GetCalcStatement(ByVal InputValue As Variant) As String

   InputValue = Trim(InputValue)
   InputValue = Mid(InputValue, 2)
   GetCalcStatement = InputValue

End Function

Private Function Calculate(ByVal CalcStatement As String) As Variant
   CalcStatement = Replace(CalcStatement, DecimalSeparator, ".")
   Calculate = Eval(CalcStatement)
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

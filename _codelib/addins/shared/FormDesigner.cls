VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: _codelib.addins.shared.FormDesigner
'---------------------------------------------------------------------------------------
'
' Form design class
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/FormDesigner.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_Form As Form

Public Property Get Form() As Form
   Set Form = m_Form
End Property

Public Property Set Form(ByVal FormRef As Form)
   Set m_Form = FormRef
End Property

Public Property Get FormName() As String
   If m_Form Is Nothing Then
      FormName = vbNullString
   Else
      FormName = m_Form.Name
   End If
End Property

Public Property Let FormName(ByVal NewValue As String)
   Set Me.Form = GetFormInDesignMode(NewValue)
End Property

Private Function GetFormInDesignMode(ByVal FormName As String) As Form
   
   Dim FormIsLoaded As Boolean
   
   FormIsLoaded = CurrentProject.AllForms(FormName).IsLoaded
   
   If FormIsLoaded Then
      If Application.Forms(FormName).CurrentView <> 0 Then
         DoCmd.Close acForm, FormName
         FormIsLoaded = False
      End If
   End If
   
   If Not FormIsLoaded Then
      DoCmd.OpenForm FormName, acDesign, , , , acWindowNormal
   End If
   
   Set GetFormInDesignMode = Application.Forms(FormName)

End Function

Public Function AddControl(ByVal ControlType As AcControlType, ByVal ControlName As String, ByVal Section As AcSection, _
                           Optional ByVal LeftPos As Variant, Optional ByVal TopPos As Variant, _
                           Optional ByVal ControlWidth As Variant, Optional ByVal ControlHeight As Variant, _
                           Optional ByVal DistanceBetweenControls As Long = 72, _
                           Optional ByVal InsertLabelControl As Boolean = True, Optional ByVal LabelControlCaption As String) As Control
   
   Dim ctl As Control
   Dim LabelCtl As Label
   
   Dim NextFreeLeftPos As Long, NextFreeTopPos As Long
   If IsMissing(LeftPos) Or IsMissing(TopPos) Then
      FindNextControlPos Section, NextFreeLeftPos, NextFreeTopPos, DistanceBetweenControls
      If IsMissing(LeftPos) Then LeftPos = NextFreeLeftPos
      If IsMissing(TopPos) Then TopPos = NextFreeTopPos
   End If
   
   Set ctl = Application.CreateControl(m_Form.Name, ControlType, Section, , , LeftPos, TopPos, ControlWidth, ControlHeight)
   
   If Len(ControlName) > 0 Then
      ControlName = GetCheckedControlName(ControlName)
      ctl.Name = ControlName
   End If
   
   Set AddControl = ctl
   
   If InsertLabelControl Then
   
      If Len(LabelControlCaption) = 0 Then
         LabelControlCaption = ControlName
      End If
   
      Set LabelCtl = Application.CreateControl(m_Form.Name, acLabel, Section, ctl.Name, , LeftPos, TopPos, ctl.Width, ctl.Height)
      LabelCtl.Name = GetCheckedControlName("lab" & ctl.Name)
      LabelCtl.Caption = LabelControlCaption
      ctl.Top = LabelCtl.Top + LabelCtl.Height

   End If
   
End Function

Private Function GetCheckedControlName(ByVal ControlName As String) As String
   
   Dim CheckedControlName As String
   Dim i As Long
   
   CheckedControlName = ControlName
   Do While ControlExists(CheckedControlName)
      i = i + 1
      CheckedControlName = ControlName & i
   Loop
   
   GetCheckedControlName = CheckedControlName
   
End Function

Private Function ControlExists(ByVal ControlName As String) As Boolean
   
   Dim ctl As Control
   For Each ctl In m_Form.Controls
      If ctl.Name = ControlName Then
         ControlExists = True
         Exit Function
      End If
   Next

End Function

Private Sub FindNextControlPos(ByVal ControlSection As AcSection, ByRef LeftPos As Long, ByRef TopPos As Long, Optional ByVal DistanceBetweenControls As Long = 72)

   Dim ctl As Control
   Dim MinLeft As Long
   Dim MaxTop As Long
   Dim ErrCnt As Long
   
On Error GoTo HandleErr
   
   If m_Form.Section(ControlSection).Controls.Count > 0 Then
      MinLeft = m_Form.InsideWidth
      For Each ctl In m_Form.Section(ControlSection).Controls
         If MinLeft > ctl.Left Then
            MinLeft = ctl.Left
         End If
         If MaxTop < (ctl.Top + ctl.Height) Then
            MaxTop = (ctl.Top + ctl.Height)
         End If
      Next
   End If
   
   If MinLeft = 0 Then
      MinLeft = DistanceBetweenControls
   End If
   LeftPos = MinLeft
   
   TopPos = MaxTop + DistanceBetweenControls

   Exit Sub
   
HandleErr:
   If Err.Number = 2462 Then 'auf Section kann nicht zugegriffen werden
   
      Select Case ControlSection
         Case acHeader, acFooter
            m_Form.SetFocus
            DoCmd.RunCommand acCmdFormHdrFtr
         Case acPageHeader, acPageFooter
            DoCmd.RunCommand acCmdPageHdrFtr
      End Select
      
      If ErrCnt < 2 Then
         ErrCnt = ErrCnt + 1
         Resume
      End If
      
   End If

   Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

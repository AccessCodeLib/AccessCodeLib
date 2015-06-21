VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormSortingExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: FormSortingControl
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erstellt einen OrderBy-String für ein Endlos-Formular
' </summary>
' <remarks>
' </remarks>
' \ingroup form
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/sort/FormSortingExtension.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/sort/FormSortingControl.cls</use>
'  <use>form/sort/FormSortingControlEventBridge.cls</use>
'  <use>api/winapi/WinApiLayoutTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const EventProcedureString = "[Event Procedure]"

Private m_SortingControls As Collection
Private WithEvents m_EventBridge As FormSortingControlEventBridge

Private m_UseDesc As Boolean
Private m_ActiveSortingControl As Control

Public Event OrderByChanged(ByVal OrderByString As String)

'--------------------------------------------------
' Variablen + API für Positionierung des "Sortier-Symbols"
'
Private m_SortSymbolLabel As Label
Private m_SortSymbolAscChar As String
Private m_SortSymbolDescChar As String
Private WithEvents m_SortSymbolForm As Form
Attribute m_SortSymbolForm.VB_VarHelpID = -1


Private Declare Function FindWindowEx Lib "USER32.DLL" Alias "FindWindowExA" ( _
         ByVal hWnd1 As Long, _
         ByVal hWnd2 As Long, _
         ByVal lpsz1 As String, _
         ByVal lpsz2 As String _
      ) As Long
   

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, ByRef lpRect As Rect) As Long
' / API
'---------------------------------------------------


Public Function Add(ByVal ControlRef As Control, ByVal OrderByStatement As String) As FormSortingControl

   With New FormSortingControl
      
      Set .Control = ControlRef
      .OrderByStatement = OrderByStatement
      Set .EventBridge = m_EventBridge
      Set Add = .This
      
      m_SortingControls.Add .This, ControlRef.Name
      
   End With
   
End Function

Private Sub Class_Initialize()
   Set m_SortingControls = New Collection
   Set m_EventBridge = New FormSortingControlEventBridge
End Sub

Private Sub m_EventBridge_SortingControlClick(ByVal SortingControl As FormSortingControl)
   
   If SortingControl.Control Is m_ActiveSortingControl Then
      m_UseDesc = Not m_UseDesc
   Else
      m_UseDesc = False
      Set m_ActiveSortingControl = SortingControl.Control
   End If
   
   ChangeSortSymbol SortingControl.Control, m_UseDesc
   
   RaiseEvent OrderByChanged(SortingControl.GetOrderByString(m_UseDesc))
   
End Sub

Private Sub ChangeSortSymbol(ByVal ControlRef As Control, ByVal SortDesc As Boolean)
   
   If Not (m_SortSymbolLabel Is Nothing) Then
      ChangeSortSymbolLabel ControlRef, SortDesc
      Exit Sub
   End If
   
   '@TODO: Sortierkennung an Labeltext, Formatierung usw. erkennbar machen
   
End Sub


'-----------------------------------------------------
' Prozeduren für Positionierung des "Sortier-Symbols"
'
' @TODO: Fix bugs
' @TODO: Refacturing!
'
Public Sub InitSortSymbolLabel(ByVal SortSymbolRef As Label, SortAscChar As String, SortDescChar As String, _
               Optional ByVal StartSortingcontrol As Control, Optional StartSortDesc As Boolean = False)

   Set m_SortSymbolLabel = SortSymbolRef
   m_SortSymbolAscChar = SortAscChar
   m_SortSymbolDescChar = SortDescChar

   If Not (StartSortingcontrol Is Nothing) Then
      ChangeSortSymbol StartSortingcontrol, StartSortDesc
   End If
   
   Set m_SortSymbolForm = m_SortSymbolLabel.Parent
   If Len(m_SortSymbolForm.OnResize) = 0 Then
      m_SortSymbolForm.OnResize = EventProcedureString
   End If
   
End Sub

Public Sub RepaintSortSymbol()
   SetSortSymbolPos m_ActiveSortingControl
End Sub

Private Sub ChangeSortSymbolLabel(ByVal ControlRef As Control, ByVal SortDesc As Boolean)
   
   SetSortSymbolPos ControlRef
   
   If SortDesc Then
      m_SortSymbolLabel.Caption = m_SortSymbolDescChar
   Else
      m_SortSymbolLabel.Caption = m_SortSymbolAscChar
   End If
   
End Sub

Private Sub SetSortSymbolPos(ByVal ControlRef As Control)
   Dim CtlRight As Long
   
   If m_SortSymbolLabel Is Nothing Then Exit Sub
   If ControlRef Is Nothing Then Exit Sub
   
   CtlRight = GetLeftPosAndWidth(ControlRef)
   
   If CtlRight > ControlRef.Parent.InsideWidth Then
      ControlRef.Parent.InsideWidth = CtlRight
   End If
   m_SortSymbolLabel.Left = CtlRight - m_SortSymbolLabel.Width
End Sub

Private Function GetLeftPosAndWidth(ByVal ControlRef As Control) As Long

   Dim frm As Form
   Set frm = ControlRef.Parent
   
   If (ControlRef.Layout <> AcLayoutType.acLayoutNone) Then
      GetLeftPosAndWidth = GetLayoutLeftPosAndWidth(ControlRef)
   ElseIf ControlRef.HorizontalAnchor = acHorizontalAnchorRight Or ControlRef.HorizontalAnchor = acHorizontalAnchorBoth Then
      GetLeftPosAndWidth = ControlRef.Left + ControlRef.Width + frm.WindowWidth - frm.Width - ScrollBarsWidth(frm) - 2 * GetScrollbarWidth() - 50
   Else
      GetLeftPosAndWidth = ControlRef.Left + ControlRef.Width
   End If
   
End Function

Private Function GetScrollbarWidth() As Long
   With New WinApiLayoutTools
      GetScrollbarWidth = .GetScrollbarWidth
   End With
End Function

Private Function ScrollbarIsVisible(ByVal frm As Form) As Boolean
   
   If (frm.ScrollBars And 2) = 0 Then
      ScrollbarIsVisible = False
      Exit Function
   End If
   
   Dim scHwnd As Long
   scHwnd = FindWindowEx(frm.Hwnd, 0&, "NUIScrollbar", "Vertikal") '
   If scHwnd = 0 Then
      scHwnd = FindWindowEx(frm.Hwnd, 0&, "NUIScrollbar", "Vertical")
   End If
   
   If scHwnd = 0 Then
      ScrollbarIsVisible = False
      Exit Function
   End If
   
   Dim rscb As Rect, rfrm As Rect
   
   GetWindowRect frm.Hwnd, rfrm
   GetWindowRect scHwnd, rscb
   
   If rfrm.Right < rscb.Right Then ' Scrollbar nicht sichtbar
      ScrollbarIsVisible = False
      Exit Function
   End If
   
   ScrollbarIsVisible = True
   
End Function

Private Function ScrollBarsWidth(ByVal frm As Form) As Long
   
   If ScrollbarIsVisible(frm) Then
      ScrollBarsWidth = GetScrollbarWidth()
      Exit Function
   End If
   
   ScrollBarsWidth = 0
   
End Function

Private Function GetLayoutLeftPosAndWidth(ByVal ControlRef As Control) As Long

   Dim AnchorWidthSum As Double
   Dim LeftSum As Double

   Dim ctl As Control
   Dim frm As Form
   Dim LayoutUsed As Boolean
   Dim LayoutId As Long
   
   Set frm = ControlRef.Parent
   Dim CtlDesignLeft As Long
   
   CtlDesignLeft = ControlRef.Left
   
   LayoutUsed = (ControlRef.Layout <> AcLayoutType.acLayoutNone)
   If LayoutUsed Then LayoutId = ControlRef.LayoutId
   
   For Each ctl In frm.Section(ControlRef.Section).Controls
   
      If ctl.Left < CtlDesignLeft Then
         If ctl.LayoutId = LayoutId Then
            If ctl.HorizontalAnchor = acHorizontalAnchorBoth Then
               LeftSum = LeftSum + ctl.Width
            End If
         End If
      End If
      
      If ctl.HorizontalAnchor = acHorizontalAnchorBoth Then
         AnchorWidthSum = AnchorWidthSum + ctl.Width
      End If
      
   Next
   
   Dim LeftPos As Long
   Dim DiffW As Long
   DiffW = frm.WindowWidth - frm.Width - ScrollBarsWidth(frm) - 2 * GetScrollbarWidth()
   LeftPos = ControlRef.Left + DiffW * LeftSum / AnchorWidthSum - 100
   
   If ControlRef.HorizontalAnchor = acHorizontalAnchorBoth Then
      GetLayoutLeftPosAndWidth = LeftPos + ControlRef.Width + ControlRef.Width * DiffW / AnchorWidthSum
   Else
      GetLayoutLeftPosAndWidth = LeftPos + ControlRef.Width
   End If
   
End Function

Private Sub m_SortSymbolForm_Resize()
   RepaintSortSymbol
End Sub

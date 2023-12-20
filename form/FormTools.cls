VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: FormTools
'---------------------------------------------------------------------------------------
' Procedure collection for forms
'
' Author:
'     Josef Poetzl
'
' Remarks:
'   Simplifies the control of Access forms
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/FormTools.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const WM_VSCROLL = &H115
Private Const WM_SWP_NOZORDER = &H4
Private Const WM_SWP_SHOWWINDOW = &H40
Private Const WM_LOGPIXELSY = 90

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

#If VBA7 Then

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                              ByVal Hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, lParam As Any) As Long
                              
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hDC As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As Rect) As Long

#Else

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                              ByVal Hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, lParam As Any) As Long
                              
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As Rect) As Long

#End If

Public Sub OpenForm(ByVal FormName As String, _
               Optional ByVal sFilter As String = vbNullString, _
               Optional ByVal sOpenarg As String = vbNullString, _
               Optional ByVal WindowMode As AcWindowMode = AcWindowMode.acWindowNormal)

   DoCmd.OpenForm FormName, acNormal, , sFilter, , WindowMode, sOpenarg

End Sub

Public Function CloseForm(Optional ByVal FormName As String)
   
   DoCmd.Close acForm, FormName
   
End Function

Public Function OpenNewRecord(ByVal FormName As String, _
                 Optional ByVal sOpenarg As String = vbNullString, _
                 Optional ByVal WindowMode As AcWindowMode = AcWindowMode.acWindowNormal)
   
   If IsLoaded(FormName) Then
      FormTools.CloseForm FormName
   End If
   DoCmd.OpenForm FormName, acNormal, , , acFormAdd, WindowMode, sOpenarg

End Function

Public Sub GotoBookmark(ByRef FormRef As Form, ByVal BookmarkID_Name As String, ByVal BookmarkID_Value As Variant)

   Dim rst As Object
   Dim strCriteria As String

   If IsNull(BookmarkID_Value) Then
      Exit Sub
   End If

   If IsNumeric(BookmarkID_Value) Then
      strCriteria = BookmarkID_Name & "=" & BookmarkID_Value
   Else
      strCriteria = BookmarkID_Name & "='" & BookmarkID_Value & "'"
   End If
   
   Set rst = FormRef.Recordset.Clone
   If TypeOf rst Is DAO.Recordset Then
      With rst
         If Not .EOF Then
         .FindFirst strCriteria
         If .NoMatch = False Then
            FormRef.Bookmark = .Bookmark
         End If
         End If
      End With
   Else 'If TypeOf rst Is ADODB.Recordset Then
      With rst
         .Find strCriteria, , 1 'adSearchForward
         If Not .EOF Then
            FormRef.Bookmark = .Bookmark
         End If
      End With
   End If
   Set rst = Nothing
   
End Sub

Public Sub SetFormMinSize(ByVal FormRef As Form, ByVal lStaticBaseInsideHeigth As Long, ByVal lBaseInsideWidth As Long)
   
   Dim bolScrollbarChanged As Boolean
   Dim lngOptimalHeight As Long
   Dim lngOptimalWidth As Long
   
   If FormRef Is Nothing Then
      Exit Sub
   End If
   
   lngOptimalWidth = FormRef.Width - 420 * ((FormRef.ScrollBars And 2) = 2)
   If Abs(FormRef.InsideWidth - lngOptimalWidth) > 50 Then
      If FormRef.InsideWidth < lngOptimalWidth Then
         If (FormRef.ScrollBars And 1) = 0 Then
            bolScrollbarChanged = True
            FormRef.ScrollBars = FormRef.ScrollBars Or 1
         End If
      Else
         If (FormRef.ScrollBars And 1) = 1 Then
            bolScrollbarChanged = True
            FormRef.ScrollBars = -2 * ((FormRef.ScrollBars And 2) = 2)
         End If
      End If
      
   End If
   
   If bolScrollbarChanged Then
      If FormRef.ScrollBars = 0 Then
         ResizeWindow FormRef, FormRef.Width, lStaticBaseInsideHeigth
      Else
         ResizeWindow FormRef, , lStaticBaseInsideHeigth - 420 * ((FormRef.ScrollBars And 1) = 1)
      End If
   End If
   
   lngOptimalHeight = FormRef.Section(0).Height + FormRef.Section(1).Height + FormRef.Section(2).Height
   If Abs(FormRef.InsideHeight - lngOptimalHeight) > 50 Then
      If FormRef.InsideHeight < lngOptimalHeight Then
         If (FormRef.ScrollBars And 2) = 0 Then
            bolScrollbarChanged = True
            FormRef.ScrollBars = FormRef.ScrollBars Or 2
         End If
      Else
         If (FormRef.ScrollBars And 2) = 2 Then
            bolScrollbarChanged = True
            FormRef.ScrollBars = -((FormRef.ScrollBars And 1) = 1)
         End If
      End If
      
   End If
   
   If bolScrollbarChanged Then
      If FormRef.ScrollBars = 0 Then
         ResizeWindow FormRef, FormRef.Width, lStaticBaseInsideHeigth
      Else
         ResizeWindow FormRef, lBaseInsideWidth - 425 * ((FormRef.ScrollBars And 2) = 2)
      End If
   End If
   bolScrollbarChanged = False

End Sub

Public Sub ResizeWindow(ByVal FormRef As Form, Optional ByVal newWidth As Long = -1, Optional ByVal NewHeight As Long = -1)
'Einheit TWIPS !!!

   Dim rctForm As Rect
   Dim rctMain As Rect
   
   Dim intNewWith As Integer
   Dim intNewHeight As Integer

   If FormRef Is Nothing Then
      Exit Sub
   End If
   
   If GetWindowRect(FormRef.Hwnd, rctForm) = False Then
      Exit Sub
   End If
   
   If GetWindowRect(Application.hWndAccessApp, rctMain) = False Then
      Exit Sub
   End If
   
   Dim DisplayLogPixelsY As Integer
   Dim WMTwipsPerPixel As Integer
   Dim hDesktopWnd As Long
   Dim hDCcaps As Long
   Dim iRtn As Integer

   hDesktopWnd = GetDesktopWindow() 'get handle to desktop
   hDCcaps = GetDC(hDesktopWnd) 'get display context for desktop
   DisplayLogPixelsY = GetDeviceCaps(hDCcaps, WM_LOGPIXELSY)
   iRtn = ReleaseDC(hDesktopWnd, hDCcaps) 'release display contex
   WMTwipsPerPixel = 1440 / DisplayLogPixelsY
   
   With rctForm
      If newWidth < 0 Then
         intNewWith = .Right - .Left
      Else
         intNewWith = Fix(newWidth / WMTwipsPerPixel)
      End If

      If NewHeight < 0 Then
         intNewHeight = .Bottom - .Top
      Else
         intNewHeight = Fix(NewHeight / WMTwipsPerPixel)
      End If

      Call SetWindowPos(FormRef.Hwnd, 0, .Left - rctMain.Left, 0, intNewWith, intNewHeight, WM_SWP_NOZORDER Or WM_SWP_SHOWWINDOW)
      
   End With
   
End Sub

Public Function IsLoaded(ByVal FormName As String) As Boolean
  IsLoaded = CurrentProject.AllForms.Item(FormName).IsLoaded
End Function

Public Sub RequeryData(ByVal FormRef As Form, Optional ByVal BookmarkFieldName As String = vbNullString)

   Dim varIDValue As Variant
   Dim lngPos As Long, lngPosDiff As Long
   Dim lngDirection As Long
   Dim i As Long, lngRet As Long
   Dim bolIsListForm As Boolean
   
   Dim SetPainting As Boolean

On Error GoTo 0

   If Len(BookmarkFieldName) = 0 Then
      FormRef.Requery
      Exit Sub
   End If

   bolIsListForm = IsListForm(FormRef)

On Error Resume Next
   If bolIsListForm Then
   With FormRef
      lngPos = (.CurrentSectionTop - .Section(acHeader).Height) \ .Section(0).Height
   End With
   End If
   
   varIDValue = Null
   If Len(BookmarkFieldName) > 0 Then
       varIDValue = FormRef.Controls(BookmarkFieldName).Value
   End If
   
   If Err.Number <> 0 Then
      Err.Clear
   End If

On Error GoTo HandleErr

   FormRef.Requery

   If IsNull(varIDValue) = False Then
      'zu Bookmark springen
      
      If FormRef.Painting = True Then
         FormRef.Painting = False
         SetPainting = True
      End If
      
      GotoBookmark FormRef, BookmarkFieldName, varIDValue

      'letzte Position wieder herstellen
      If bolIsListForm Then
On Error Resume Next
         With FormRef
            lngPosDiff = lngPos - ((.CurrentSectionTop - .Section(acHeader).Height) \ .Section(0).Height)
         End With
On Error GoTo HandleErr
         If lngPosDiff <> 0 Then
            If lngPosDiff > 0 Then
               lngDirection = SB_LINEUP
            Else
               lngDirection = SB_LINEDOWN
               lngPosDiff = Abs(lngPosDiff)
            End If
            For i = 1 To lngPosDiff
               lngRet = SendMessage(FormRef.Hwnd, WM_VSCROLL, lngDirection, 0&)
            Next i
         End If
      End If
   End If

ExitHere:
On Error Resume Next
   If SetPainting Then
      FormRef.Painting = True
   End If
   Exit Sub

HandleErr:
   SetPainting = True
   Resume ExitHere

End Sub

Private Function IsListForm(ByVal FormRef As Form) As Boolean

   Dim bolIsListForm As Boolean

   With FormRef
      Select Case .CurrentView
         Case 0 'Entwurfsansicht
            bolIsListForm = False
         Case 1 'Formularansicht
            bolIsListForm = (.DefaultView = 1)
         Case 2 'Datenblatt
            bolIsListForm = True
      End Select
   End With

   IsListForm = bolIsListForm

End Function

Public Sub SetScrollBars(ByVal FormRef As Form, Optional ByVal RecordCount As Long = -1)

On Error Resume Next

   If RecordCount = -1 Then
      RecordCount = FormRef.Recordset.RecordCount
      
      If FormRef.AllowAdditions Then
         RecordCount = RecordCount + 1
      End If
      
   End If
   
   Dim lngDetailHeight As Long
   lngDetailHeight = FormRef.Section(acDetail).Height
   
   Dim lngHeaderHeight As Long
   lngHeaderHeight = FormRef.Section(acHeader).Height
   
   Dim lngFooterHeight As Long
   lngFooterHeight = FormRef.Section(acFooter).Height

   If FormRef.InsideHeight < RecordCount * lngDetailHeight + lngHeaderHeight + lngFooterHeight Then
      If (FormRef.ScrollBars And 2) = 0 Then
         FormRef.ScrollBars = FormRef.ScrollBars + 2
      End If
   Else
      If (FormRef.ScrollBars And 2) = 2 Then
         FormRef.ScrollBars = FormRef.ScrollBars - 2
      End If
   End If
   
End Sub

Public Sub RemoveMultiValueSelection(ByVal FormRef As Form, ByVal MultiValueControl As Control)

   Dim rsMV As DAO.Recordset2

   If IsArray(MultiValueControl.Value) Then

      Set rsMV = FormRef.Recordset.Fields(MultiValueControl.ControlSource).Value
      With rsMV
         Do While Not .EOF
            .Delete
            .Requery
         Loop
      End With

      FormRef.Dirty = False
      MultiValueControl.Requery

   End If

End Sub

Public Sub CheckSubFormSourceObject(ByVal SubFormRef As SubForm, ByVal SourceObjectNameIfEmpty As String, Optional ByVal RequeryIfNoEmpty As Boolean = False)
   
   With SubFormRef
      If Len(.SourceObject) = 0 Then
         .SourceObject = SourceObjectNameIfEmpty
      ElseIf RequeryIfNoEmpty Then
         .Requery
      End If
   End With
   
End Sub

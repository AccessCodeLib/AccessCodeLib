VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: InfoMessageBox
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Text in Anwendung einblenden und optional automatisch ausblenden inkl. fadein/fadeout-Feature
' </summary>
' <remarks>
' <code>
' With New InfoMessageBox
'   .Show "Hallo Du!", 3000, 0, 3000, 6000
' End With
' </code>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/InfoMessageBox.cls</file>
'  <use>form/InfoMessageBoxForm.frm</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private WithEvents m_Form As Form
Attribute m_Form.VB_VarHelpID = -1
Private WithEvents m_Form_DetailSection As Section
Attribute m_Form_DetailSection.VB_VarHelpID = -1
Private Const EventProcedureTag As String = "[Event Procedure]"

#Const DEBUG_FADEINOUT_TIMER = 0
#Const DEBUG_FADEINOUT_TIMER_STEP = 0

Private m_Self As InfoMessageBox
Private m_FadeOutIsRunning As Boolean
Private m_FadeInIsRunning As Boolean
Private m_FadeOutTime As Long
Private m_FadeInTransparencyTarget As Long
Private m_CloseTimerInterval As Long
Private m_FadeInOutStep As Long

'********************************************************************************
' Deklarationen für Transparente Darstellung

Private m_Transparency As Byte

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'********************************************************************************
' Deklarationen für GetTextHeight

Private Const TWIPSPERINCH = 1440
Private Const LOGPIXELSY = 90
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_NOCLIP = &H100
Private Const CLIP_LH_ANGLES = 16
Private Const LF_FACESIZE = 32

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function apiCreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function apiGetDC Lib "user32" Alias "GetDC" (ByVal Hwnd As Long) As Long
Private Declare Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal Hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function apiDrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'********************************************************************************

Public Event InfoMessageClosed()

Public Sub Show(Optional ByVal NewInfoText As String = vbNullString, _
                Optional ByVal CloseTimer As Long = 0, _
                Optional ByVal TransparencyFactor As Long = -1, _
                Optional ByVal FadeOutTime As Long = 0, _
                Optional ByVal FadeInTime As Long = 0, _
                Optional ByVal KeepSelfReference As Boolean = True)
   
   
   If KeepSelfReference Then
      Set m_Self = Me ' .. eigene Instanz halten
   End If
   
   Me.Form.CheckWidth
   
   If Len(NewInfoText) > 0 Then
      Text = NewInfoText
   End If
   
   If FadeInTime > 0 Then
      Transparency = 100
   ElseIf TransparencyFactor >= 0 Then
      Transparency = TransparencyFactor
   End If

   m_FadeOutTime = FadeOutTime
   m_CloseTimerInterval = CloseTimer
   
   m_Form.Visible = True
   
   If FadeInTime > 0 Then
      StartFadeIn FadeInTime, TransparencyFactor
   Else
      m_Form.TimerInterval = m_CloseTimerInterval
   End If
   
End Sub

Public Property Let Text(ByVal NewText As String)

   NewText = GetCheckedHtmlRichText(NewText)
   m_Form.InfoText = NewText
   
   If Len(NewText) > 0 Then
      AlignText m_Form.InfoTextControl
   End If
   
End Property

Private Sub AlignText(ByVal ctl As Control)
   
   Dim TextHeight As Long
   TextHeight = GetTextHeight(ctl)
   
   If TextHeight > ctl.Height Then
      m_Form.SetInfoTextControlHeight TextHeight + 100
      ctl.TopMargin = 50
   Else
      ctl.TopMargin = ((ctl.Height - TextHeight) / 2)
   End If
   
End Sub

Private Function GetCheckedHtmlRichText(ByVal TextToCheck As String) As String
   
   Dim strText As String
   strText = TextToCheck
   
   If Val(Application.SysCmd(acSysCmdAccessVer)) >= 12 Then ' ab Ac2007
   If m_Form.InfoTextControl.Properties("TextFormat") = 1 Then ' 1 = acTextFormatHTMLRichText
      If InStr(1, strText, "<div>", vbTextCompare) = 0 Then ' normaler Text, der im RichTextformat dargestellt werden soll
         strText = "<div>" & Replace(strText, vbNewLine, "<br />") & "</div>"
      End If
   End If
   End If
   
   GetCheckedHtmlRichText = strText
   
End Function

Public Property Get Form() As Form
   CheckFormRef
   Set Form = m_Form
End Property

Public Property Set Form(ByVal NewRef As Form)
   
   Set m_Form = NewRef
   Set m_Form_DetailSection = m_Form.Section(acDetail)
   
   If Len(m_Form.OnTimer) = 0 Then
      m_Form.OnTimer = EventProcedureTag
   End If
   If Len(m_Form.OnClick) = 0 Then
      m_Form.OnClick = EventProcedureTag
   End If
   If Len(m_Form.OnClose) = 0 Then
      m_Form.OnClose = EventProcedureTag
   End If
   
   If Len(m_Form_DetailSection.OnClick) = 0 Then
      m_Form_DetailSection.OnClick = EventProcedureTag
   End If
   
End Property

Public Sub Move(Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Width As Variant, Optional ByVal Height As Variant)
   Me.Form.Move Left, Top, Width, Height
End Sub

Private Sub CheckFormRef()
   If m_Form Is Nothing Then
      Set Form = New Form_InfoMessageBoxForm
   End If
End Sub

Public Property Let Transparency(ByVal NewTransparency As Byte)
   
   Dim ret As Long
   Dim TransparencyValue As Long
   Dim bAlpha As Byte
   
   m_Transparency = NewTransparency
   If m_Transparency > 100 Then
      m_Transparency = 100
   ElseIf m_Transparency < 0 Then
      m_Transparency = 0
   End If
   
   TransparencyValue = CDbl(m_Transparency) * 2.55
   If TransparencyValue > 255 Then
      TransparencyValue = 255
   ElseIf TransparencyValue < 0 Then
      TransparencyValue = 0
   End If
   
   bAlpha = (255 - TransparencyValue)
   
   ret = GetWindowLong(Me.Form.Hwnd, (-20))
   ret = ret Or &H80000
   SetWindowLong m_Form.Hwnd, (-20), ret
                               
   SetLayeredWindowAttributes m_Form.Hwnd, 0, bAlpha, &H2
    
End Property

Public Property Get Transparency() As Byte
   Transparency = m_Transparency
End Property

Public Property Get BackColor() As Long
   BackColor = m_Form.BackColor
End Property

Public Property Let BackColor(ByVal NewValue As Long)
   Me.Form.BackColor = NewValue
End Property

Public Property Get ForeColor() As Long
   ForeColor = m_Form.ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As Long)
   Me.Form.ForeColor = NewValue
End Property

Private Sub m_Form_Close()
   RaiseEvent InfoMessageClosed
   Dispose
End Sub

Private Sub m_Form_DetailSection_Click()
   CloseForm
End Sub

Private Sub m_Form_Click()
   CloseForm
End Sub

Private Sub m_Form_Timer()
   
   If m_FadeOutIsRunning Then
      RunFadeOutStep
   ElseIf m_FadeInIsRunning Then
      RunFadeInStep
   Else
      m_Form.TimerInterval = 0
      If m_FadeOutTime > 0 Then
         StartFadeOut
      Else
#If DEBUG_FADEINOUT_TIMER Then
   Debug.Print ObjPtr(m_Form), "CloseForm", Timer, Now()
#End If
         CloseForm
      End If
   End If
   
End Sub

Private Sub StartFadeOut()

   Me.Form.TimerInterval = 0

   m_FadeOutIsRunning = True

   If Me.Transparency >= 100 Then
      m_FadeInOutStep = 1
      m_Form.TimerInterval = 1
      Exit Sub
   End If
   
   m_Form.TimerInterval = CalcFadeInFadeOutTimerInverval(m_FadeOutTime, (100 - Me.Transparency))

#If DEBUG_FADEINOUT_TIMER Then
   Debug.Print ObjPtr(Me), "StartFadeOut", Me.TimerInterval, Timer, Now()
#End If
      
End Sub

Private Function CalcFadeInFadeOutTimerInverval(ByVal FadeInOutTime As Long, ByVal TransparenceDiff As Long) As Long

   Dim lngTimerInveral As Long

   m_FadeInOutStep = Int(CDbl(TransparenceDiff) / (CDbl(FadeInOutTime) / 100#))
   If m_FadeInOutStep = 0 Then
      m_FadeInOutStep = 1
   End If
   
   lngTimerInveral = Int(CDbl(FadeInOutTime) / (CDbl(TransparenceDiff) / CDbl(m_FadeInOutStep)))
   If lngTimerInveral = 0 Then
      lngTimerInveral = 1
   End If
   
#If DEBUG_FADEINOUT_TIMER Then
   Debug.Print ObjPtr(Me), "CalcFadeInFadeOutTimerInverval", FadeInOutTime, TransparenceDiff, lngTimerInveral, m_FadeInOutStep
#End If
   
   CalcFadeInFadeOutTimerInverval = lngTimerInveral

End Function


Private Sub RunFadeOutStep()

   Dim NewTransparency As Long

   NewTransparency = Me.Transparency + m_FadeInOutStep
   If NewTransparency > 100 Then
      NewTransparency = 100
   End If
   
   Me.Transparency = NewTransparency

#If DEBUG_FADEINOUT_TIMER_STEP Then
   Debug.Print ObjPtr(Me), "RunFadeOutStep", m_FadeInOutStep, Timer, Now()
#End If
   
   If Me.Transparency >= 90 Then
#If DEBUG_FADEINOUT_TIMER Then
      Debug.Print ObjPtr(Me), "CloseFadeOutStep", Timer, Now()
#End If
      
      CloseForm
   End If

End Sub

Public Sub CloseForm()
   Dispose
End Sub

Private Sub Dispose()
   Set m_Form_DetailSection = Nothing
   Set m_Form = Nothing
   Set m_Self = Nothing
End Sub

Private Sub StartFadeIn(ByVal FadeInTime As Long, ByVal TargetTransparencyFactor As Long)
   
   If TargetTransparencyFactor >= 0 Then
      m_FadeInTransparencyTarget = TargetTransparencyFactor
   Else
      m_FadeInTransparencyTarget = 0
   End If
   
   m_FadeInIsRunning = True
   
   If m_FadeInTransparencyTarget >= 100 Then
      m_FadeInOutStep = 1
      m_Form.TimerInterval = 1
      Stop
      Exit Sub
   End If
   
   m_Form.TimerInterval = CalcFadeInFadeOutTimerInverval(FadeInTime, 100 - m_FadeInTransparencyTarget)

#If DEBUG_FADEINOUT_TIMER Then
   Debug.Print ObjPtr(Me), "StartFadeIn", m_Form.TimerInterval, Timer, Now()
#End If

End Sub

Private Sub RunFadeInStep()
   
   Dim NewTransparency As Long
   
   NewTransparency = Me.Transparency - m_FadeInOutStep
   
   If NewTransparency <= m_FadeInTransparencyTarget Then
      CloseFadeIn
      Exit Sub
   End If
   
   Me.Transparency = NewTransparency
   
#If DEBUG_FADEINOUT_TIMER_STEP Then
   Debug.Print ObjPtr(Me), "RunFadeInStep", m_FadeInOutStep, Timer, Now()
#End If
   
End Sub

Private Sub CloseFadeIn()
   
   m_Form.TimerInterval = 0
   m_FadeInIsRunning = False
   
#If DEBUG_FADEINOUT_TIMER Then
   Debug.Print ObjPtr(Me), "CloseFadeIn", Me.TimerInterval, Timer, Now()
#End If
      
   Me.Transparency = m_FadeInTransparencyTarget
   
   m_Form.TimerInterval = m_CloseTimerInterval
   
#If DEBUG_FADEINOUT_TIMER Then
   If m_CloseTimerInterval > 0 Then
      Debug.Print ObjPtr(Me), "StartCloseTimer", m_Form.TimerInterval, Timer, Now()
   End If
#End If
   
End Sub

Private Function GetTextHeight(ByVal ctl As Control) As Long
' = angepasster Code der Funktion fTextWidthOrHeight von Stephen Lebans
'
' Original-Code: http://www.lebans.com/verticaljustification.htm
'
' Änderungen:
'     * gekürzt, da nur Höhe benötigt wird
'     * Funktionalität für Berichte entfernt
'     * Text wird aus Steuerelement gelesen
'     * Nicht mehr benötigte Code-Blöcke entfernt
'
'
' Prozedur-Kopf von "fTextWidthOrHeight":
'*********************************************************
 'Name                   FUNCTION() fTextWidthOrHeight
 '
 'Purpose:               Returns the Height or Width needed to
 '                       display the contents of the Control passed
 '                       to this function. This function
 '                       uses the Control's font attributes to build
 '                       a Font for the required calculations.
 '
 '                       This function replaces the Report object's TextHeight
 '                       and TextWidth methods which only work for a single line of text.
 '                       This function works with multiple lines of text and
 '                       also with both Forms and Reports.
 '
 'Version:               4.1
 '
 'Calls:                 Text API stuff. DrawText performs the actual
 '                       calculation to determine Control Height.
 '
 'Returns:               Height or width of Control in TWIPS required
 '                       to display current contents.
 '
 'Created by:            Stephen Lebans
 '
 'Credits:               If you want some...take some.
 '
 'Date:                  May 22, 2001
 '
 'Time:                  10:10:10pm
 '
 'Feedback:              Stephen@lebans.com
 '
 'My Web Page:           www.lebans.com
 '
 'Copyright:             Lebans Holdings Ltd.
 '                       Please feel free to use this code
 '                       without restriction in any application you develop.
 '                       This code may not be resold by itself or as
 '                       part of a collection.
 '
 'What's Missing:        Let me know!
 '
 '
 '
 'Bugs:
 'None at this point.
 '
 'Enjoy
 'Stephen Lebans
 
'*********************************************************
 
   Dim strText As String
   
   ' Structure for DrawText calc
   Dim sRect As RECT
   
   ' Reports Device Context
   Dim hDC As Long
   
   ' Holds the current screen resolution
   Dim lngDPI As Long
   
   Dim newfont As Long
   ' Handle to our Font Object we created.
   ' We must destroy it before exiting main function
   
   Dim oldfont As Long
   ' Device Context's Font we must Select back into the DC
   ' before we exit this function.
 
   ' Temporary holder for returns from API calls
   Dim lngRet As Long
   
   ' Logfont struct
   Dim myfont As LOGFONT
   
   ' LineSpacing Amount
   Dim lngLineSpacing As Long
   
   ' Ttemp var
   Dim numLines As Long
   
   ' Temp string var for current printer name
   Dim strName As String
   
   ' Temp vars
   Dim sngTemp1 As Single
   Dim sngTemp2 As Single
 
   hDC = apiGetDC(0&)

   Select Case ctl.ControlType
       Case acTextBox
         strText = Nz(ctl.Value, vbNullString)
         If Val(Application.SysCmd(acSysCmdAccessVer)) >= 12 Then ' ab Ac2007
         If ctl.Properties("TextFormat") = 1 Then ' 1 = acTextFormatHTMLRichText
            strText = GetClearedTextFromHtmlRichTextBox(ctl)
         End If
         End If
       
       Case acLabel, acCommandButton
         strText = Nz(ctl.Caption, vbNullString)
       
       Case acListBox
         strText = Nz(ctl.ItemData(0), vbNullString)
       
       Case Else
         Err.Raise vbObjectError, "GetTextHeight", "ControlType not supported"
   End Select
   
   ' Get current device resolution
   lngDPI = apiGetDeviceCaps(hDC, LOGPIXELSY)

   ' We use a negative value to signify
   ' to the CreateFont function that we want a Glyph
   ' outline of this size not a bounding box.
   ' Copy font stuff from Text Control's property sheet
   With ctl
      myfont.lfClipPrecision = CLIP_LH_ANGLES
      myfont.lfOutPrecision = 0 ' sl APRIL 08/2002OUT_TT_ONLY_PRECIS
      myfont.lfEscapement = 0
      myfont.lfFaceName = .FontName & Chr$(0)
      myfont.lfWeight = .FontWeight
      myfont.lfItalic = .FontItalic
      myfont.lfUnderline = .FontUnderline
      'Must be a negative figure for height or system will return
      'closest match on character cell not glyph
      myfont.lfHeight = (.FontSize / 72) * -lngDPI
      ' Create our temp font
      newfont = apiCreateFontIndirect(myfont)
   End With
   
   If newfont = 0 Then
      Err.Raise vbObjectError + 256, "GetTextHeight", "Cannot Create Font"
   End If

   ' Select the new font into our DC.
   oldfont = apiSelectObject(hDC, newfont)
   ' Use DrawText to Calculate height of Rectangle required to hold
   ' the current contents of the Control passed to this function.
   With sRect
   
      .Left = 0
      .Top = 0
      .Bottom = 0
      .Right = (ctl.Width / (TWIPSPERINCH / lngDPI)) - 10
      
      ' Calculate our bounding box based on the controls current width
      lngRet = apiDrawText(hDC, strText, -1, sRect, DT_CALCRECT Or DT_TOP Or DT_LEFT Or DT_WORDBREAK Or DT_NOCLIP)

      ' Cleanup
      lngRet = apiSelectObject(hDC, oldfont)
      ' Delete the Font we created
      apiDeleteObject (newfont)
      
      lngRet = apiReleaseDC(0&, hDC)
      
      GetTextHeight = .Bottom * (TWIPSPERINCH / lngDPI)
      
   End With

End Function

Private Function GetClearedTextFromHtmlRichTextBox(ByVal tb As TextBox) As String
   
   Dim strTemp As String
   
   strTemp = Nz(tb.Value, vbNullString)

   If InStr(1, strTemp, "<div>", vbTextCompare) > 0 Then ' Text wird Access-Html-Richtext entsprechen => Zeilenumbruch = div-Tag
      'strTemp = Application.PlainText(strTemp) ' PlainText-Funktion ist erst ab 2007 verfügbar => CallByName, damit auch mit 2003 kompiliert werden kann
      strTemp = VBA.CallByName(Application, "PlainText", VbMethod, strTemp)
   Else 'use Replace - => Zeilenumbruch = vbNewline oder <br>-Tag
      strTemp = Replace(strTemp, "<br />", vbNewLine)
      strTemp = Replace(strTemp, "<br/>", vbNewLine)
      strTemp = Replace(strTemp, "<br>", vbNewLine)
   End If
   
   GetClearedTextFromHtmlRichTextBox = strTemp

End Function

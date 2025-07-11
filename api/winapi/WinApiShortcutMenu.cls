VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiShortcutMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.winapi.WinApiShortcutMenu
'---------------------------------------------------------------------------------------
'
' ShortcutMenu via Win API
'
' Author:
'     Josef Poetzl
'
' @todo Clean up code and create doc
'
' Remarks:
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' | Adapted variant from API example by J�rg Ostendorp of AEK10
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiShortcutMenu.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>api/winapi/WinApiLayoutTools.cls</use>
'  <use>api/winapi/WinApiWindowTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

#If VBA7 Then

Private m_Helper As LongPtr 'Auxiliary window handle

Private Type APIMSG
   Hwnd As LongPtr
   Message As LongPtr
   wParam As Long
   lParam As Long
   Time As LongPtr
   pt As POINTAPI
End Type

'Auxiliary window
Private Declare PtrSafe Function CreateWindowEx _
      Lib "user32.dll" _
      Alias "CreateWindowExA" ( _
      ByVal dwExStyle As Long, _
      ByVal lpClassname As String, _
      ByVal lpWindowName As String, _
      ByVal dwStyle As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hwndParent As LongPtr, _
      ByVal hMenu As LongPtr, _
      ByVal hInstance As LongPtr, _
      ByRef lpParam As Any) As Long

'Standard menu functions
Private Declare PtrSafe Function CreatePopupMenu _
      Lib "user32.dll" () As Long

Private Declare PtrSafe Function AppendMenu _
      Lib "user32.dll" _
      Alias "AppendMenuA" ( _
      ByVal hMenu As LongPtr, _
      ByVal wFlags As Long, _
      ByVal wIDNewItem As LongPtr, _
      ByVal lpNewItem As Any _
      ) As Long

Private Declare PtrSafe Function TrackPopupMenu _
      Lib "user32.dll" ( _
      ByVal hMenu As LongPtr, _
      ByVal wFlags As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal nReserved As Long, _
      ByVal Hwnd As LongPtr, _
      lprc As Any _
      ) As LongPtr 'lprc As RECT

Private Declare PtrSafe Function DestroyMenu _
      Lib "user32.dll" ( _
         ByVal hMenu As LongPtr _
      ) As Long

'Message Loop
Private Declare PtrSafe Function GetMessage _
      Lib "user32.dll" _
      Alias "GetMessageA" ( _
      ByRef lpMsg As APIMSG, _
      ByVal Hwnd As LongPtr, _
      ByVal wMsgFilterMin As Long, _
      ByVal wMsgFilterMax As Long _
      ) As Long

Private Declare PtrSafe Function DispatchMessage _
      Lib "user32.dll" _
      Alias "DispatchMessageA" ( _
      ByRef lpMsg As APIMSG _
      ) As Long

Private Declare PtrSafe Function TranslateMessage _
      Lib "user32.dll" (ByRef lpMsg As APIMSG) As Long

Private Declare PtrSafe Function GetWindowRect _
      Lib "user32.dll" ( _
         ByVal Hwnd As LongPtr, _
         ByRef lpRect As Rect _
      ) As LongPtr
      
Private Declare PtrSafe Function SetMenuDefaultItem _
      Lib "user32" ( _
         ByVal hMenu As LongPtr, _
         ByVal uItem As LongPtr, _
         ByVal fByPos As LongPtr _
      ) As LongPtr
      
#Else

Private m_Helper As Long 'Auxiliary window handle

Private Type APIMSG
   Hwnd As Long
   Message As Long
   wParam As Long
   lParam As Long
   Time As Long
   pt As POINTAPI
End Type

'Auxiliary window
Private Declare Function CreateWindowEx _
      Lib "user32.dll" _
      Alias "CreateWindowExA" ( _
      ByVal dwExStyle As Long, _
      ByVal lpClassname As String, _
      ByVal lpWindowName As String, _
      ByVal dwStyle As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hwndParent As Long, _
      ByVal hMenu As Long, _
      ByVal hInstance As Long, _
      ByRef lpParam As Any) As Long

'Standard menu functions
Private Declare Function CreatePopupMenu _
      Lib "user32.dll" () As Long

Private Declare Function AppendMenu _
      Lib "user32.dll" _
      Alias "AppendMenuA" ( _
      ByVal hMenu As Long, _
      ByVal wFlags As Long, _
      ByVal wIDNewItem As Long, _
      ByVal lpNewItem As Any _
      ) As Long

Private Declare Function TrackPopupMenu _
      Lib "user32.dll" ( _
      ByVal hMenu As Long, _
      ByVal wFlags As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal nReserved As Long, _
      ByVal Hwnd As Long, _
      lprc As Any _
      ) As Long 'lprc As RECT

Private Declare Function DestroyMenu _
      Lib "user32.dll" ( _
         ByVal hMenu As Long _
      ) As Long

'Message Loop
Private Declare Function GetMessage _
      Lib "user32.dll" _
      Alias "GetMessageA" ( _
      ByRef lpMsg As APIMSG, _
      ByVal Hwnd As Long, _
      ByVal wMsgFilterMin As Long, _
      ByVal wMsgFilterMax As Long _
      ) As Long

Private Declare Function DispatchMessage _
      Lib "user32.dll" _
      Alias "DispatchMessageA" ( _
      ByRef lpMsg As APIMSG _
      ) As Long

Private Declare Function TranslateMessage _
      Lib "user32.dll" (ByRef lpMsg As APIMSG) As Long

Private Declare Function GetWindowRect _
      Lib "user32.dll" ( _
         ByVal Hwnd As Long, _
         ByRef lpRect As RECT _
      ) As Long
      
Private Declare Function SetMenuDefaultItem _
      Lib "user32" ( _
         ByVal hMenu As Long, _
         ByVal uItem As Long, _
         ByVal fByPos As Long _
      ) As Long

#End If
      
'Message
Private Const WM_COMMAND As Long = &H111
Private Const WM_MOUSELEAVE As Long = &H2A3

'Menu or item styles

Public Enum MenuItemStyle
   MF_SEPARATOR = &H800&
   MF_STRING = &H0&
   MF_GRAYED = &H1&
   MF_BITMAP = &H4&
   MF_POPUP = &H10&
   MF_BYCOMMAND = &H0&
   MF_MENUBARBREAK = &H20&
   MF_SYSMENU = &H2000&
   MF_CHECKED = &H8&
End Enum

'Menu alignment
Private Const TPM_BOTTOMALIGN As Long = &H20&
'Private Const TPM_TOPALIGN As Long = &H0&

'Auxiliary window
Private Const WS_CHILD As Long = &H40000000

'---- additions:

Private Type WinAPIMenuHandlerItem
   ItemNumber As Long
   ItemText As String
   ItemType As Long
   SubMenu As Long
   DefaultItem As Boolean
End Type

Private m_MenuItems() As WinAPIMenuHandlerItem
Private m_ItemCnt As Long
Private m_AccessForm As Access.Form
Private m_MenuControl As Access.Control
Private m_Section As Long
Private m_WindowTools As WinApiWindowTools
Private m_LayoutTools As WinApiLayoutTools

Private Sub Class_Initialize()
On Error Resume Next
   m_ItemCnt = 0
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   Set m_MenuControl = Nothing
   Set m_AccessForm = Nothing
End Sub

Public Property Get Object() As WinApiShortcutMenu
   Set Object = Me
End Property

Private Property Get WindowTools() As WinApiWindowTools
   If m_WindowTools Is Nothing Then
      Set m_WindowTools = New WinApiWindowTools
   End If
   Set WindowTools = m_WindowTools
End Property

Private Property Get LayoutTools() As WinApiLayoutTools
   If m_LayoutTools Is Nothing Then
      Set m_LayoutTools = New WinApiLayoutTools
   End If
   Set LayoutTools = m_LayoutTools
End Property


Public Property Let ControlSection(ByVal SectionIndex As Long)
On Error Resume Next
   m_Section = SectionIndex
End Property

Public Property Get MenuControl() As Access.Control
   Set MenuControl = m_MenuControl
End Property

Public Property Set MenuControl(ByRef MnuCtl As Access.Control)
   
#If VBA7 Then
   Dim lngSectionHwnd As LongPtr
#Else
   Dim lngSectionHwnd As Long
#End If
   Dim ParentObj As Object

   Set m_MenuControl = MnuCtl
   
   Set ParentObj = m_MenuControl.Parent
   If TypeOf ParentObj Is Access.Page Then
      Set ParentObj = ParentObj.Parent.Parent
   End If
   If m_AccessForm Is Nothing Then
      Set m_AccessForm = ParentObj
   End If
   
   Select Case m_Section
      Case 0 'Detailbereich
         lngSectionHwnd = WindowTools.GetDetailSection(ParentObj.Hwnd)
      Case 1 'Kopf
         lngSectionHwnd = WindowTools.GetHeaderSection(ParentObj.Hwnd)
      Case 2 ' Fu�
         lngSectionHwnd = WindowTools.GetFooterSection(ParentObj.Hwnd)
      Case Else
         lngSectionHwnd = WindowTools.GetDetailSection(ParentObj.Hwnd)
   End Select
   Set ParentObj = Nothing
   
   'DoCmd.Restore '??? wozu ???
   'Hilfsfenster am Ursprung des Detailbereichs erstellen zur Positionsbestimmung
   'wenn das Formular gescrollt wird
   
   m_Helper = CreateWindowEx(0, _
                     "Static", _
                     "Helper", _
                     WS_CHILD, _
                     0, 0, 0, 0, _
                     lngSectionHwnd, _
                     0, 0, 0)
   'Wenn die Buttons nicht im Detailbereich sondern im Header oder Footer eingef�gt werden,
   'kann auf das Hilsfenter verzichtet werden. Da diese Bereiche nicht gescrollt werden
   'k�nnen, k�nnen Sie auch direkt zur Positionsbestimmung des Kontextmen�s herangezogen
   'werden
   ' => ??? wie ?

End Property

Public Property Get AccessForm() As Access.Form
   Set AccessForm = m_AccessForm
End Property

Public Property Set AccessForm(ByRef FrmRef As Access.Form)
   Set m_AccessForm = FrmRef
End Property

'---------------------------------------------------------------------------------------
' Sub: AddMenuItem
'---------------------------------------------------------------------------------------
'
' Create menu item
'
' Parameters:
'     MenuItemNumber - Number of the menu item
'     ItemText       - Menu item text
'     ItemType       - Menu item type
'     SubMenu        - Number of the assigned submenu
'
'---------------------------------------------------------------------------------------
Public Sub AddMenuItem(ByVal MenuItemNumber As Long, ByVal ItemText As String, _
              Optional ByVal ItemType As MenuItemStyle = MF_STRING, _
              Optional ByVal SubMenu As Long = 0, _
              Optional ByVal DefaultItem As Boolean = False)
              
   m_ItemCnt = m_ItemCnt + 1
   ReDim Preserve m_MenuItems(m_ItemCnt)
   
   With m_MenuItems(m_ItemCnt)
      .ItemNumber = MenuItemNumber
      .ItemText = ItemText
      .ItemType = ItemType
      .SubMenu = SubMenu
      .DefaultItem = DefaultItem
   End With

End Sub

'---------------------------------------------------------------------------------------
' Function: OpenMenu
'---------------------------------------------------------------------------------------
'
' �ffnet das Popup-Men�
'
' Parameters:
'     X - desired X position ... can be omitted if Control was specified
'     X - desired Y-position ... can be omitted if Control was specified
'
' Returns:
'     Long - Number of the selected menu item
'
'---------------------------------------------------------------------------------------
Public Function OpenMenu(Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As Long

#If VBA7 Then
   Dim MenuHwnd As LongPtr
   Dim TmpHwnd As LongPtr
   Dim Hwnd As LongPtr
#Else
   Dim MenuHwnd As Long
   Dim TmpHwnd As Long
   Dim Hwnd As Long
#End If

   Dim Message As APIMSG
   Dim RcHelper As Rect
   Dim ButtonLeft As Long
   Dim ButtonTop As Long
   Dim XPos As Long
   Dim YPos As Long
   
   Dim ParentObj As Object
   Dim AccFormRect As Rect

   'Otherwise, nothing is displayed when ENTER is pressed:
   DoEvents
   
   'Create empty menu
   MenuHwnd = CreatePopupMenu
   
   'Add entries
   Dim i As Long
   
   For i = 1 To m_ItemCnt
      With m_MenuItems(i)
         If .SubMenu > 0 Then
            If TmpHwnd = 0 Then
               TmpHwnd = CreatePopupMenu
            End If
            If .ItemType And MF_POPUP = MF_POPUP Then
               AppendMenu MenuHwnd, MF_STRING Or MF_POPUP, TmpHwnd, .ItemText
               TmpHwnd = 0
            Else
               AppendMenu TmpHwnd, .ItemType, .ItemNumber, .ItemText
            End If
         Else
            AppendMenu MenuHwnd, .ItemType, .ItemNumber, .ItemText
         End If
      
         If .DefaultItem Then SetMenuDefaultItem MenuHwnd, .ItemNumber - 1, &H400&
      
      End With
      
   Next i

   'Determine position
   '- Auxiliary window (~section window) absolute
   GetWindowRect m_Helper, RcHelper
   
   If m_MenuControl Is Nothing Then
      GetWindowRect m_AccessForm.Hwnd, AccFormRect
      ButtonLeft = LayoutTools.GetPixelFromTwips(X) + AccFormRect.Left
      ButtonTop = LayoutTools.GetPixelFromTwips(Y) + AccFormRect.Top
   Else
   '- Buttom relative to the section window
      ButtonLeft = LayoutTools.GetPixelFromTwips(m_MenuControl.Left) + LayoutTools.GetPixelFromTwips(X)
      ButtonTop = LayoutTools.GetPixelFromTwips(m_MenuControl.Top) + LayoutTools.GetPixelFromTwips(Y)
   End If
   '- calc with each other
   XPos = RcHelper.Left + ButtonLeft - 1
   YPos = RcHelper.Top + ButtonTop
   
   'Show menu
   
   If m_MenuControl Is Nothing Then
      Set ParentObj = m_AccessForm
   Else
      Set ParentObj = m_MenuControl.Parent
      If TypeOf ParentObj Is Access.Page Then
         Set ParentObj = ParentObj.Parent.Parent
      End If
   End If
   
   TrackPopupMenu MenuHwnd, TPM_BOTTOMALIGN, XPos, YPos, _
               ByVal 0&, ParentObj.Hwnd, ByVal 0&
   
   'Message-Loop (Attention: for normal forms Application.hWndAccessApp, for POPUPS Form.Hwnd)
   
   If m_AccessForm.PopUp = True Then
      Hwnd = m_AccessForm.Hwnd
   Else
      Hwnd = Application.hWndAccessApp
   End If
   
   Dim RepeatGetMessage As Boolean
   Do ' Security sleep in case there are problems with windows message. (WM_MOUSELEAVE]
      RepeatGetMessage = False
      GetMessage Message, Hwnd, ByVal 0&, ByVal 0&
      TranslateMessage Message
      DispatchMessage Message
      
      'Evaluate message
      If Message.Message = WM_COMMAND Then
         OpenMenu = Message.wParam
      ElseIf Message.Message = WM_MOUSELEAVE Then
         RepeatGetMessage = True
      Else
         OpenMenu = 0
      End If
   Loop While RepeatGetMessage
   
   'Destroy menu
   DestroyMenu MenuHwnd

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiShortcutMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ShortcutMenu über Win-API"
'---------------------------------------------------------------------------------------
' Class Module: WinApiShortcutMenu
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ShortcutMenu über Win-API
' </summary>
' <remarks>
' </remarks>
'\ingroup WinAPI
'
' @todo Code aufräumen und Doxygen-Doku erstellen
'
'**/
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

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Angepasste Variante aus API-Beispiel von Jörg Ostendorp der AEK10
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type APIMSG
   Hwnd As Long
   Message As Long
   wParam As Long
   lParam As Long
   time As Long
   pt As POINTAPI
End Type

#If VBA7 Then

'Hilfsfenster
Private Declare PtrSafe Function CreateWindowEx _
      Lib "USER32.DLL" _
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

'Menü-Standardfunktionen
Private Declare PtrSafe Function CreatePopupMenu _
      Lib "USER32.DLL" () As Long

Private Declare PtrSafe Function AppendMenu _
      Lib "USER32.DLL" _
      Alias "AppendMenuA" ( _
      ByVal hMenu As Long, _
      ByVal wFlags As Long, _
      ByVal wIDNewItem As Long, _
      ByVal lpNewItem As Any _
      ) As Long

Private Declare PtrSafe Function TrackPopupMenu _
      Lib "USER32.DLL" ( _
      ByVal hMenu As Long, _
      ByVal wFlags As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal nReserved As Long, _
      ByVal Hwnd As Long, _
      lprc As Any _
      ) As Long 'lprc As RECT

Private Declare PtrSafe Function DestroyMenu _
      Lib "USER32.DLL" ( _
         ByVal hMenu As Long _
      ) As Long

'Message-Loop
Private Declare PtrSafe Function GetMessage _
      Lib "USER32.DLL" _
      Alias "GetMessageA" ( _
      ByRef lpMsg As APIMSG, _
      ByVal Hwnd As Long, _
      ByVal wMsgFilterMin As Long, _
      ByVal wMsgFilterMax As Long _
      ) As Long

Private Declare PtrSafe Function DispatchMessage _
      Lib "USER32.DLL" _
      Alias "DispatchMessageA" ( _
      ByRef lpMsg As APIMSG _
      ) As Long

Private Declare PtrSafe Function TranslateMessage _
      Lib "USER32.DLL" (ByRef lpMsg As APIMSG) As Long

Private Declare PtrSafe Function GetWindowRect _
      Lib "USER32.DLL" ( _
         ByVal Hwnd As Long, _
         ByRef lpRect As RECT _
      ) As Long
      
Private Declare PtrSafe Function SetMenuDefaultItem _
      Lib "user32" ( _
         ByVal hMenu As Long, _
         ByVal uItem As Long, _
         ByVal fByPos As Long _
      ) As Long
      
#Else

'Hilfsfenster
Private Declare Function CreateWindowEx _
      Lib "USER32.DLL" _
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

'Menü-Standardfunktionen
Private Declare Function CreatePopupMenu _
      Lib "USER32.DLL" () As Long

Private Declare Function AppendMenu _
      Lib "USER32.DLL" _
      Alias "AppendMenuA" ( _
      ByVal hMenu As Long, _
      ByVal wFlags As Long, _
      ByVal wIDNewItem As Long, _
      ByVal lpNewItem As Any _
      ) As Long

Private Declare Function TrackPopupMenu _
      Lib "USER32.DLL" ( _
      ByVal hMenu As Long, _
      ByVal wFlags As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal nReserved As Long, _
      ByVal Hwnd As Long, _
      lprc As Any _
      ) As Long 'lprc As RECT

Private Declare Function DestroyMenu _
      Lib "USER32.DLL" ( _
         ByVal hMenu As Long _
      ) As Long

'Message-Loop
Private Declare Function GetMessage _
      Lib "USER32.DLL" _
      Alias "GetMessageA" ( _
      ByRef lpMsg As APIMSG, _
      ByVal Hwnd As Long, _
      ByVal wMsgFilterMin As Long, _
      ByVal wMsgFilterMax As Long _
      ) As Long

Private Declare Function DispatchMessage _
      Lib "USER32.DLL" _
      Alias "DispatchMessageA" ( _
      ByRef lpMsg As APIMSG _
      ) As Long

Private Declare Function TranslateMessage _
      Lib "USER32.DLL" (ByRef lpMsg As APIMSG) As Long

Private Declare Function GetWindowRect _
      Lib "USER32.DLL" ( _
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

'Menü- bzw. Item-Styles

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

'Menü-Ausrichtung
Private Const TPM_BOTTOMALIGN As Long = &H20&
'Private Const TPM_TOPALIGN As Long = &H0&

'Hilfsfenster
Private Const WS_CHILD As Long = &H40000000

Private Helper As Long 'Hilfsfenster-Handle

'---- Ergänzung

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

'Public Event Selected(ByVal lMnuItemNr As Long)

Private Sub Class_Initialize()
On Error Resume Next
   m_ItemCnt = 0
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   Set m_MenuControl = Nothing
   Set m_AccessForm = Nothing
End Sub

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


Public Property Let ControlSection(ByVal lSection As Long)
On Error Resume Next
   m_Section = lSection
End Property

Public Property Get MenuControl() As Access.Control
   Set MenuControl = m_MenuControl
End Property

Public Property Set MenuControl(ByRef MnuCtl As Access.Control)
   
   Dim lngSectionHwnd As Long
   Dim objParent As Object

   Set m_MenuControl = MnuCtl
   
   Set objParent = m_MenuControl.Parent
   If TypeOf objParent Is Access.Page Then
      Set objParent = objParent.Parent.Parent
   End If
   If m_AccessForm Is Nothing Then
      Set m_AccessForm = objParent
   End If
   
   Select Case m_Section
      Case 0 'Detailbereich
         lngSectionHwnd = WindowTools.GetDetailSection(objParent.Hwnd)
      Case 1 'Kopf
         lngSectionHwnd = WindowTools.GetHeaderSection(objParent.Hwnd)
      Case 2 ' Fuß
         lngSectionHwnd = WindowTools.GetFooterSection(objParent.Hwnd)
      Case Else
         lngSectionHwnd = WindowTools.GetDetailSection(objParent.Hwnd)
   End Select
   Set objParent = Nothing
   
   'DoCmd.Restore '??? wozu ???
   'Hilfsfenster am Ursprung des Detailbereichs erstellen zur Positionsbestimmung
   'wenn das Formular gescrollt wird
   
   Helper = CreateWindowEx(0, _
                     "Static", _
                     "Helper", _
                     WS_CHILD, _
                     0, 0, 0, 0, _
                     lngSectionHwnd, _
                     0, 0, 0)
   'Wenn die Buttons nicht im Detailbereich sondern im Header oder Footer eingefügt werden,
   'kann auf das Hilsfenter verzichtet werden. Da diese Bereiche nicht gescrollt werden
   'können, können Sie auch direkt zur Positionsbestimmung des Kontextmenüs herangezogen
   'werden
   ' => ??? wie ?

End Property

Public Property Get AccessForm() As Access.Form
   Set AccessForm = m_AccessForm
End Property

Public Property Set AccessForm(ByRef frm As Access.Form)
   Set m_AccessForm = frm
End Property

'---------------------------------------------------------------------------------------
' Sub: AddMenuItem
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Menü-Eintrag erstellen
' </summary>
' <param name="lMenuItemNumber">Nummer des Menüeintrags</param>
' <param name="sItemText">Text des Menüeintrags</param>
' <param name="lItemType">Art des Menüeintrags</param>
' <param name="lSubMenu">Nummer des zugeordneten Untermenüs</param>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub AddMenuItem(ByVal lMenuItemNumber As Long, ByVal sItemText As String, _
              Optional ByVal lItemType As MenuItemStyle = MF_STRING, _
              Optional ByVal lSubMenu As Long = 0, _
              Optional ByVal bDefaultItem As Boolean = False)
              
   m_ItemCnt = m_ItemCnt + 1
   ReDim Preserve m_MenuItems(m_ItemCnt)
   
   With m_MenuItems(m_ItemCnt)
      .ItemNumber = lMenuItemNumber
      .ItemText = sItemText
      .ItemType = lItemType
      .SubMenu = lSubMenu
      .DefaultItem = bDefaultItem
   End With

End Sub

'---------------------------------------------------------------------------------------
' Function: OpenMenu
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Öffnet das Popup-Menü
' </summary>
' <param name="X">gewünschte X-Positon .. kann entfallen, wenn Control angegeben wurde</param>
' <param name="X">gewünschte Y-Positon .. kann entfallen, wenn Control angegeben wurde</param>
' <returns>Nummer des ausgewählten Menüeintrags</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenMenu(Optional ByVal X As Single = 0, Optional ByVal Y As Single = 0) As Long

   Dim MenuHwnd As Long
   Dim tmpHwnd As Long
   
   Dim rcHelper As RECT
   Dim ButtonLeft As Long
   Dim ButtonTop As Long
   Dim Xpos As Long
   Dim Ypos As Long
   Dim Message As APIMSG
 
   Dim objParent As Object
   Dim rcAccForm As RECT

   'Sonst wird bei ENTER nix angezeigt:
   DoEvents
   
   'Leeres Menü erstellen
   MenuHwnd = CreatePopupMenu
   
   'Einträge hinzufügen
   Dim i As Long
   
   For i = 1 To m_ItemCnt
      With m_MenuItems(i)
         If .SubMenu > 0 Then
            If tmpHwnd = 0 Then
               tmpHwnd = CreatePopupMenu
            End If
            If .ItemType And MF_POPUP = MF_POPUP Then
               AppendMenu MenuHwnd, MF_STRING Or MF_POPUP, tmpHwnd, .ItemText
               tmpHwnd = 0
            Else
               AppendMenu tmpHwnd, .ItemType, .ItemNumber, .ItemText
            End If
         Else
            AppendMenu MenuHwnd, .ItemType, .ItemNumber, .ItemText
         End If
      
         If .DefaultItem Then SetMenuDefaultItem MenuHwnd, .ItemNumber - 1, &H400&
      
      End With
      
   Next i

   'Position ermitteln
   '- Hilfsfenster (~Sektionsfenster) absolut
   GetWindowRect Helper, rcHelper
   
   If m_MenuControl Is Nothing Then
      GetWindowRect m_AccessForm.Hwnd, rcAccForm
      ButtonLeft = LayoutTools.GetPixelFromTwips(X) + rcAccForm.Left
      ButtonTop = LayoutTools.GetPixelFromTwips(Y) + rcAccForm.Top
   Else
   '- Buttom relativ zum Sektionsfenster
      ButtonLeft = LayoutTools.GetPixelFromTwips(m_MenuControl.Left) + LayoutTools.GetPixelFromTwips(X)
      ButtonTop = LayoutTools.GetPixelFromTwips(m_MenuControl.Top) + LayoutTools.GetPixelFromTwips(Y)
   End If
   '- Miteinander verrechnen
   Xpos = rcHelper.Left + ButtonLeft - 1
   Ypos = rcHelper.Top + ButtonTop
   
   'Menü anzeigen
   
   If m_MenuControl Is Nothing Then
      Set objParent = m_AccessForm
   Else
      Set objParent = m_MenuControl.Parent
      If TypeOf objParent Is Access.Page Then
         Set objParent = objParent.Parent.Parent
      End If
   End If
   
   TrackPopupMenu MenuHwnd, TPM_BOTTOMALIGN, Xpos, Ypos, _
               ByVal 0&, objParent.Hwnd, ByVal 0&
   
   'Message-Loop (Obacht: für normale Formulare Application.hWndAccessApp, für POPUPS Me.hwnd)
   
   Dim Hwnd As Long
   If m_AccessForm.PopUp = True Then
      Hwnd = m_AccessForm.Hwnd
   Else
      Hwnd = Application.hWndAccessApp
   End If
   
   Dim repeatGetMessage As Boolean
   Do 'Sicherheitsschliefe, falls es Probleme mit Windows-Message gibt. (WM_MOUSELEAVE)
      repeatGetMessage = False
      GetMessage Message, Hwnd, ByVal 0&, ByVal 0&
      TranslateMessage Message
      DispatchMessage Message
      
      'Message auswerten
      If Message.Message = WM_COMMAND Then
         OpenMenu = Message.wParam
      ElseIf Message.Message = WM_MOUSELEAVE Then
         repeatGetMessage = True
      Else
         OpenMenu = 0
      End If
   Loop While repeatGetMessage
   
   'Menü zerstören
   DestroyMenu MenuHwnd

End Function

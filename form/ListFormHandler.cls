VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListFormHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Klassenmodul: ListFormHandler (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formularsteuerungsklasse für Listen/Endlosformulare
' </summary>
' <remarks>
' Vereinfacht die Steuerung von Access-Formularen, die als Endlosformular Listen darstellen.
'
' DRAFT! - Code muss noch mit ACLib-Modulen kombiniert werden
'
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/ListFormHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/FormHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'/**
' <summary>
' Spezielle Auswahl von möglichen Formularsteuerelmenten mit besonderer Behandlung
' </summary>
' <remarks>1-99 sind identisch mit <see cref="FormHandler::FH_HandledControls">FH_HandledControls</see> aus FormHandler</remarks>
' <list type="table">
'   <item><term>FC_CloseCommandButton (1)</term><description>Schließen-Schaltfläche</description></item>
'   <item><term>FC_SelectionFirstControl (2)</term><description>Steuerelement, das beim deaktiviert/unsichtbar Setzen eines aktiven Steuerelements den Fokus erhalten soll</description></item>
'   <item><term>FC_EditModeCommandButton (3)</term><description>Schaltfläche für Bearbeitungsmoduswechsel (Readonly, Edit)</description></item>
'   <item><term>FC_EditModeInfoLabel (4)</term><description>Label für Anzeige des aktuellen Bearbeitungsmodues</description></item>
'   <item><term>FC_RequeryCommandButton (5)</term><description>Requery-Schaltfläche</description></item>
'   <item><term>FC_SelectCommandButton (101)</term><description>Auswahl-Schaltfläche</description></item>
' </list>
'**/
Public Enum LF_HandledControls
   LFC_CloseCommandButton = 1
   LFC_SelectionFirstControl = 2
   LFC_EditModeCommandButton = 3
   LFC_EditModeInfoLabel = 4
   LFC_RequeryCommandButton = 5
   LFC_SelectCommandButton = 101
   LFC_MultiSelectionCommitCommandButton = 111
   LFC_MultiSelectionMarkAllCheckBox = 112
End Enum

'/**
' <summary>
' Einstellmöglichkeiten für Auswahl (inkl. Darstellung
' </summary>
' <list type="table">
'   <item><term>LM_Normal (0)</term><description>Normales Endlosformular / Wechsel zu Detailformular o.ä. bei Betätigung der Auswahlschaltfläche</description></item>
'   <item><term>LM_Select (1)</term><description>Endlosformular für Auswahl eines Datensatzes bzw. mehrerer Datensätze</description></item>
' </list>
'**/
Public Enum LF_ListMode
   LM_Normal = 0
   LM_Select = 1
End Enum

'Allg. FormHandler
Private WithEvents m_FormHandler As FormHandler
Attribute m_FormHandler.VB_VarHelpID = -1

'ListMode
Private m_ListMode As LF_ListMode

' Formular-Objekt
Private WithEvents m_ListForm As Access.Form
Attribute m_ListForm.VB_VarHelpID = -1

' Detailbereich
Private WithEvents m_fsecDetail As Access.Section
Attribute m_fsecDetail.VB_VarHelpID = -1

' Button Select
Private WithEvents m_SelectCommandButton As CommandButton
Attribute m_SelectCommandButton.VB_VarHelpID = -1

Private m_lngSelectShiftMode As Long

'Form-Handler-Events

'/**
' <summary>
' Ereignis für Meldungen, die im Frontend behandelt werden sollen.
' </summary>
' <param name="Msg">Nachrichtentext</param>
' <param name="MsgNumber">Nachrichtennummer</param>
' <remarks></remarks>
'**/
Public Event FormHandlerMessage(ByVal Msg As String, ByVal MsgNumber As Long)

'/**
' <summary>
' Ereignis für Fehlermeldungen, die im Frontend behandelt werden sollen.
' </summary>
' <param name="ErrMsg">Fehlermeldung</param>
' <param name="ErrNumber">Fehlernummer</param>
' <remarks></remarks>
'**/
Public Event FormHandlerError(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef Response As Integer)

'/**
' <summary>
' Ereignis für Fehlermeldungen, wenn Referenz des Access-Forms auf Nothing steht und benötigt wird.
' </summary>
' <param name="ErrMsg">Fehlermeldung</param>
' <param name="ErrNumber">Fehlernummer</param>
' <param name="NewForm">Damit kann vom Access-FE erneut das Formobjekt übergeben werden.</param>
' <remarks></remarks>
'**/
Public Event FormHandlerErrorMissingForm(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewForm As Object)

'/**
' <summary>
' Ereignis für Fehlermeldungen, wenn Referenz zu <see cref="DbConnection">DbConnection</see> auf Nothing steht und benötigt wird.
' </summary>
' <param name="ErrMsg">Fehlermeldung</param>
' <param name="ErrNumber">Fehlernummer</param>
' <param name="NewCurrentDbConnection">Damit kann vom Access-Frontend erneut das <see cref="DbConnection">DbConnection-Objekt</see> übergeben werden.</param>
' <remarks></remarks>
'**/
Public Event FormHandlerErrorMissingDbConnection(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewCurrentDbConnection As Object)

'/**
' <summary>
' Ereignis-Meldung für Schließen des Access-Forms
' </summary>
' <param name="Cancel">ermöglicht Abbruch der Schließen-Aktion</param>
' <remarks></remarks>
'**/
Public Event CloseForm(ByRef Cancel As Boolean)

'/**
' <summary>
' Ereignis-Meldung wenn Access-Form-Unload-Ereignis eintritt
' </summary>
' <param name="Cancel">ermöglicht Abbruch der Unload-Aktion</param>
' <remarks>Wird hauptsächlich für sauberes Entladen der Objekt-Referenzen benötigt</remarks>
'**/
Public Event AccessFormUnload(ByRef Cancel As Integer)

'/**
' <summary>
' Ereignis-Meldung wenn EditMode geändert wurde
' </summary>
' <param name="NewMode">Neu gesetzter Modus</param>
' <remarks></remarks>
'**/
Public Event EditModeChanged(ByVal NewMode As Long)

'/**
' <summary>
' Ereignis bei Auswahl einer Datensatzes
' </summary>
' <param name="ItemID">Ausgewählte Datensatz-Kennung</param>
' <param name="ShiftMode">Kennung für Zusatztaste (Umschalt und/oder Strg)</param>
' <param name="OpenArg">Optionale Paramter fürs Öffnen</param>
' <remarks></remarks>
'**/
Public Event Selected(ByVal ItemID As Variant, ByVal ShiftMode As Long, ByVal OpenArg As Long)

'/**
' <summary>
' Ereignis-Meldung wenn Objektreferenzen entfernt wurden
' </summary>
' <remarks></remarks>
'**/
Public Event Disposed()

Private m_Disposing As Boolean
Private m_Disposed As Boolean

Private m_SelectionReturnDataSourceFieldName As String

Private m_MouseWheelFix As FormMouseWheelFix

'Init/Terminate
Private Sub Class_Initialize()
   Set m_FormHandler = New FormHandler
End Sub

Private Sub Class_Terminate()
   If Not m_Disposed Then Dispose
End Sub

Public Property Get Self() As ListFormHandler
   Set Self = Me
End Property

Public Sub Dispose()

   If m_Disposed Then Exit Sub

On Error Resume Next

   m_Disposing = True

   ClearFormObjectReferences
   ClearObjectReferences

   m_Disposed = True
   m_Disposing = False
   RaiseEvent Disposed
   
End Sub

Private Sub ClearObjectReferences()

On Error Resume Next

'FormHandler
   If Not (m_FormHandler Is Nothing) Then
      m_FormHandler.Dispose
   End If
   Set m_FormHandler = Nothing
   
End Sub

Private Sub ClearFormObjectReferences()
   
On Error Resume Next
   
   ' Button Select
   Set m_SelectCommandButton = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Property: FormHandler (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt Referenz auf <see cref="FormHandler">FormHandler-Basisklasse</see> zurück.
' </summary>
' <returns><see cref="FormHandler">FormHandler-Basisklasse</see></returns>
' <remarks>Ist nur für zusätzliche Code-Erweiterung implementiert.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get BaseFormHandler() As FormHandler
Attribute BaseFormHandler.VB_MemberFlags = "40"
'Attribute BaseFormHandler.VB_MemberFlags = "40"

   If m_FormHandler Is Nothing Then
      Set m_FormHandler = New FormHandler
      If m_Disposed Then m_Disposed = False
   ElseIf m_FormHandler.ActiveForm Is Nothing Then
      If Not (m_SelectCommandButton Is Nothing) Then
         On Error Resume Next
         Set m_FormHandler.ActiveForm = m_SelectCommandButton.Parent
         If m_Disposed Then m_Disposed = False
         On Error GoTo 0
      End If
   End If
   Set BaseFormHandler = m_FormHandler

End Property


'---------------------------------------------------------------------------------------
' Sub: CloseForm (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Access-Formular schließen
' </summary>
' <remarks>siehe <see cref="FormHandler::CloseForm">CloseForm</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub CloseForm()
   BaseFormHandler.CloseForm
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DefaultOrderParams
' Date     : 15.09.2007
'---------------------------------------------------------------------------------------
'/**
' <summary>
'
' </summary>
' <param name="sOrderParams">Parameter für Standard-Sortierung</param>
' <remarks>siehe <see cref="FormHandler::DefaultOrderParams">DefaultOrderParams</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let DefaultOrderParams(sOrderParams As String)
   BaseFormHandler.DefaultOrderParams = sOrderParams
End Property

'---------------------------------------------------------------------------------------
' Property: ActiveDbConnection (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ActiveDbConnection
' </summary>
' <returns>DbConnection</returns>
' <remarks>siehe <see cref="FormHandler::ActiveDbConnection">ActiveDbConnection</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ActiveDbConnection() As DbConnectionHandler
   Set ActiveDbConnection = BaseFormHandler.ActiveDbConnection
End Property

'---------------------------------------------------------------------------------------
' Property: ActiveDbConnection (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zuweisung einer neuen <see cref="DbConnection">DbConnection</see>
' </summary>
' <param name="NewDbConnection">NewDbConnection</param>
' <remarks>siehe <see cref="FormHandler::ActiveDbConnection">ActiveDbConnection</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set ActiveDbConnection(ByRef NewDbConnection As DbConnectionHandler)
   Set BaseFormHandler.ActiveDbConnection = NewDbConnection
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceType (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DataSourceType
' </summary>
' <param name="NewDataSourceType">FH_DataSourceType zur Einstellung der Datenquelle</param>
' <remarks>siehe <see cref="FormHandler::DataSourceType">DataSourceType</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let DataSourceType(NewDataSourceType As FH_DataSourceType)
   BaseFormHandler.DataSourceType = NewDataSourceType
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceType (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DataSourceType
' </summary>
' <returns>FH_DataSourceType</returns>
' <remarks>siehe <see cref="FormHandler::DataSourceType">DataSourceType</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DataSourceType() As FH_DataSourceType
   DataSourceType = BaseFormHandler.DataSourceType
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceUniqueFieldName (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Feldbezeichnung der Datenherkunft, die eindeutige Datensatzkennung zurückgibt
' </summary>
' <param name="FieldName">Feldbezeichnung der Datenherkunft</param>
' <remarks>siehe <see cref="FormHandler::DataSourceUniqueFieldName">DataSourceUniqueFieldName</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let DataSourceUniqueFieldName(ByVal FieldName As String)
   BaseFormHandler.DataSourceUniqueFieldName = FieldName
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceUniqueFieldName (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Feldbezeichnung der Datenherkunft, die eindeutige Datensatzkennung zurückgibt
' </summary>
' <returns>String</returns>
' <remarks>siehe <see cref="FormHandler::DataSourceUniqueFieldName">DataSourceUniqueFieldName</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DataSourceUniqueFieldName() As String
   DataSourceUniqueFieldName = BaseFormHandler.DataSourceUniqueFieldName
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceUniqueFieldName (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Feldbezeichnung der Datenherkunft, die eindeutige Datensatzkennung zurückgibt
' </summary>
' <param name="FieldName">Feldbezeichnung der Datenherkunft</param>
' <remarks>siehe <see cref="FormHandler::DataSourceUniqueFieldName">DataSourceUniqueFieldName</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let SelectionReturnDataSourceFieldName(ByVal FieldName As String)
   m_SelectionReturnDataSourceFieldName = FieldName
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceUniqueFieldName (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Feldbezeichnung der Datenherkunft, die eindeutige Datensatzkennung zurückgibt
' </summary>
' <returns>String</returns>
' <remarks>siehe <see cref="FormHandler::DataSourceUniqueFieldName">DataSourceUniqueFieldName</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get SelectionReturnDataSourceFieldName() As String

   If Len(m_SelectionReturnDataSourceFieldName) = 0 Then
      m_SelectionReturnDataSourceFieldName = DataSourceUniqueFieldName
   End If
   SelectionReturnDataSourceFieldName = m_SelectionReturnDataSourceFieldName

End Property

'---------------------------------------------------------------------------------------
' Property: ActiveForm (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' verwendetes Access-Form-Objekt
' </summary>
' <returns>Access.Form</returns>
' <remarks>siehe <see cref="FormHandler::ActiveForm">ActiveForm</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ActiveForm() As Access.Form

   Set ActiveForm = BaseFormHandler.ActiveForm

End Property

'---------------------------------------------------------------------------------------
' Property: ActiveForm (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Access-Form-Objekt welches ListFormHandler steuern soll
' </summary>
' <param name="FormRef">zu übergebendes Access-Form</param>
' <remarks>siehe <see cref="FormHandler::ActiveForm">ActiveForm</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set ActiveForm(ByRef FormRef As Access.Form)

   Set BaseFormHandler.ActiveForm = FormRef
   InitListForm FormRef
   
End Property

Private Sub InitListForm(ByRef NewForm As Access.Form)

   Set m_ListForm = NewForm
   If Not (m_ListForm Is Nothing) Then
      Set m_fsecDetail = m_ListForm.Section(0)
   End If
   
   With New FormMouseWheelFix
      If .Init(m_ListForm) Then
         Set m_MouseWheelFix = .Self
      End If
   End With
   
End Sub

'---------------------------------------------------------------------------------------
' Property: HandledControls (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Steuerelement(e) mit besonderer Behandlung
' </summary>
' <param name="Index">LF_HandledControls-Kennung für auszuwählendes Steuerelement</param>
' <returns>Access.Control</returns>
' <remarks>siehe <see cref="FormHandler::HandledControls">HandledControls</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get HandledControls(ByVal Index As LF_HandledControls) As Access.Control

   Select Case Index
      Case LF_HandledControls.LFC_SelectCommandButton
         Set HandledControls = selectCommandButton
      Case Else
         Set HandledControls = BaseFormHandler.HandledControls(Index)
   End Select

End Property

'---------------------------------------------------------------------------------------
' Property: HandledControls (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Steuerelement(e) mit besonderer Behandlung
' </summary>
' <param name="Index">LF_HandledControls-Kennung für auszuwählendes Steuerelement</param>
' <param name="FormControl">Steuerelement das der LF_HandledControls-Kennung zugeordnet werden soll</param>
' <remarks>siehe <see cref="FormHandler::HandledControls">HandledControls</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set HandledControls(ByVal Index As LF_HandledControls, ByRef FormControl As Access.Control)

   Select Case Index
      Case LF_HandledControls.LFC_SelectCommandButton
         Set selectCommandButton = FormControl
      Case Else
         Set BaseFormHandler.HandledControls(Index) = FormControl
   End Select

End Property

'---------------------------------------------------------------------------------------
' Sub: InitDataSource (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datenquelle einstellen
' </summary>
' <param name="lDataSourceType">Art der Datenquelle</param>
' <param name="sDataSource">Definition der Datenquelle (z.B.: SQL-Anweisung)</param>
' <param name="sDataSourceNoData">Definition der Datenquelle für leere Liste</param>
' <param name="sOptionalArg">Optionaler Parameterwert</param>
' <remarks>siehe <see cref="FormHandler::InitDataSource">InitDataSource</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub InitDataSource(ByVal lDataSourceType As FH_DataSourceType, _
             Optional ByVal sDataSource As String = vbNullString, _
             Optional ByVal sDataSourceNoData As String = vbNullString, _
             Optional ByVal sOptionalArg As String = vbNullString, _
             Optional ByVal sOptionalBindingDataSourceName As String = vbNullString)
             
   BaseFormHandler.InitDataSource lDataSourceType, sDataSource, sDataSourceNoData, sOptionalArg, sOptionalBindingDataSourceName
   
End Sub

'---------------------------------------------------------------------------------------
' Sub: RequeryData (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formular-Datensätze aktualisieren
' </summary>
' <remarks>siehe <see cref="FormHandler::RequeryData">RequeryData</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub RequeryData()
   BaseFormHandler.RequeryData
End Sub

'---------------------------------------------------------------------------------------
' Sub: SetData (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formular-Datenquelle filtern und sortieren
' </summary>
' <param name="sFilterParams">Filterparameter</param>
' <param name="sOrderByParams">Sortierparamter</param>
' <remarks>siehe <see cref="FormHandler::SetData">SetData</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub SetData(Optional ByVal sFilterParams As String = vbNullString, Optional ByVal sOrderByParams As String = vbNullString)
   BaseFormHandler.SetData sFilterParams, sOrderByParams
End Sub

''---------------------------------------------------------------------------------------
'' Sub: SetDataFilterArray (26.04.2008)
''---------------------------------------------------------------------------------------
''/**
'' <summary>
'' Formular-Datenquelle filtern und sortieren
'' </summary>
'' <param name="filterParams">Collection mit FormFilterItems</param>
'' <param name="sOrderByParams">Sortierparamter</param>
'' <remarks>siehe <see cref="FormHandler::SetDataFilterArray">SetDataFilterArray</see> von FormHandler</remarks>
''**/
''---------------------------------------------------------------------------------------
'Public Sub SetDataFilterArray(Optional ByVal filterParams As Collection, _
'                       Optional ByVal sOrderByParams As String = vbNullString)
'   BaseFormHandler.SetDataFilterArray filterParams, sOrderByParams
'End Sub

'---------------------------------------------------------------------------------------
' Sub: SetNoData (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formulardatenbasis entfernen und "NULL-Datensatz" anzeigen
' </summary>
' <remarks>siehe <see cref="FormHandler::SetNoData">SetNoData</see> von FormHandler</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub SetNoData()
   BaseFormHandler.SetNoData
End Sub

Private Sub m_FormHandler_AccessFormUnload(Cancel As Integer)

   RaiseEvent AccessFormUnload(Cancel)
   If Cancel = True Then
      Exit Sub
   End If
   ClearFormObjectReferences

End Sub

Private Sub m_FormHandler_BeforeCalculateRecordsetPosition()
   On Error Resume Next
   If Not (m_SelectCommandButton Is Nothing) Then
      m_SelectCommandButton.SetFocus
   End If
End Sub

Private Sub m_FormHandler_CloseForm(ByRef Cancel As Boolean)

   RaiseEvent CloseForm(Cancel)
   
   If Cancel = True Then
      Exit Sub
   End If

   ClearFormObjectReferences
   
End Sub

Private Sub m_FormHandler_Disposed()
   If Not m_Disposing Then Set m_FormHandler = Nothing
End Sub

Private Sub m_FormHandler_EditModeChanged(ByVal NewMode As Long)
   RaiseEvent EditModeChanged(NewMode)
End Sub

Private Sub m_FormHandler_FormHandlerError(ByVal ErrMsg As String, ByVal ErrNumber As Long, Response As Integer)
   RaiseEvent FormHandlerError(ErrMsg, ErrNumber, Response)
End Sub

Private Sub m_FormHandler_FormHandlerErrorMissingForm(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewForm As Object)
   RaiseEvent FormHandlerErrorMissingForm(ErrMsg, ErrNumber, NewForm)
End Sub

Private Sub m_FormHandler_FormHandlerErrorMissingDbConnection(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewCurrentDbConnection As Object)
   RaiseEvent FormHandlerErrorMissingDbConnection(ErrMsg, ErrNumber, NewCurrentDbConnection)
End Sub

Private Sub m_FormHandler_FormHandlerMessage(ByVal Msg As String, ByVal MsgNumber As Long)
   RaiseEvent FormHandlerMessage(Msg, MsgNumber)
End Sub


'--------------------------------------------------------------------
'
' ListForm-Zusatz
'


'---------------------------------------------------------------------------------------
' Property: ListMode (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Art der Listensteuerung (Anzeige- oder Auswahl-Formular)
' </summary>
' <returns>LF_ListMode</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ListMode() As LF_ListMode
   ListMode = m_ListMode
End Property

'---------------------------------------------------------------------------------------
' Property: ListMode (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Art der Listensteuerung (Anzeige- oder Auswahl-Formular)
' </summary>
' <param name="NewMode">LF_ListMode-Wert</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ListMode(NewMode As LF_ListMode)
'0...normale Liste (qCompanyList)
'1...Auswahlliste (qCompanySelectionList)

   m_ListMode = NewMode

   If Not (m_SelectCommandButton Is Nothing) Then
      With m_SelectCommandButton
      If NewMode = 1& Then
         .FontName = "Symbol"
         .Caption = "®"  ' entspricht -->
      Else
         .FontName = "Tahoma"
         .Caption = "?"
      End If
      End With
   End If

End Property

Public Sub SelectItem(Optional lOpenArg As Long = 0)

   Dim varItemIDValue As Variant

   If Len(SelectionReturnDataSourceFieldName) > 0 Then
      varItemIDValue = BaseFormHandler.ActiveForm.Controls(SelectionReturnDataSourceFieldName).Value
   End If
   
   RaiseEvent Selected(varItemIDValue, m_lngSelectShiftMode, lOpenArg)

End Sub


'-------------------------------------------------------------------------------------------
' CommandButton: Select
'
Private Property Get selectCommandButton() As Access.Control

   Set selectCommandButton = m_SelectCommandButton

End Property

Private Property Set selectCommandButton(ByRef FormControl As Access.Control)

   If Not m_SelectCommandButton Is Nothing Then
      Set m_SelectCommandButton = Nothing
   End If
   Set m_SelectCommandButton = FormControl
   
   If Not (m_SelectCommandButton Is Nothing) Then
      If m_Disposed Then m_Disposed = False
      'Auswahl noch einmal setzen um Button-Layout einzustellen
      ListMode = ListMode
   End If

End Property

Private Sub m_SelectCommandButton_Click()
   SelectItem 0&
End Sub

Private Sub m_SelectCommandButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ShiftMode = Shift
End Sub


' Optionen
Public Property Let ShiftMode(ByVal Shift As Integer)
   m_lngSelectShiftMode = Shift
End Property

Public Property Get ShiftMode() As Integer
   ShiftMode = m_lngSelectShiftMode
End Property

'zusätzliche Form-Events
Private Sub m_ListForm_DblClick(Cancel As Integer)
   SelectItem
End Sub

Private Sub m_ListForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ShiftMode = Shift
End Sub

' Detailbereich
Private Sub m_fsecDetail_DblClick(Cancel As Integer)
   SelectItem
End Sub

Private Sub m_ListForm_Unload(Cancel As Integer)
   ClearFormObjectReferences
End Sub

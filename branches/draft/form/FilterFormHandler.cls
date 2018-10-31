VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterFormHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: FilterFormHandler
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formularsteuerungsklasse für Formulare mit ungebundnen Filtersteuerelementen
' und einem Endlosformular für die Anzeige der gefilterten Daten
' </summary>
' <remarks>
' Vereinfacht die Steuerung von Access-Formularen zur Filterung von Daten
' (Hauptformular mit Filterfelder = FilterFormHandler, Unterformular in Endlosdarstellung = ListFormHandler)
'
' DRAFT! - Code muss noch mit ACLib-Modulen kombiniert werden
'
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/FilterFormHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/FormHandler.cls</use>
'  <use>form/ListFormHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

#Const IncludeOldDataModeCode = 1
Private Const DisableStartUpFilterDefaultValue = True

'/**
' <summary>
' Steuerung, wenn im Formular kein Filter gesetzt ist
' </summary>
' <list type="table">
'   <item><term>VM_ShowNothing (1)</term><description>Keine Datensätze anzeigen</description></item>
'   <item><term>VM_ShowAll (2)</term><description>Alle vorhandenen Datensätze anzeigen</description></item>
' </list>
'**/
Public Enum FF_ViewModeNoFilter
   VM_ShowNothing = 0
   VM_ShowAll = 1
End Enum

'/**
' <summary>
' Spezielle Auswahl von möglichen Formularsteuerelmenten mit besonderer Behandlung
' </summary>
' <remarks>1-99 sind identisch mit <see cref="FormHandler::FH_HandledControls">FH_HandledControls</see> aus FormHandler</remarks>
' <list type="table">
'   <item><term>FFC_CloseCommandButton (1)</term><description>Schließen-Schaltfläche</description></item>
'   <item><term>FFC_SelectionFirstControl (2)</term><description>Steuerelement, das beim deaktiviert/unsichtbar Setzen eines aktiven Steuerelements den Fokus erhalten soll</description></item>
'   <item><term>FFC_EditModeCommandButton (3)</term><description>Schaltfläche für Bearbeitungsmoduswechsel (Readonly, Edit)</description></item>
'   <item><term>FFC_EditModeInfoLabel (4)</term><description>Label für Anzeige des aktuellen Bearbeitungsmodues</description></item>
'   <item><term>FFC_RequeryCommandButton (5)</term><description>Requery-Schaltfläche</description></item>
'   <item><term>FFC_SelectCommandButton (101)</term><description>Auswahl-Schaltfläche</description></item>
'   <item><term>FFC_MultipleChoiceCommitCommandButton (111)</term><description>Auswahlbestätigung bei Mehrfachauswahl</description></item>
'   <item><term>FFC_SelectCommandButton (112)</term><description></description></item>
'   <item><term>FFC_UseFilterCommandButton (201)</term><description>"Filter anwenden"-Schaltfläche</description></item>
'   <item><term>FFC_ClearFilterCommandButton (202)</term><description>"Filter entfernen"-Schaltfläche</description></item>
' </list>
'**/
Public Enum FF_HandledControls
   FFC_CloseCommandButton = 1
   FFC_SelectionFirstControl = 2
   FFC_EditModeCommandButton = 3
   FFC_EditModeInfoLabel = 4
   FFC_RequeryCommandButton = 5
   FFC_SelectCommandButton = 101
   FFC_MultiSelectionCommitCommandButton = 111
   FFC_MultiSelectionMarkAllCheckBox = 112
   FFC_UseFilterCommandButton = 201
   FFC_ClearFilterCommandButton = 202
End Enum

#If IncludeOldDataModeCode = 1 Then

Public Enum FF_FieldType
   [_FT_undefined] = 0
   FT_Number = 1
   FT_Text = 2
   FT_Date = 3
   FT_Boolean = 4
   FT_DateTime = 5
   FT_Time = 6
End Enum

Public Enum FF_FilterMode
   [_FM_undefined] = 0
   FM_Equal = 1
   FM_Larger = 2
   FM_Smaller = 4
   FM_Between = 8
   FM_Like = 16
   FM_WildcardRight = 32
   FM_WildcardLeft = 64
   FM_SearchEngine = 128
End Enum

#End If

'Allg. FormHandler (= Hauptform)
Private WithEvents m_FormHandler As FormHandler
Attribute m_FormHandler.VB_VarHelpID = -1

' ListFormHandler (kann im Prinzip auch der selbe Objektverweis wie m_FormHandler sein, ist aber normalerweise UF)
Private WithEvents m_ListFormHandler As ListFormHandler
Attribute m_ListFormHandler.VB_VarHelpID = -1

'Filter-Collection (FormFilterItems)
Private WithEvents m_FilterControlMngr As FilterControlManager
Attribute m_FilterControlMngr.VB_VarHelpID = -1

' Button UseFilter
Private WithEvents m_UseFilterCommandButton As Access.CommandButton
Attribute m_UseFilterCommandButton.VB_VarHelpID = -1

' Button ClearFilter
Private WithEvents m_ClearFilterCommandButton As Access.CommandButton
Attribute m_ClearFilterCommandButton.VB_VarHelpID = -1

'Variablen
Private m_OrderParams As String
Private m_AdditionialFilter As String

Private m_DisableStartUpFilter As Boolean
Private m_FirstUserApplyFilterActivated As Boolean

Private m_lngUseFilterShiftMode As Long
Private m_Disposing As Boolean
Private m_Disposed As Boolean

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
' Ereignis für Fehlermeldungen
' </summary>
' <param name="ErrMsg">Fehlermeldung</param>
' <param name="ErrNumber">Fehlernummer</param>
' <param name="Response">AcUtilErrorResumeMode-Wert</param>
' <remarks>Ermöglich die Behebung des Fehlers bzw. die Behandlung des Fehlers im Frontend</remarks>
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
' Ereignis für Fehlermeldungen, wenn Referenz des Access-Forms auf Nothing steht und benötigt wird.
' </summary>
' <param name="ErrMsg">Fehlermeldung</param>
' <param name="ErrNumber">Fehlernummer</param>
' <param name="NewForm">Damit kann vom Access-FE erneut das Formobjekt übergeben werden.</param>
' <remarks></remarks>
'**/
Public Event FormHandlerErrorMissionListForm(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewForm As Object)

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
' Ereignis-Meldung wenn Access-ListForm-Unload-Ereignis eintritt
' </summary>
' <param name="Cancel">ermöglicht Abbruch der Unload-Aktion</param>
' <remarks>Wird hauptsächlich für sauberes Entladen der Objekt-Referenzen benötigt</remarks>
'**/
Public Event AccessListFormUnload(ByRef Cancel As Integer)

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

'/**
' <summary>
' Filterwert über Formular berechnen lassen
' </summary>
' <param name="ItemID">Ausgewählte Datensatz-Kennung</param>
' <param name="ShiftMode">Kennung für Zusatztaste (Umschalt und/oder Strg)</param>
' <param name="OpenArg">Optionale Paramter fürs Öffnen</param>
' <remarks></remarks>
'**/
Public Event FilterValueRequest(ByVal FilterValueCode As String, ByRef FilterValue As String)

'Init/Terminate
Private Sub Class_Initialize()
   Set m_FormHandler = New FormHandler
   m_DisableStartUpFilter = DisableStartUpFilterDefaultValue
End Sub

Private Sub Class_Terminate()
   If Not m_Disposed Then Dispose
End Sub

Public Property Get Self() As FilterFormHandler
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

Private Sub ClearFormObjectReferences()

On Error Resume Next

   ' Button UseFilter
   Set m_UseFilterCommandButton = Nothing
   
   ' Button ClearFilter
   Set m_ClearFilterCommandButton = Nothing
   
   Set m_ListFormHandler = Nothing
   Set m_FormHandler = Nothing
   
End Sub

Private Sub ClearObjectReferences()

On Error Resume Next

   DeleteFilters
   Set m_FilterControlMngr = Nothing

   If Not (m_ListFormHandler Is Nothing) Then
      m_ListFormHandler.Dispose
      Set m_ListFormHandler = Nothing
   End If
   
   If Not (m_FormHandler Is Nothing) Then
      m_FormHandler.Dispose
      Set m_FormHandler = Nothing
   End If

End Sub

Public Property Get DisableStartUpFilter() As Boolean
   DisableStartUpFilter = m_DisableStartUpFilter
End Property

Public Property Let DisableStartUpFilter(ByVal NewValue As Boolean)
   m_DisableStartUpFilter = NewValue
End Property

Public Property Get FilterControlManager() As FilterControlManager
   If m_FilterControlMngr Is Nothing Then
      Set m_FilterControlMngr = New FilterControlManager
      If Not (m_ListFormHandler Is Nothing) Then
         SetFilterControlMngerDataMode m_ListFormHandler.DataSourceType
      End If
   End If
   Set FilterControlManager = m_FilterControlMngr
End Property

Public Property Get FilterControls() As FilterControlCollection
   Set FilterControls = FilterControlManager.FilterControls
End Property

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
End Property

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
Public Property Get HandledControls(ByVal Index As FF_HandledControls) As Access.Control

   Select Case Index
      Case 0 To 99 'Formhandler
         Set HandledControls = BaseFormHandler.HandledControls(Index)
      Case 100 To 199 'Listform
         Set HandledControls = ListForm.HandledControls(Index)
      Case FF_HandledControls.FFC_UseFilterCommandButton
         Set HandledControls = UseFilterCommandButton
      Case FF_HandledControls.FFC_ClearFilterCommandButton
         Set HandledControls = ClearFilterCommandButton
      Case Else
         '???
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
Public Property Set HandledControls(ByVal Index As FF_HandledControls, ByRef FormControl As Access.Control)

   Select Case Index
      Case 0 To 99 'Formhandler
         Set BaseFormHandler.HandledControls(Index) = FormControl
      Case 100 To 199 'ListFormHandler
         Set ListForm.HandledControls(Index) = FormControl
      Case FF_HandledControls.FFC_UseFilterCommandButton
         Set UseFilterCommandButton = FormControl
      Case FF_HandledControls.FFC_ClearFilterCommandButton
         Set ClearFilterCommandButton = FormControl
      Case Else 'einfach an Base weitergeben und dort Fehlerbehandlung durchführen
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

   ListForm.InitDataSource lDataSourceType, sDataSource, sDataSourceNoData, sOptionalArg, sOptionalBindingDataSourceName
   
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
   ListForm.RequeryData
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
   ListForm.SetData sFilterParams, sOrderByParams
End Sub

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
   ListForm.SetNoData
End Sub

Private Sub m_FormHandler_AccessFormUnload(Cancel As Integer)
   RaiseEvent AccessFormUnload(Cancel)
   If Cancel = True Then
      Exit Sub
   End If
   ClearFormObjectReferences
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
' FilterForm-Zusatz
'


'---------------------------------------------------------------------------------------
' Property: ListMode (16.09.2007)
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
   ListMode = ListForm.ListMode
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
   ListForm.ListMode = NewMode
End Property


#If IncludeOldDataModeCode = 1 Then

Public Sub AddFilter(ByVal cFilterControl As Access.Control, _
                ByVal sFieldName As String, _
                ByVal lFieldType As FF_FieldType, _
                ByVal lFilterMode As FF_FilterMode, _
                Optional ByVal bNotErasable As Boolean = False, _
                Optional ByVal cFilterControl2 As Access.Control = Nothing, _
                Optional ByVal vNoFilterControlValue As Variant = Null)
                
   Dim EnableProcedureMode As Boolean
   Dim TempFilterMode As FF_FilterMode
   Dim RelOp As SqlRelationalOperators
                           
   If Not (m_ListFormHandler Is Nothing) Then
      Select Case m_ListFormHandler.DataSourceType
         Case FH_DataSourceType.DST_Procedure, FH_DataSourceType.DST_ProcedurePT, FH_DataSourceType.DST_DaoParamQueryDef
            EnableProcedureMode = True
      End Select
   End If
   
   If EnableProcedureMode Then
   
      If (lFilterMode And FM_Like) = FM_Like Then
         RelOp = SQL_Equal + SQL_UseLikeBehavior
         'Ausnahme: *Text*
         TempFilterMode = 0
         If (lFilterMode And FF_FilterMode.FM_WildcardRight) = FF_FilterMode.FM_WildcardRight Then
            TempFilterMode = TempFilterMode + FF_FilterMode.FM_WildcardRight
         End If
         If (lFilterMode And FF_FilterMode.FM_WildcardLeft) = FF_FilterMode.FM_WildcardLeft Then
            TempFilterMode = TempFilterMode + FF_FilterMode.FM_WildcardLeft
         End If
         RelOp = RelOp + ConverFilterMode2SqlRelationalOperators(TempFilterMode)
      ElseIf lFieldType = FT_Text Then
         RelOp = SQL_Equal + SQL_UseLikeBehavior
      Else
         RelOp = SQL_Equal
      End If

   Else
      RelOp = ConverFilterMode2SqlRelationalOperators(lFilterMode)
   End If
       
   With FilterControlManager.FilterControls.Add(sFieldName, ConvertFieldType2SqlFieldDataType(lFieldType), RelOp, cFilterControl, cFilterControl2, vNoFilterControlValue)
      .DisableRemoveFilterValue = bNotErasable
   End With

End Sub

Private Function ConvertFieldType2SqlFieldDataType(ByVal FT As FF_FieldType) As SqlFieldDataType
   
'Public Enum FF_FieldType
'   [_FT_undefined] = 0
'   FT_Number = 1
'   FT_Text = 2
'   FT_Date = 3
'   FT_Boolean = 4
'   FT_DateTime = 5
'   FT_Time = 6
'End Enum
'
'Public Enum SqlFieldDataType
'   SQL_Boolean = 1
'   SQL_Numeric = 2
'   SQL_Text = 3
'   SQL_Date = 4
'End Enum
   
   Dim DT As SqlFieldDataType
   
   Select Case FT
      Case FT_Number
         DT = SQL_Numeric
      Case FT_Text
         DT = SQL_Text
      Case FT_Boolean
         DT = SQL_Boolean
      Case FT_Date, FT_Time, FT_Date
         DT = SQL_Date
   End Select
   
   ConvertFieldType2SqlFieldDataType = DT

End Function

Private Function ConverFilterMode2SqlRelationalOperators(ByVal FM As FF_FilterMode) As SqlRelationalOperators

   Dim RelOp As SqlRelationalOperators

'Public Enum FF_FilterMode
'   [_FM_undefined] = 0
'   FM_Equal = 1
'   FM_Larger = 2
'   FM_Smaller = 4
'   FM_Between = 8
'   FM_Like = 16
'   FM_WildcardRight = 32
'   FM_WildcardLeft = 64
'   FM_SearchEngine = 128
'End Enum
'
'Public Enum SqlRelationalOperators
'   [_IgnoreAll] = &H80000000
'   SQL_Not = 1
'   SQL_Equal = 2
'   SQL_LessThan = 4
'   SQL_GreaterThan = 8
'   SQL_Like = 256
'   SQL_Between = 512
'   SQL_In = 1024
'   SQL_Add_WildCardSuffix = 2048
'   SQL_Add_WildCardPrefix = 4096
'   SQL_SplitValueToArray = 8192
'End Enum

   If (FM And FM_Equal) = FM_Equal Then
      RelOp = SQL_Equal
   End If

   If (FM And FM_Larger) = FM_Larger Then
      RelOp = RelOp + SQL_GreaterThan
   End If
   
   If (FM And FM_Smaller) = FM_Smaller Then
      RelOp = RelOp + SQL_LessThan
   End If
   
   If (FM And FM_Between) = FM_Between Then
      RelOp = RelOp + SQL_Between
   End If
   
   If (FM And FM_Like) = FM_Like Then
      RelOp = RelOp + SQL_Like
   End If
   
   If (FM And FM_WildcardRight) = FM_WildcardRight Then
      RelOp = RelOp + SQL_Add_WildCardSuffix
   End If
   
   If (FM And FM_WildcardLeft) = FM_WildcardLeft Then
      RelOp = RelOp + SQL_Add_WildCardPrefix
   End If
   
   If (FM And FM_SearchEngine) = FM_SearchEngine Then
      Err.Raise vbObjectError, "FilterFormHandler.ConverFilterMode2SqlRelationalOperators", "FilterMode 'SearchEngine' is not supported"
   End If
   
   ConverFilterMode2SqlRelationalOperators = RelOp

End Function

Public Property Get NoDirectFiltering() As Boolean
   NoDirectFiltering = Not FilterControlManager.AutoFilterOn
End Property

Public Property Let NoDirectFiltering(ByVal NewValue As Boolean)
   FilterControlManager.AutoFilterOn = Not NewValue
End Property

#End If

Public Property Get AdditionialFilter() As String
   AdditionialFilter = m_AdditionialFilter
End Property

Public Property Let AdditionialFilter(sFilter As String)
   m_AdditionialFilter = sFilter
End Property

Public Property Get OrderParams() As String
   OrderParams = m_OrderParams
End Property

Public Property Let OrderParams(sOrderParams As String)
   m_OrderParams = sOrderParams
End Property

Public Sub SetFilter(sFieldName As String, vFilterValue As Variant)
'xxx???xxx
   m_FilterControlMngr.FilterControls.Items(sFieldName).FilterValue = vFilterValue

End Sub

Public Property Get Filters() As FilterControlCollection

   If m_FilterControlMngr Is Nothing Then
      Set m_FilterControlMngr = New FilterControlManager
      If m_Disposed Then m_Disposed = False
   End If
   Set Filters = m_FilterControlMngr.FilterControls

End Property

Public Sub DeleteFilters()
   
On Error Resume Next
   If m_FilterControlMngr Is Nothing Then
      Exit Sub
   End If
   m_FilterControlMngr.FilterControls.RemoveControls
   
End Sub

Public Sub ClearFilters(Optional ViewMode = FF_ViewModeNoFilter.VM_ShowAll)

   If m_FilterControlMngr Is Nothing Then
      Exit Sub
   End If
   
   Filters.RemoveFilterValues False
   
   If ViewMode = FF_ViewModeNoFilter.VM_ShowAll Then 'alle DS anzeigen
      UseFilter True
   ElseIf ViewMode = FF_ViewModeNoFilter.VM_ShowNothing Then 'keine DS anzeigen
      ListForm.SetNoData
   End If

End Sub

Public Function GetFormFilter(Optional ByVal lDataSourceType As FH_DataSourceType = FH_DataSourceType.[_DST_Undefined]) As String

   Dim NewFilterString As String

   If lDataSourceType = FH_DataSourceType.[_DST_Undefined] Then
      lDataSourceType = ListForm.DataSourceType
   End If
   
   SetFilterControlMngerDataMode lDataSourceType
   
   NewFilterString = m_FilterControlMngr.CreateFilterString()
   
   If Len(m_AdditionialFilter) > 0 Then
      With New FilterStringBuilder
         Set .SqlTool = m_FilterControlMngr.SqlTool
         .AddCriteria NewFilterString
         .AddCriteria m_AdditionialFilter
         NewFilterString = .ToString(m_FilterControlMngr.FilterControls.ItemConcatOperator)
      End With
   End If
   
   GetFormFilter = NewFilterString

End Function

Public Sub UseFilter(Optional FilterNow As Boolean = True)
   ApplyFilter FilterNow
End Sub

Public Sub ApplyFilter(Optional FilterNow As Boolean = True)

   If Not m_FirstUserApplyFilterActivated Then
      m_FirstUserApplyFilterActivated = True
   End If

   With FilterControlManager
      If .AutoFilterOn Or FilterNow Then
         .ApplyFilter
      End If
   End With
   
End Sub

Private Sub m_FilterControlMngr_FilterStringChanged(ByVal NewFilterString As String)
   ApplyFilterString NewFilterString
End Sub

Private Sub ApplyFilterString(ByVal NewFilterString As String)

   If Not m_FirstUserApplyFilterActivated Then
   If m_DisableStartUpFilter Then
      Exit Sub
   End If
   End If

   If Len(m_AdditionialFilter) > 0 Then
      With New FilterStringBuilder
         Set .SqlTool = m_FilterControlMngr.SqlTool
         .AddCriteria NewFilterString
         .AddCriteria m_AdditionialFilter
         NewFilterString = .ToString(m_FilterControlMngr.FilterControls.ItemConcatOperator)
      End With
   End If
   
   With ListForm
      .SetData NewFilterString, m_OrderParams
   End With

End Sub

Private Sub m_FilterControlMngr_FilterValueRequest(ByVal FilterValueCode As String, FilterValue As String)
   RaiseEvent FilterValueRequest(FilterValueCode, FilterValue)
End Sub

Public Property Get ListForm() As ListFormHandler

   If m_ListFormHandler Is Nothing Then
      Set m_ListFormHandler = New ListFormHandler
   End If
   Set ListForm = m_ListFormHandler

End Property

Public Property Set ListForm(frm As ListFormHandler)

   If Not m_ListFormHandler Is Nothing Then
      Set m_ListFormHandler = Nothing
   End If
   Set m_ListFormHandler = frm
   
   SetFilterControlMngerDataMode frm.DataSourceType

End Property

Private Sub SetFilterControlMngerDataMode(ByVal DataSourceType As FH_DataSourceType)
   
   If m_FilterControlMngr Is Nothing Then
      Dim x As FilterControlManager
      Set x = FilterControlManager
      Exit Sub
   End If
   
   With m_FilterControlMngr
   
      Select Case DataSourceType
         Case FH_DataSourceType.DST_Procedure, FH_DataSourceType.DST_ProcedurePT
            .ConfigSqlFormat "'yyyymmdd hh:nn:ss'", "1", "%"
            .FilterControls.ItemConcatOperator = SQL_CommaSeparator
         Case FH_DataSourceType.DST_ADORS, FH_DataSourceType.DST_SQLPTado
            .ConfigSqlFormat "'yyyymmdd hh:nn:ss'", "1", "%"
            .FilterControls.ItemConcatOperator = SQL_And
         Case FH_DataSourceType.DST_DaoParamQueryDef
            .ConfigSqlFormat "'yyyymmdd hh:nn:ss'", "1", "%"
            .FilterControls.ItemConcatOperator = SQL_CommaSeparator
         Case FH_DataSourceType.DST_ADORSJET
            .ConfigSqlFormat "\#yyyy-mm-dd hh:nn:ss\#", "True", "%"
            .FilterControls.ItemConcatOperator = SQL_And
         Case Else 'DAO-SQL
            .ConfigSqlFormat "\#yyyy-mm-dd hh:nn:ss\#", "True", "*"
            .FilterControls.ItemConcatOperator = SQL_And
      End Select

   End With
   
End Sub

Private Sub m_ListFormHandler_AccessFormUnload(Cancel As Integer)
   RaiseEvent AccessListFormUnload(Cancel)
End Sub

Private Sub m_ListFormHandler_CloseForm(Cancel As Boolean)

On Error Resume Next
   RaiseEvent CloseForm(Cancel)

   If Cancel = True Then
      Exit Sub
   End If

   ClearFormObjectReferences
   
End Sub

Private Sub m_ListFormHandler_Disposed()
   If Not m_Disposing Then Set m_ListFormHandler = Nothing
End Sub

Private Sub m_ListFormHandler_EditModeChanged(ByVal NewMode As Long)
   RaiseEvent EditModeChanged(NewMode)
End Sub

Private Sub m_ListFormHandler_FormHandlerError(ByVal ErrMsg As String, ByVal ErrNumber As Long, Response As Integer)
   RaiseEvent FormHandlerError(ErrMsg, ErrNumber, Response)
End Sub

Private Sub m_ListFormHandler_FormHandlerErrorMissingDbConnection(ByVal ErrMsg As String, ByVal ErrNumber As Long, NewCurrentDbConnection As Object)
   RaiseEvent FormHandlerErrorMissingDbConnection(ErrMsg, ErrNumber, NewCurrentDbConnection)
End Sub

Private Sub m_ListFormHandler_FormHandlerErrorMissingForm(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewForm As Object)
   RaiseEvent FormHandlerErrorMissionListForm(ErrMsg, ErrNumber, NewForm)
End Sub

Private Sub m_ListFormHandler_FormHandlerMessage(ByVal Msg As String, ByVal MsgNumber As Long)
   RaiseEvent FormHandlerMessage(Msg, MsgNumber)
End Sub

Private Sub m_ListFormHandler_Selected(ByVal ItemID As Variant, ByVal ShiftMode As Long, ByVal OpenArg As Long)
   RaiseEvent Selected(ItemID, ShiftMode, OpenArg)
End Sub

'-------------------------------------------------------------------------------------------
' CommandButton: UseFilter
'
Private Property Get UseFilterCommandButton() As Access.Control

   Set UseFilterCommandButton = m_UseFilterCommandButton

End Property

Private Property Set UseFilterCommandButton(ByRef FormControl As Access.Control)

   If Not (m_UseFilterCommandButton Is Nothing) Then
      Set m_UseFilterCommandButton = Nothing
   End If
   Set m_UseFilterCommandButton = FormControl
   
   If Not (m_UseFilterCommandButton Is Nothing) Then
      If m_Disposed Then m_Disposed = False
   End If

End Property

Private Sub m_UseFilterCommandButton_Click()
   UseFilter True
End Sub

Private Sub m_UseFilterCommandButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ShiftMode = Shift
End Sub


'-------------------------------------------------------------------------------------------
' CommandButton: ClearFilter
'
Private Property Get ClearFilterCommandButton() As Access.Control

   Set ClearFilterCommandButton = m_ClearFilterCommandButton

End Property

Private Property Set ClearFilterCommandButton(ByRef FormControl As Access.Control)

   If Not (m_ClearFilterCommandButton Is Nothing) Then
      Set m_ClearFilterCommandButton = Nothing
   End If
   Set m_ClearFilterCommandButton = FormControl
   
   If Not (m_ClearFilterCommandButton Is Nothing) Then
      If m_Disposed Then m_Disposed = False
   End If

End Property

Private Sub m_ClearFilterCommandButton_Click()
   ClearFilters FF_ViewModeNoFilter.VM_ShowAll
End Sub

Private Sub m_ClearFilterCommandButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ShiftMode = Shift
End Sub

Public Property Let ShiftMode(ByVal Shift As Integer)
   m_lngUseFilterShiftMode = Shift
End Property

Public Property Get ShiftMode() As Integer
   ShiftMode = m_lngUseFilterShiftMode
End Property

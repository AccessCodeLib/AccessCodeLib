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
' Class: form.FilterFormHandler
'---------------------------------------------------------------------------------------
'
' Form control class for forms with unbound filter controls
' and a continuous form for displaying the filtered data
'
' Remarks:
'     Simplifies the control of Access forms for filtering data
'     (main form with filter fields = FilterFormHandler, subform in continuous view = ListFormHandler)
'
'     DRAFT! - code still has to be combined with ACLib modules
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/FilterFormHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/FormHandler.cls</use>
'  <use>form/ListFormHandler.cls</use>
'  <description>Form controller for a search form with unbound filter controls and a continuous (list) form that shows the filtered data; based on FormHandler/ListFormHandler</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

#Const IncludeOldDataModeCode = 1
Private Const DisableStartUpFilterDefaultValue = True

'---------------------------------------------------------------------------------------
' Enum: FF_ViewModeNoFilter
'---------------------------------------------------------------------------------------
'
' Behaviour if no filter is set in the form
'
' VM_ShowNothing (1) - Show no records
' VM_ShowAll (2) - Show all records
'
'---------------------------------------------------------------------------------------
Public Enum FF_ViewModeNoFilter
   VM_ShowNothing = 0
   VM_ShowAll = 1
End Enum

'---------------------------------------------------------------------------------------
' Enum: FF_HandledControls
'---------------------------------------------------------------------------------------
'
' Form controls with special handling
'
' FFC_CloseCommandButton (1) - Close button
' FFC_SelectionFirstControl (2) - Control that receives the focus when an active control is disabled/hidden
' FFC_EditModeCommandButton (3) - Button for switching the edit mode (read-only, edit)
' FFC_EditModeInfoLabel (4) - Label showing the current edit mode
' FFC_RequeryCommandButton (5) - Requery button
' FFC_SelectCommandButton (101) - Select button
' FFC_MultiSelectionCommitCommandButton (111) - Confirmation of a multiple selection
' FFC_MultiSelectionMarkAllCheckBox (112) - Check box for marking all records of a multiple selection
' FFC_UseFilterCommandButton (201) - "Apply filter" button
' FFC_ClearFilterCommandButton (202) - "Remove filter" button
'
' Remarks:
'     1-99 are identical to <FormHandler.FH_HandledControls>
'
'---------------------------------------------------------------------------------------
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

'---------------------------------------------------------------------------------------
' Event: FormHandlerMessage
'---------------------------------------------------------------------------------------
'
' Event for messages to be handled in the frontend.
'
' Parameters:
'     Msg       - Message text
'     MsgNumber - Message number
'
'---------------------------------------------------------------------------------------
Public Event FormHandlerMessage(ByVal Msg As String, ByVal MsgNumber As Long)

'---------------------------------------------------------------------------------------
' Event: FormHandlerError
'---------------------------------------------------------------------------------------
'
' Event for error messages
'
' Parameters:
'     ErrMsg    - Error message
'     ErrNumber - Error number
'     Response  - AcUtilErrorResumeMode value
'
' Remarks:
'     Allows the error to be fixed or handled in the frontend
'
'---------------------------------------------------------------------------------------
Public Event FormHandlerError(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef Response As Integer)

'---------------------------------------------------------------------------------------
' Event: FormHandlerErrorMissingForm
'---------------------------------------------------------------------------------------
'
' Event for error messages if the reference to the Access form is Nothing but required.
'
' Parameters:
'     ErrMsg    - Error message
'     ErrNumber - Error number
'     NewForm   - Allows the Access frontend to pass the form object again.
'
'---------------------------------------------------------------------------------------
Public Event FormHandlerErrorMissingForm(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewForm As Object)

'---------------------------------------------------------------------------------------
' Event: FormHandlerErrorMissionListForm
'---------------------------------------------------------------------------------------
'
' Event for error messages if the reference to the Access form is Nothing but required.
'
' Parameters:
'     ErrMsg    - Error message
'     ErrNumber - Error number
'     NewForm   - Allows the Access frontend to pass the form object again.
'
'---------------------------------------------------------------------------------------
Public Event FormHandlerErrorMissionListForm(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewForm As Object)

'---------------------------------------------------------------------------------------
' Event: FormHandlerErrorMissingDbConnection
'---------------------------------------------------------------------------------------
'
' Event for error messages if the reference to <DbConnection> is Nothing but required.
'
' Parameters:
'     ErrMsg                 - Error message
'     ErrNumber              - Error number
'     NewCurrentDbConnection - Allows the Access frontend to pass the DbConnection object (<DbConnection>) again.
'
'---------------------------------------------------------------------------------------
Public Event FormHandlerErrorMissingDbConnection(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewCurrentDbConnection As Object)

'---------------------------------------------------------------------------------------
' Event: CloseForm
'---------------------------------------------------------------------------------------
'
' Event raised when the Access form is closed
'
' Parameters:
'     Cancel - allows the close action to be cancelled
'
'---------------------------------------------------------------------------------------
Public Event CloseForm(ByRef Cancel As Boolean)

'---------------------------------------------------------------------------------------
' Event: AccessFormUnload
'---------------------------------------------------------------------------------------
'
' Event raised when the Access form Unload event occurs
'
' Parameters:
'     Cancel - allows the unload action to be cancelled
'
' Remarks:
'     Mainly required for cleanly releasing the object references
'
'---------------------------------------------------------------------------------------
Public Event AccessFormUnload(ByRef Cancel As Integer)

'---------------------------------------------------------------------------------------
' Event: AccessListFormUnload
'---------------------------------------------------------------------------------------
'
' Event raised when the Unload event of the Access list form occurs
'
' Parameters:
'     Cancel - allows the unload action to be cancelled
'
' Remarks:
'     Mainly required for cleanly releasing the object references
'
'---------------------------------------------------------------------------------------
Public Event AccessListFormUnload(ByRef Cancel As Integer)

'---------------------------------------------------------------------------------------
' Event: EditModeChanged
'---------------------------------------------------------------------------------------
'
' Event raised when the edit mode has changed
'
' Parameters:
'     NewMode - Newly set mode
'
'---------------------------------------------------------------------------------------
Public Event EditModeChanged(ByVal NewMode As Long)

'---------------------------------------------------------------------------------------
' Event: Selected
'---------------------------------------------------------------------------------------
'
' Event raised when a record is selected
'
' Parameters:
'     ItemID    - Identifier of the selected record
'     ShiftMode - Modifier key identifier (Shift and/or Ctrl)
'     OpenArg   - Optional parameters for opening
'
'---------------------------------------------------------------------------------------
Public Event Selected(ByVal ItemID As Variant, ByVal ShiftMode As Long, ByVal OpenArg As Long)

'---------------------------------------------------------------------------------------
' Event: Disposed
'---------------------------------------------------------------------------------------
'
' Event raised when object references have been removed
'
'---------------------------------------------------------------------------------------
Public Event Disposed()

'---------------------------------------------------------------------------------------
' Event: FilterValueRequest
'---------------------------------------------------------------------------------------
'
' Let the form calculate a filter value
'
' Parameters:
'     FilterValueCode - Code of the requested filter value
'     FilterValue     - (ByRef) calculated filter value
'
'---------------------------------------------------------------------------------------
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
' Property: BaseFormHandler
'---------------------------------------------------------------------------------------
'
' Returns the reference to the FormHandler base class (<FormHandler>).
'
' Returns:
'     FormHandler base class (<FormHandler>)
'
' Remarks:
'     Only implemented for additional code extensions.
'
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
' Sub: CloseForm
'---------------------------------------------------------------------------------------
'
' Close the Access form
'
' Remarks:
'     see <FormHandler.CloseForm>
'
'---------------------------------------------------------------------------------------
Public Sub CloseForm()
   BaseFormHandler.CloseForm
End Sub

'---------------------------------------------------------------------------------------
' Property: DefaultOrderParams
'---------------------------------------------------------------------------------------
'
' Default sort parameters
'
' Parameters:
'     sOrderParams - Parameters for the default sort order
'
' Remarks:
'     see <FormHandler.DefaultOrderParams>
'
'---------------------------------------------------------------------------------------
Public Property Let DefaultOrderParams(sOrderParams As String)
   BaseFormHandler.DefaultOrderParams = sOrderParams
End Property

'---------------------------------------------------------------------------------------
' Property: ActiveDbConnection
'---------------------------------------------------------------------------------------
'
' ActiveDbConnection
'
' Returns:
'     DbConnection
'
' Remarks:
'     see <FormHandler.ActiveDbConnection>
'
'---------------------------------------------------------------------------------------
Public Property Get ActiveDbConnection() As DbConnectionHandler
   Set ActiveDbConnection = BaseFormHandler.ActiveDbConnection
End Property

'---------------------------------------------------------------------------------------
' Property: ActiveDbConnection
'---------------------------------------------------------------------------------------
'
' Assign a new <DbConnection>
'
' Parameters:
'     NewDbConnection - New DbConnection object
'
' Remarks:
'     see <FormHandler.ActiveDbConnection>
'
'---------------------------------------------------------------------------------------
Public Property Set ActiveDbConnection(ByRef NewDbConnection As DbConnectionHandler)
   Set BaseFormHandler.ActiveDbConnection = NewDbConnection
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceType
'---------------------------------------------------------------------------------------
'
' DataSourceType
'
' Parameters:
'     NewDataSourceType - FH_DataSourceType for setting the data source
'
' Remarks:
'     see <FormHandler.DataSourceType>
'
'---------------------------------------------------------------------------------------
Public Property Let DataSourceType(NewDataSourceType As FH_DataSourceType)
   BaseFormHandler.DataSourceType = NewDataSourceType
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceType
'---------------------------------------------------------------------------------------
'
' DataSourceType
'
' Returns:
'     FH_DataSourceType
'
' Remarks:
'     see <FormHandler.DataSourceType>
'
'---------------------------------------------------------------------------------------
Public Property Get DataSourceType() As FH_DataSourceType
   DataSourceType = BaseFormHandler.DataSourceType
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceUniqueFieldName
'---------------------------------------------------------------------------------------
'
' Field name of the record source that returns the unique record identifier
'
' Parameters:
'     FieldName - Field name of the record source
'
' Remarks:
'     see <FormHandler.DataSourceUniqueFieldName>
'
'---------------------------------------------------------------------------------------
Public Property Let DataSourceUniqueFieldName(ByVal FieldName As String)
   BaseFormHandler.DataSourceUniqueFieldName = FieldName
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceUniqueFieldName
'---------------------------------------------------------------------------------------
'
' Field name of the record source that returns the unique record identifier
'
' Returns:
'     String
'
' Remarks:
'     see <FormHandler.DataSourceUniqueFieldName>
'
'---------------------------------------------------------------------------------------
Public Property Get DataSourceUniqueFieldName() As String
   DataSourceUniqueFieldName = BaseFormHandler.DataSourceUniqueFieldName
End Property

'---------------------------------------------------------------------------------------
' Property: ActiveForm
'---------------------------------------------------------------------------------------
'
' Access form object in use
'
' Returns:
'     Access.Form
'
' Remarks:
'     see <FormHandler.ActiveForm>
'
'---------------------------------------------------------------------------------------
Public Property Get ActiveForm() As Access.Form
   Set ActiveForm = BaseFormHandler.ActiveForm
End Property

'---------------------------------------------------------------------------------------
' Property: ActiveForm
'---------------------------------------------------------------------------------------
'
' Access form object to be controlled by the handler
'
' Parameters:
'     FormRef - Access form to be passed
'
' Remarks:
'     see <FormHandler.ActiveForm>
'
'---------------------------------------------------------------------------------------
Public Property Set ActiveForm(ByRef FormRef As Access.Form)
   Set BaseFormHandler.ActiveForm = FormRef
End Property

'---------------------------------------------------------------------------------------
' Property: HandledControls
'---------------------------------------------------------------------------------------
'
' Control(s) with special handling
'
' Parameters:
'     Index - LF_HandledControls identifier of the control to select
'
' Returns:
'     Access.Control
'
' Remarks:
'     see <FormHandler.HandledControls>
'
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
' Property: HandledControls
'---------------------------------------------------------------------------------------
'
' Control(s) with special handling
'
' Parameters:
'     Index       - LF_HandledControls identifier of the control to select
'     FormControl - Control to be assigned to the LF_HandledControls identifier
'
' Remarks:
'     see <FormHandler.HandledControls>
'
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
' Sub: InitDataSource
'---------------------------------------------------------------------------------------
'
' Set the data source
'
' Parameters:
'     lDataSourceType   - Type of data source
'     sDataSource       - Definition of the data source (e.g. SQL statement)
'     sDataSourceNoData - Definition of the data source for an empty list
'     sOptionalArg      - Optional parameter value
'     sOptionalBindingDataSourceName - (optional) Name of the data source used for binding
'
' Remarks:
'     see <FormHandler.InitDataSource>
'
'---------------------------------------------------------------------------------------
Public Sub InitDataSource(ByVal lDataSourceType As FH_DataSourceType, _
             Optional ByVal sDataSource As String = vbNullString, _
             Optional ByVal sDataSourceNoData As String = vbNullString, _
             Optional ByVal sOptionalArg As String = vbNullString, _
             Optional ByVal sOptionalBindingDataSourceName As String = vbNullString)

   ListForm.InitDataSource lDataSourceType, sDataSource, sDataSourceNoData, sOptionalArg, sOptionalBindingDataSourceName
   
End Sub

'---------------------------------------------------------------------------------------
' Sub: RequeryData
'---------------------------------------------------------------------------------------
'
' Requery the form records
'
' Remarks:
'     see <FormHandler.RequeryData>
'
'---------------------------------------------------------------------------------------
Public Sub RequeryData()
   ListForm.RequeryData
End Sub

'---------------------------------------------------------------------------------------
' Sub: SetData
'---------------------------------------------------------------------------------------
'
' Filter and sort the form record source
'
' Parameters:
'     sFilterParams  - Filter parameters
'     sOrderByParams - Sort parameters
'
' Remarks:
'     see <FormHandler.SetData>
'
'---------------------------------------------------------------------------------------
Public Sub SetData(Optional ByVal sFilterParams As String = vbNullString, Optional ByVal sOrderByParams As String = vbNullString)
   ListForm.SetData sFilterParams, sOrderByParams
End Sub

'---------------------------------------------------------------------------------------
' Sub: SetNoData
'---------------------------------------------------------------------------------------
'
' Remove the form record source and show a "NULL record"
'
' Remarks:
'     see <FormHandler.SetNoData>
'
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
' Property: ListMode
'---------------------------------------------------------------------------------------
'
' Type of list control (display or selection form)
'
' Returns:
'     LF_ListMode
'
'---------------------------------------------------------------------------------------
Public Property Get ListMode() As LF_ListMode
   ListMode = ListForm.ListMode
End Property

'---------------------------------------------------------------------------------------
' Property: ListMode
'---------------------------------------------------------------------------------------
'
' Type of list control (display or selection form)
'
' Parameters:
'     NewMode - LF_ListMode value
'
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
      Dim X As FilterControlManager
      Set X = FilterControlManager
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

Private Sub m_UseFilterCommandButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub m_ClearFilterCommandButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ShiftMode = Shift
End Sub

Public Property Let ShiftMode(ByVal Shift As Integer)
   m_lngUseFilterShiftMode = Shift
End Property

Public Property Get ShiftMode() As Integer
   ShiftMode = m_lngUseFilterShiftMode
End Property

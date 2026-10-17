VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListFormHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: form.ListFormHandler
'---------------------------------------------------------------------------------------
'
' Form control class for lists/continuous forms
'
' Remarks:
'     Simplifies the control of Access forms that display lists as continuous forms.
'
'     DRAFT! - code still has to be combined with ACLib modules
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/ListFormHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/FormHandler.cls</use>
'  <description>Form controller for list/continuous forms based on FormHandler: list mode, default sort order, selection event</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'---------------------------------------------------------------------------------------
' Enum: LF_HandledControls
'---------------------------------------------------------------------------------------
'
' Form controls with special handling
'
' FC_CloseCommandButton (1) - Close button
' FC_SelectionFirstControl (2) - Control that receives the focus when an active control is disabled/hidden
' FC_EditModeCommandButton (3) - Button for switching the edit mode (read-only, edit)
' FC_EditModeInfoLabel (4) - Label showing the current edit mode
' FC_RequeryCommandButton (5) - Requery button
' FC_SelectCommandButton (101) - Select button
'
' Remarks:
'     1-99 are identical to <FormHandler.FH_HandledControls>
'
'---------------------------------------------------------------------------------------
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

'---------------------------------------------------------------------------------------
' Enum: LF_ListMode
'---------------------------------------------------------------------------------------
'
' Selection options (incl. display)
'
' LM_Normal (0) - Normal continuous form / switch to detail form or similar when the select button is clicked
' LM_Select (1) - Continuous form for selecting one or more records
'
'---------------------------------------------------------------------------------------
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
' Event for error messages to be handled in the frontend.
'
' Parameters:
'     ErrMsg    - Error message
'     ErrNumber - Error number
'     Response  - AcUtilErrorResumeMode value
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

Private m_Disposing As Boolean
Private m_Disposed As Boolean

Private m_SelectionReturnDataSourceFieldName As String

#If UseMouseWheelFix Then
Private m_MouseWheelFix As FormMouseWheelFix
#End If

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
' Property: SelectionReturnDataSourceFieldName
'---------------------------------------------------------------------------------------
'
' Field name of the record source whose value is returned when a record is selected
'
' Parameters:
'     FieldName - Field name of the record source
'
' Remarks:
'     Default: <DataSourceUniqueFieldName>
'
'---------------------------------------------------------------------------------------
Public Property Let SelectionReturnDataSourceFieldName(ByVal FieldName As String)
   m_SelectionReturnDataSourceFieldName = FieldName
End Property

'---------------------------------------------------------------------------------------
' Property: SelectionReturnDataSourceFieldName
'---------------------------------------------------------------------------------------
'
' Field name of the record source whose value is returned when a record is selected
'
' Returns:
'     String
'
' Remarks:
'     Default: <DataSourceUniqueFieldName>
'
'---------------------------------------------------------------------------------------
Public Property Get SelectionReturnDataSourceFieldName() As String

   If Len(m_SelectionReturnDataSourceFieldName) = 0 Then
      m_SelectionReturnDataSourceFieldName = DataSourceUniqueFieldName
   End If
   SelectionReturnDataSourceFieldName = m_SelectionReturnDataSourceFieldName

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
   InitListForm FormRef
   
End Property

Private Sub InitListForm(ByRef NewForm As Access.Form)

   Set m_ListForm = NewForm
   If Not (m_ListForm Is Nothing) Then
      Set m_fsecDetail = m_ListForm.Section(0)
   End If
   
#If UseMouseWheelFix Then
   With New FormMouseWheelFix
      If .Init(m_ListForm) Then
         Set m_MouseWheelFix = .Self
      End If
   End With
#End If
   
End Sub

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
Public Property Get HandledControls(ByVal Index As LF_HandledControls) As Access.Control

   Select Case Index
      Case LF_HandledControls.LFC_SelectCommandButton
         Set HandledControls = selectCommandButton
      Case Else
         Set HandledControls = BaseFormHandler.HandledControls(Index)
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
Public Property Set HandledControls(ByVal Index As LF_HandledControls, ByRef FormControl As Access.Control)

   Select Case Index
      Case LF_HandledControls.LFC_SelectCommandButton
         Set selectCommandButton = FormControl
      Case Else
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
             
   BaseFormHandler.InitDataSource lDataSourceType, sDataSource, sDataSourceNoData, sOptionalArg, sOptionalBindingDataSourceName
   
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
   BaseFormHandler.RequeryData
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
   ListMode = m_ListMode
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

Private Sub m_SelectCommandButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub m_ListForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ShiftMode = Shift
End Sub

' Detailbereich
Private Sub m_fsecDetail_DblClick(Cancel As Integer)
   SelectItem
End Sub

Private Sub m_ListForm_Unload(Cancel As Integer)
   ClearFormObjectReferences
End Sub

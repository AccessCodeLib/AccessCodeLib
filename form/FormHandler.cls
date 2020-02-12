VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: FormHandler
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formularsteuerungsklasse für Access-Formulare (Basisklasse)
' </summary>
' <remarks>
' Vereinfacht die Steuerung von Access-Formularen.
'
' DRAFT! - Code muss noch mit ACLib-Modulen kombiniert werden
'
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/FormHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/DbConnectionHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'für DST_DaoParamQueryDef, damit Parameter aus String getrennt werden können
Private Const m_conSplitString As String = "|#@#|"

Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1

#If VBA7 Then
   Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                              ByVal Hwnd As LongPtr, ByVal wMsg As Long, _
                              ByVal wParam As LongPtr, lParam As Any) As Long
#Else
   Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                              ByVal Hwnd As Long, ByVal wMsg As Long, _
                              ByVal wParam As Long, lParam As Any) As Long
#End If
'/**
' <summary>
' Einstellmöglichkeit für Datenherkunft des Formulars
' </summary>
' <list type="table">
'   <item><term>DST_Undefined (0)</term><description>nur für Prozedurparameter, damit wird der in FormHanlder gespeicherte DataSourceType verwendet</description></item>
'   <item><term>DST_FormSource (1)</term><description>SQL-Quelle direkt an Form.RecordSource übergeben, Filter und Sortierung wird per Form.Filter und Form.OrderBy durchgeführt</description></item>
'   <item><term>DST_SQL (2)</term><description>SQL-Quelle inkl. Filter und Sortierung an Form.RecordSource übergeben</description></item>
'   <item><term>DST_DAO (3)</term><description>DAO-Recordsetobjekt an Form.Recordset übergeben</description></item>
'   <item><term>DST_ADO (4)</term><description>ADODB-Recordsetobjekt an Form.Recordset übergeben</description></item>
'   <item><term>DST_Procedure (5)</term><description>DBMS-Prozedur (T-SQL-Syntax) in ADODB-Recordset ausführen und ADODB-Recordset an Form.Recordset übergeben</description></item>
'   <item><term>DST_ProcedurePT (6)</term><description>DBMS-Prozedur (T-SQL-Syntax) in gespeicherter Pass-Through-Abfrage eintragen und PT-Abfragenamen an Form.RecordSource übergeben</description></item>
'   <item><term>DST_SQLPTdao (7)</term><description>SQL-Anweisung (DAO-Syntax) in Pass-Through-Abfrage eintragen und PT-Abfragenamen an Form.RecordSource übergeben</description></item>
'   <item><term>DST_SQLPTado (8)</term><description>SQL-Anweisung (ADO-Syntax) in Pass-Through-Abfrage eintragen und PT-Abfragenamen an Form.RecordSource übergeben</description></item>
'   <item><term>DST_DaoParamQueryDef (9)</term><description>An gespeicherte Abfrage Parameter-Werte übergeben und DAO-Recordset (aus Querydef.OpenRecordset) an Form.Recordset übergeben</description></item>
' </list>
'**/
Public Enum FH_DataSourceType
   [_DST_Undefined] = 0 'nur für Prozedurparameter, damit wird der in FormHanlder gespeicherte DataSourceType verwendet
   DST_FormSource = 1
   DST_SQL = 2
   DST_DAORS = 3
   DST_ADORS = 4
   DST_Procedure = 5
   DST_ProcedurePT = 6
   DST_SQLPTdao = 7
   DST_SQLPTado = 8
   DST_DaoParamQueryDef = 9
   DST_ADORSJET = 10
End Enum


'/**
' <summary>
' Spezielle Auswahl von möglichen Datenzugriffseinstellungen (Editmodus)
' </summary>
' <list type="table">
'   <item><term>EM_ReadOnly (0)</term><description>Schreibgeschützt</description></item>
'   <item><term>EM_Edit (1)</term><description>Änderung möglich</description></item>
'   <item><term>EM_New (2)</term><description>Neuer DS (Änderung möglich)</description></item>
'   <item><term>EM_Locked (3)</term><description>Gesperrt (Umschalten auf Edit nicht möglich)</description></item>
' </list>
'**/
Public Enum FH_EditMode
   EM_ReadOnly = 0
   EM_Edit = 1
   EM_New = 2
   EM_Locked = 3
End Enum


'/**
' <summary>
' Spezielle Auswahl von möglichen Formularsteuerelmenten mit besonderer Behandlung
' </summary>
' <list type="table">
'   <item><term>FHC_CloseCommandButton (1)</term><description>Schließen-Schaltfläche</description></item>
'   <item><term>FHC_SelectionFirstControl (2)</term><description>Steuerelement, das beim deaktiviert/unsichtbar Setzen eines aktiven Steuerelements den Fokus erhalten soll</description></item>
'   <item><term>FHC_EditModeCommandButton (3)</term><description>Schaltfläche für Bearbeitungsmoduswechsel (Readonly, Edit)</description></item>
'   <item><term>FHC_EditModeInfoLabel (4)</term><description>Label für Anzeige des aktuellen Bearbeitungsmodues</description></item>
'   <item><term>FHC_RequeryCommandButton (5)</term><description>Requery-Schaltfläche</description></item>
' </list>
'**/
Public Enum FH_HandledControls
   FHC_CloseCommandButton = 1
   FHC_SelectionFirstControl = 2
   FHC_EditModeCommandButton = 3
   FHC_EditModeInfoLabel = 4
   FHC_RequeryCommandButton = 5
End Enum

'Kennungart für die Tabellenerweiterung (Objekt-Instanz-Kennung)
Public Enum InstanceTabNameMode
    ITNM_NoInstanceTab = 0&         'Tabellenname wird _nicht_ mit einer Instanz-Kennung erweitert (reicht vollkommen aus, wenn Tab. nicht verküpft wird)
    ITNM_ApplicationInstance = 1&   'Erweiterung mit Kennung für die Application-Instanz
    ITNM_FormInstance = 2&          'Erweiterung mit Kennung für die Form-Instanz
End Enum


'DbConnection
Private m_DbConnection As DbConnectionHandler

' Datenbasis
Private m_DataSourceType As FH_DataSourceType
Private m_DataSource As String
Private m_DataSourceNoData As String
Private m_DataQueryDefName As String
Private m_BindingDataSourceName As String

Private m_FilterParams As String
Private m_OrderParams As String
Private m_DefaultOrderParams As String

' Feldname für DS-Kennung
Private m_strDataSourceUniqueFieldName As String

'Access-Application (aus Form-Objekt geholt)
Private m_AccessApplication As Access.Application

' Formular-Objekt
Private WithEvents m_AccessForm As Access.Form
Attribute m_AccessForm.VB_VarHelpID = -1

' Button Select
Private WithEvents m_CloseCommandButton As Access.CommandButton
Attribute m_CloseCommandButton.VB_VarHelpID = -1

' Button Requery
Private WithEvents m_RequeryCommandButton As Access.CommandButton
Attribute m_RequeryCommandButton.VB_VarHelpID = -1

' Button EditMode
Private WithEvents m_EditModeCommandButton As Access.CommandButton
Attribute m_EditModeCommandButton.VB_VarHelpID = -1
Private m_EditModeLabel As Access.Label
Private m_CurrentMode As FH_EditMode


' 1. Steuerelement in Auswahl
Private m_SelectFirstFormControl As Access.Control       'SetFocus

' Events

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
' Ereignis-Meldung wenn Objektreferenzen entfernt wurden
' </summary>
' <remarks></remarks>
'**/
Public Event Disposed()

'/**
' <summary>
' Ereignis-Meldung bevor Änderung der Datensatz-Postion berechnet wird
' </summary>
' <remarks>... z.B.: Fokus auf Detailbereich setzen</remarks>
'**/
Public Event BeforeCalculateRecordsetPosition()


Private m_Disposed As Boolean

Private Sub Class_Initialize()
   m_DataSourceType = FH_DataSourceType.[_DST_Undefined]
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub

Public Sub Dispose()

   If m_Disposed Then Exit Sub

On Error Resume Next

   ClearFormObjectReferences
   ClearObjectReferences

   m_Disposed = True
   RaiseEvent Disposed

End Sub

Private Sub ClearObjectReferences()

On Error Resume Next

   Set m_DbConnection = Nothing

End Sub

Private Sub ClearFormObjectReferences()

On Error Resume Next

   ' Button Select
   Set m_CloseCommandButton = Nothing

   ' Button+Label EditMode
   Set m_EditModeCommandButton = Nothing
   Set m_EditModeLabel = Nothing

   ' Button Requery
   Set m_RequeryCommandButton = Nothing

   '1. Steuerelement in Auswahl
   Set m_SelectFirstFormControl = Nothing

   'Formular-Objekt
   Set m_AccessForm = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Property: ActiveDbConnection (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DbConnection-Objekt das für die Bereitstellung der Daten benutzt wird
' </summary>
' <returns>DbConnection</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ActiveDbConnection() As DbConnectionHandler

   If (m_DbConnection Is Nothing) Then

      RaiseEvent FormHandlerErrorMissingDbConnection("FormHandler.ActiveDbConnection", 0&, m_DbConnection)

      If m_DbConnection Is Nothing Then
         Set m_DbConnection = DbCon
      End If

   End If
   Set ActiveDbConnection = m_DbConnection

End Property

'---------------------------------------------------------------------------------------
' Property: ActiveDbConnection (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zuweisung einer neuen <see cref="DbConnection">DbConnection</see>
' </summary>
' <param name="NewDbConnection">NewDbConnection</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set ActiveDbConnection(ByRef NewDbConnection As DbConnectionHandler)

   Set m_DbConnection = NewDbConnection
   If m_Disposed Then m_Disposed = False

End Property

'---------------------------------------------------------------------------------------
' Sub: CloseForm (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Access-Formular schließen
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub CloseForm()

   Dim bolCancel As Boolean

   RaiseEvent CloseForm(bolCancel)
   If bolCancel = True Then
      Exit Sub
   End If

   If Not (m_AccessForm Is Nothing) Then
      If m_AccessForm.Dirty Then
         m_AccessForm.Dirty = False
      End If
   End If

   On Error Resume Next 'falls Close abgebrochen wird

   DoCmd.Close acForm, ActiveForm.Name

End Sub


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
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub InitDataSource(ByVal lDataSourceType As FH_DataSourceType, _
                    Optional ByVal sDataSource As String = vbNullString, _
                    Optional ByVal sDataSourceNoData As String = vbNullString, _
                    Optional ByVal sOptionalArg As String = vbNullString, _
                    Optional ByVal sOptionalBindingDataSourceName As String = vbNullString)

   m_DataSourceType = lDataSourceType
   If m_DataSourceType <> FH_DataSourceType.DST_FormSource Then
      m_DataSource = sDataSource
      m_DataSourceNoData = sDataSourceNoData
      If m_DataSourceType = FH_DataSourceType.DST_ProcedurePT Or _
         m_DataSourceType = FH_DataSourceType.DST_SQLPTado Or _
         m_DataSourceType = FH_DataSourceType.DST_SQLPTdao Then
            m_DataQueryDefName = sOptionalArg
            m_BindingDataSourceName = sOptionalBindingDataSourceName
      End If
   End If

End Sub

'---------------------------------------------------------------------------------------
' Property: DataSourceType (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DataSourceType
' </summary>
' <returns>FH_DataSourceType</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DataSourceType() As FH_DataSourceType
   DataSourceType = m_DataSourceType
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceType (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DataSourceType
' </summary>
' <param name="NewDataSourceType">FH_DataSourceType zur Einstellung der Datenquelle</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let DataSourceType(newType As FH_DataSourceType)
   m_DataSourceType = newType
End Property

'---------------------------------------------------------------------------------------
' Sub: RequeryData (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formular-Datensätze aktualisieren
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub RequeryData()

   Dim varIDValue As Variant

   Dim lngPos As Long, lngPosDiff As Long
   Dim lngDirection As LongPtr
   Dim i As Long, lngRet As Long
   Dim bolIsListForm As Boolean

   bolIsListForm = IsListForm

On Error Resume Next
   If bolIsListForm Then
   With m_AccessForm
      lngPos = (.CurrentSectionTop - .Section(acHeader).Height) \ .Section(0).Height
   End With
   End If
   

'On Error GoTo HandleErr
   varIDValue = Null
   If Len(m_strDataSourceUniqueFieldName) > 0 Then
       varIDValue = m_AccessForm.Controls(m_strDataSourceUniqueFieldName).Value
'      With m_AccessForm.Recordset
'         If .RecordCount > 0 Then
'            varIDValue = .Fields(m_strDataSourceUniqueFieldName).Value
'         End If
'      End With
   End If

On Error GoTo 0
   SetData m_FilterParams, m_OrderParams

   If IsNull(varIDValue) = False Then
      'zu Bookmark springen
      m_AccessForm.Painting = False
      GotoBookmark m_strDataSourceUniqueFieldName, varIDValue

      'letzte Position wieder herstellen
      If bolIsListForm Then
On Error Resume Next
         With m_AccessForm
            If Not (m_SelectFirstFormControl Is Nothing) Then
               m_SelectFirstFormControl.SetFocus
            End If
            RaiseEvent BeforeCalculateRecordsetPosition
            lngPosDiff = lngPos - ((.CurrentSectionTop - .Section(acHeader).Height) \ .Section(0).Height)
            
         End With
On Error GoTo 0
         If lngPosDiff <> 0 Then
            If lngPosDiff > 0 Then
               lngDirection = SB_LINEUP
            Else
               lngDirection = SB_LINEDOWN
               lngPosDiff = Abs(lngPosDiff)
            End If
            For i = 1 To lngPosDiff
               lngRet = SendMessage(m_AccessForm.Hwnd, WM_VSCROLL, lngDirection, 0&)
            Next i
         End If
      End If
   End If

   m_AccessForm.Painting = True

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
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub SetData(Optional ByVal sFilterParams As String = vbNullString, _
               Optional ByVal sOrderByParams As String = vbNullString)

   SetAccessApplicationHourglass True

   m_FilterParams = sFilterParams
   m_OrderParams = sOrderByParams

   If Len(m_OrderParams) = 0 Then
      m_OrderParams = m_DefaultOrderParams
   End If

   Select Case m_DataSourceType
      Case FH_DataSourceType.DST_FormSource
         SetData_Form m_FilterParams, m_OrderParams
      Case FH_DataSourceType.DST_SQL
         SetData_SQL m_FilterParams, m_OrderParams
      Case FH_DataSourceType.DST_SQLPTdao, FH_DataSourceType.DST_SQLPTado
         SetData_SqlPT m_FilterParams, m_OrderParams
      Case FH_DataSourceType.DST_DAORS
         SetData_DaoRecordset m_FilterParams, m_OrderParams
      Case FH_DataSourceType.DST_ADORS, FH_DataSourceType.DST_ADORSJET
         SetData_AdoRecordset m_FilterParams, m_OrderParams
      Case FH_DataSourceType.DST_Procedure
         SetData_Procedure m_FilterParams, m_OrderParams
      Case FH_DataSourceType.DST_ProcedurePT
         SetData_ProcedurePT m_FilterParams, m_OrderParams
      Case FH_DataSourceType.DST_DaoParamQueryDef
         SetData_DaoRecordsetParamQueryDef m_FilterParams, m_OrderParams
      Case Else
         RaiseEvent FormHandlerMessage("Art der Datenbasis wurde nicht eingestellt!", 0&)
   End Select

   SetAccessApplicationHourglass False

End Sub


Private Sub SetData_Form(ByVal sFilterParams As String, _
                     ByVal sOrderByParams As String)

   With ActiveForm

      If Len(sFilterParams) > 0 Then
         .Filter = sFilterParams
         .FilterOn = True
      Else
         .Filter = vbNullString
         .FilterOn = False
      End If

      If Len(sOrderByParams) > 0 Then
         .OrderBy = sOrderByParams
         .OrderByOn = True
      Else
         .OrderBy = vbNullString
         .OrderByOn = False
      End If

   End With

End Sub

Private Sub SetData_SQL(ByVal sFilterParams As String, _
                  ByVal sOrderByParams As String)

   Dim strSQL As String

   strSQL = m_DataSource
   If Len(sFilterParams) > 0 Then
      strSQL = strSQL & " WHERE " & sFilterParams
   End If
   If Len(sOrderByParams) > 0 Then
      strSQL = strSQL & " ORDER BY " & sOrderByParams
   End If

   With ActiveForm
      .RecordSource = strSQL
   End With

End Sub

Private Sub SetData_SqlPT(ByVal sFilterParams As String, _
                    ByVal sOrderByParams As String)

   Dim strSQL As String
   Dim qdf As DAO.QueryDef
   Dim strQdfName As String
   
   strSQL = m_DataSource
   If Len(sFilterParams) > 0 Then
      strSQL = strSQL & " WHERE " & sFilterParams
   End If
   If Len(sOrderByParams) > 0 Then
      strSQL = strSQL & " ORDER BY " & sOrderByParams
   End If

   If Len(m_DataQueryDefName) > 0 Then
      strQdfName = m_DataQueryDefName
   Else
      strQdfName = "q_frm" & ActiveForm.Name
   End If
   Set qdf = ActiveDbConnection.CurrentDb.QueryDefs(strQdfName)
   qdf.SQL = strSQL
   qdf.Close
   Set qdf = Nothing

   If Len(m_BindingDataSourceName) > 0 Then
      strQdfName = m_BindingDataSourceName
   End If
   With ActiveForm
      .RecordSource = strQdfName
   End With

End Sub


Private Sub SetData_DaoRecordset(ByVal sFilterParams As String, _
                           ByVal sOrderByParams As String)

   Dim strSQL As String

   strSQL = m_DataSource
   If Len(sFilterParams) > 0 Then
      strSQL = strSQL & " WHERE " & sFilterParams
   End If
   If Len(sOrderByParams) > 0 Then
      strSQL = strSQL & " ORDER BY " & sOrderByParams
   End If

   With ActiveForm

      On Error Resume Next
      If Not .Recordset Is Nothing Then
         .Recordset.Close
         Set .Recordset = Nothing
      End If
      On Error GoTo 0
      Set .Recordset = ActiveDbConnection.DAO.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges, dbOptimistic)

   End With

End Sub

Private Sub SetData_DaoRecordsetParamQueryDef( _
                           ByVal sFilterParams As String, _
                           ByVal sOrderByParams As String)

   Dim qdf As DAO.QueryDef
   Dim strQdfName As String
   Dim lngParamCount As Long, i As Long
   Dim AstrFilterParams() As String
   Dim strFilterParam As String
   Dim strParamName As String
   Dim lngEqualPos As String
   Dim varParamValue As Variant

   AstrFilterParams = Split(sFilterParams, m_conSplitString)
   lngParamCount = UBound(AstrFilterParams) + 1


   If Len(m_DataQueryDefName) > 0 Then
      strQdfName = m_DataQueryDefName
   Else
      strQdfName = "q_frm" & ActiveForm.Name
   End If
   Set qdf = ActiveDbConnection.CurrentDb.QueryDefs(strQdfName)

   With qdf
      For i = 0 To (lngParamCount - 1)
         strFilterParam = AstrFilterParams(i)
         lngEqualPos = InStr(1, strFilterParam, "=", vbTextCompare)
         strParamName = Trim$(Left$(strFilterParam, lngEqualPos - 1))
         varParamValue = Trim(Mid(strFilterParam, lngEqualPos + 1))

         If varParamValue = "NULL" Then
            .Parameters(strParamName).Value = Null
         Else
         Select Case .Parameters(strParamName).Type
            Case DAO.DataTypeEnum.dbText, DAO.DataTypeEnum.dbChar, DAO.DataTypeEnum.dbMemo
               'Text
               If Left(varParamValue, 1) = "'" Then ' Hochkomma entfernen
                  varParamValue = Mid(varParamValue, 2, Len(varParamValue - 2))
               End If
               .Parameters(strParamName).Value = Str$(varParamValue)
            Case DAO.DataTypeEnum.dbByte, DAO.DataTypeEnum.dbInteger, DAO.DataTypeEnum.dbLong
               'Ganzzahlen
               .Parameters(strParamName).Value = CLng(Val(varParamValue))
            Case DAO.DataTypeEnum.dbDecimal, DAO.DataTypeEnum.dbDouble, DAO.DataTypeEnum.dbFloat, DAO.DataTypeEnum.dbSingle, DAO.DataTypeEnum.dbNumeric
               'Kommazahlen
               .Parameters(strParamName).Value = CDbl(Val(varParamValue))
            Case DAO.DataTypeEnum.dbDate, DAO.DataTypeEnum.dbTime
               'Datums/Zeitwerte
               .Parameters(strParamName).Value = CDate(Val(varParamValue))
            Case Else
               'Einfach Text probieren
               'Text
               If Left(varParamValue, 1) = "'" Then ' Hochkomma entfernen
                  varParamValue = Mid(varParamValue, 2, Len(varParamValue - 2))
               End If
               .Parameters(strParamName).Value = Str$(varParamValue)
         End Select
         End If
      Next
   End With

   With ActiveForm

      On Error Resume Next
      If Not .Recordset Is Nothing Then
         .Recordset.Close
         Set .Recordset = Nothing
      End If
      On Error GoTo 0
      Set .Recordset = qdf.OpenRecordset(dbOpenDynaset, dbSeeChanges, dbOptimistic)

   End With

   qdf.Close
   Set qdf = Nothing

End Sub

Private Sub SetData_AdoRecordset(ByVal sFilterParams As String, _
                           ByVal sOrderByParams As String)

   Dim strSQL As String

   strSQL = m_DataSource
   If Len(sFilterParams) > 0 Then
      strSQL = strSQL & " WHERE " & sFilterParams
   End If
   If Len(sOrderByParams) > 0 Then
      strSQL = strSQL & " ORDER BY " & sOrderByParams
   End If

   With ActiveForm
   
      On Error Resume Next
      If Not .Recordset Is Nothing Then
         .Recordset.Close
         Set .Recordset = Nothing
      End If
      On Error GoTo 0
      Set .Recordset = ActiveDbConnection.ADODB.OpenRecordset(strSQL, adOpenKeyset, adLockOptimistic, adUseClient)
   End With

End Sub

Private Sub SetData_Procedure(ByVal sFilterParams As String, _
                       ByVal sOrderByParams As String)

   Dim strSQL As String

   strSQL = "exec " & m_DataSource
   If Len(sFilterParams) > 0 Then
      strSQL = strSQL & " " & sFilterParams
   End If

   If Len(m_OrderParams) > 0 Then
      If Len(sFilterParams) > 0 Then
         strSQL = strSQL & ","
      End If
      strSQL = strSQL & " " & sOrderByParams
   End If

   With ActiveForm

      On Error Resume Next
      If Not .Recordset Is Nothing Then
         .Recordset.Close
         Set .Recordset = Nothing
      End If
      On Error GoTo 0
      Set .Recordset = ActiveDbConnection.ADODB.OpenRecordset(strSQL, adOpenKeyset, adLockReadOnly, adUseClient)

   End With

End Sub

'Private Sub setData_ProcedureFilterArray(ByRef FilterParamDefs As Collection, _
'                                         ByVal sOrderByParams As String)
'
'   Dim strSQL As String
'   Dim cmd As ADODB.Command
'
'   Dim cmdParams() As Variant
'   Dim i As Long
'   Dim fc As Long
'   Dim FilterItem As FormFilterItem
'
'On Error GoTo HandleErr
'
'   fc = FilterParamDefs.Count
'   If fc > 0 Then
'      ReDim cmdParams(fc - 1, 1)
'      i = 0
'      For Each FilterItem In FilterParamDefs
'         With FilterItem
'            If Len(.FilterString) > 0 Then
'               cmdParams(i, 0) = .FieldName
'               cmdParams(i, 0) = .FilterControl.Value
'            End If
'         End With
'         i = i + 1
'      Next
'   End If
'
'   With ActiveForm
'
'      On Error Resume Next
'      If Not .Recordset Is Nothing Then
'         .Recordset.Close
'         Set .Recordset = Nothing
'      End If
'      On Error GoTo HandleErr
'      Set .Recordset = ActiveDbConnection.ADODB.OpenRecordsetCommandParam(m_DataSource, ADODB.CommandTypeEnum.adCmdStoredProc, cmdParams)
'
'   End With
'
'ExitHere:
'   Exit Sub
'
'HandleErr:
'   Select Case HandleFormHandlerError(Err.Number, "setData_ProcedureFilterArray : " & Err.Source, Err.Description)
'   Case AcUtilErrorResumeMode.acutilErrResume
'      Resume
'   Case AcUtilErrorResumeMode.acutilErrResumeNext
'      Resume Next
'   Case Else
'      On Error Resume Next
'      Resume ExitHere
'   End Select
'
'End Sub

Private Sub SetData_ProcedurePT(ByVal sFilterParams As String, _
                           ByVal sOrderByParams As String)

   Dim strSQL As String
   Dim qdf As DAO.QueryDef
   Dim strQdfName As String
   
   strSQL = "exec " & m_DataSource
   If Len(sFilterParams) > 0 Then
      strSQL = strSQL & " " & sFilterParams
   End If

   If Len(m_OrderParams) > 0 Then
      If Len(sFilterParams) > 0 Then
         strSQL = strSQL & ","
      End If
      strSQL = strSQL & " " & sOrderByParams
   End If

   If Len(m_DataQueryDefName) > 0 Then
      strQdfName = m_DataQueryDefName
   Else
      strQdfName = "q_frm" & ActiveForm.Name
   End If
   Set qdf = ActiveDbConnection.CurrentDb.QueryDefs(strQdfName)
   qdf.SQL = strSQL
   qdf.Close
   Set qdf = Nothing

   If Len(m_BindingDataSourceName) > 0 Then
      strQdfName = m_BindingDataSourceName
   End If

   With ActiveForm
      .RecordSource = strQdfName
      'Problemumgehung Verbindungsabbruch (Access 2016)
      With .Recordset
         If .RecordCount <> 0 Then
            .MoveLast
            .MoveFirst
         End If
      End With

   End With

End Sub

'---------------------------------------------------------------------------------------
' Sub: SetNoData (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formulardatenbasis entfernen und "NULL-Datensatz" anzeigen
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub SetNoData()

   Select Case m_DataSourceType
      Case FH_DataSourceType.DST_FormSource
         SetData_Form "1=0", vbNullString
      Case FH_DataSourceType.DST_SQL
         SetData_SQL "1=0", vbNullString
      Case FH_DataSourceType.DST_SQLPTdao, FH_DataSourceType.DST_SQLPTado
         SetNoData_ChangeDataSource
      Case FH_DataSourceType.DST_DAORS
         SetData_DaoRecordset "1=0", vbNullString
      Case FH_DataSourceType.DST_ADORS, FH_DataSourceType.DST_ADORSJET
         SetData_AdoRecordset "1=0", vbNullString
      Case FH_DataSourceType.DST_Procedure
         SetNoData_ChangeDataSource
      Case FH_DataSourceType.DST_ProcedurePT
         SetNoData_ChangeDataSource
      Case Else
         RaiseEvent FormHandlerMessage("Art der Datenbasis wurde nicht eingestellt!", 0&)
   End Select

End Sub

' Sub: SetNoData (2009-08-10)
Private Sub SetNoData_ChangeDataSource()

   If ActiveForm Is Nothing Then
      Exit Sub
   End If

   With m_AccessForm

      If m_DataSourceType = FH_DataSourceType.DST_Procedure Then
         On Error Resume Next
         If Not (.Recordset Is Nothing) Then
            .Recordset.Close
            Set .Recordset = Nothing
         End If
         On Error GoTo 0
         If Len(m_DataSourceNoData) > 0 Then
            Set .Recordset = ActiveDbConnection.DAO.OpenRecordset(m_DataSourceNoData, dbOpenDynaset)
         Else
            Set .Recordset = Nothing
         End If
      Else
         .RecordSource = m_DataSourceNoData
      End If

   End With


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
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let DefaultOrderParams(sOrderParams As String)
   m_DefaultOrderParams = sOrderParams
End Property


'---------------------------------------------------------------------------------------
' Property: ActiveForm (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' verwendetes Access-Form-Objekt
' </summary>
' <returns>Access.Form</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ActiveForm() As Access.Form

On Error Resume Next
   Dim acForm As Object

   If m_AccessForm Is Nothing Then
      RaiseEvent FormHandlerErrorMissingForm("FormHandler.Form", 0&, acForm)
      Set Me.ActiveForm = acForm
   End If
   Set ActiveForm = m_AccessForm

End Property

'---------------------------------------------------------------------------------------
' Property: ActiveForm (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Access-Form-Objekt welches ListFormHandler steuern soll
' </summary>
' <param name="FormRef">zu übergebendes Access-Form</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set ActiveForm(ByRef FormRef As Access.Form)

   If Not m_AccessForm Is Nothing Then
      Set m_AccessForm = Nothing
   End If
   Set m_AccessForm = FormRef

   If Not (m_AccessForm Is Nothing) Then

      With m_AccessForm

         'Access-Application auslesen
         Set m_AccessApplication = .Application

         'auf Unload-Ereignis reagieren, damit Objektreferenzen entladen werden können
         If Len(.OnUnload) = 0 Then .OnUnload = g_conEventProcedurePropertyTag
         'selbes auch für Close
         If Len(.OnClose) = 0 Then .OnClose = g_conEventProcedurePropertyTag

      End With
   Else
      ClearFormObjectReferences
   End If

End Property

'Private Sub m_AccessForm_Unload(Cancel As Integer)
'   RaiseEvent AccessFormUnload(Cancel)
'   If Cancel = True Then
'      Exit Sub
'   End If
'   clearFormObjectReferences
'End Sub

'---------------------------------------------------------------------------------------
' Property: DataSourceUniqueFieldName (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Feldbezeichnung der Datenherkunft, die eindeutige Datensatzkennung zurückgibt
' </summary>
' <returns>String</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DataSourceUniqueFieldName() As String
   DataSourceUniqueFieldName = m_strDataSourceUniqueFieldName
End Property

'---------------------------------------------------------------------------------------
' Property: DataSourceUniqueFieldName (15.09.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Feldbezeichnung der Datenherkunft, die eindeutige Datensatzkennung zurückgibt
' </summary>
' <param name="FieldName">Feldbezeichnung der Datenherkunft</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let DataSourceUniqueFieldName(ByVal FieldName As String)
   m_strDataSourceUniqueFieldName = FieldName
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
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get HandledControls(ByVal Index As FH_HandledControls) As Access.Control

   Select Case Index
      Case FH_HandledControls.FHC_CloseCommandButton
         Set HandledControls = closeCommandButton
      Case FH_HandledControls.FHC_EditModeCommandButton
         Set HandledControls = EditModeCommandButton
      Case FH_HandledControls.FHC_EditModeInfoLabel
         Set HandledControls = EditModeInfoLabel
      Case FH_HandledControls.FHC_SelectionFirstControl
         Set HandledControls = SelectionFirstControl
      Case Else
         RaiseEvent FormHandlerError("Index-Wert ist nicht verfügbar", -1, 0)
         HandledControls = Nothing
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
' <returns>Access.Control</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set HandledControls(ByVal Index As FH_HandledControls, ByRef FormControl As Access.Control)

   Select Case Index
      Case FH_HandledControls.FHC_CloseCommandButton
         Set closeCommandButton = FormControl
      Case FH_HandledControls.FHC_EditModeCommandButton
         Set EditModeCommandButton = FormControl
      Case FH_HandledControls.FHC_EditModeInfoLabel
         Set EditModeInfoLabel = FormControl
      Case FH_HandledControls.FHC_RequeryCommandButton
         Set RequeryCommandButton = FormControl
      Case FH_HandledControls.FHC_SelectionFirstControl
         Set SelectionFirstControl = FormControl
      Case Else
         RaiseEvent FormHandlerError("Index-Wert ist nicht verfügbar", -1, 0)
   End Select

End Property

'-------------------------------------------------------------------------------------------
' CommandButton: Close
'
Private Property Get closeCommandButton() As Access.CommandButton

   Set closeCommandButton = m_CloseCommandButton

End Property

Private Property Set closeCommandButton(ByRef newCloseCommandButton As Access.CommandButton)

   If Not m_CloseCommandButton Is Nothing Then
      Set m_CloseCommandButton = Nothing
   End If
   Set m_CloseCommandButton = newCloseCommandButton

   If Not (m_CloseCommandButton Is Nothing) Then _
   If Len(m_CloseCommandButton.OnClick) = 0 Then m_CloseCommandButton.OnClick = g_conEventProcedurePropertyTag

End Property

Private Sub m_AccessForm_Unload(Cancel As Integer)
   RaiseEvent AccessFormUnload(Cancel)
End Sub

Private Sub m_CloseCommandButton_Click()
   CloseForm
End Sub


'-------------------------------------------------------------------------------------------
' CommandButton: EditMode
'
Private Property Get EditModeCommandButton() As Access.CommandButton

   Set EditModeCommandButton = m_EditModeCommandButton

End Property


Private Property Set EditModeCommandButton(ByRef newCloseCommandButton As Access.CommandButton)

   If Not m_EditModeCommandButton Is Nothing Then
      Set m_EditModeCommandButton = Nothing
   End If
   Set m_EditModeCommandButton = newCloseCommandButton

   If Not (m_EditModeCommandButton Is Nothing) Then _
   If Len(m_EditModeCommandButton.OnClick) = 0 Then m_EditModeCommandButton.OnClick = g_conEventProcedurePropertyTag

End Property

Private Sub m_EditModeCommandButton_Click()
   ToggleEditMode
End Sub


'-------------------------------------------------------------------------------------------
' Label: EditModeInfoLabel
'
Private Property Get EditModeInfoLabel() As Access.Label

   Set EditModeInfoLabel = m_EditModeLabel

End Property

Private Property Set EditModeInfoLabel(ByRef newEditModeInfoLabel As Access.Label)

   If Not m_EditModeLabel Is Nothing Then
      Set m_EditModeLabel = Nothing
   End If
   Set m_EditModeLabel = newEditModeInfoLabel

End Property

'-------------------------------------------------------------------------------------------
' Control: Select First
'
Private Property Get SelectionFirstControl() As Access.Control

   Set SelectionFirstControl = m_SelectFirstFormControl

End Property

Private Property Set SelectionFirstControl(ByRef FormControl As Access.Control)

   If Not m_SelectFirstFormControl Is Nothing Then
      Set m_SelectFirstFormControl = Nothing
   End If
   Set m_SelectFirstFormControl = FormControl

End Property

'-------------------------------------------------------------------------------------------
' CommandButton: Requery
'
Private Property Get RequeryCommandButton() As Access.Control

   Set RequeryCommandButton = m_RequeryCommandButton

End Property

Private Property Set RequeryCommandButton(ByRef FormControl As Access.Control)

   If Not m_RequeryCommandButton Is Nothing Then
      Set m_RequeryCommandButton = Nothing
   End If
   Set m_RequeryCommandButton = FormControl

   If Not (m_RequeryCommandButton Is Nothing) Then _
   If Len(m_RequeryCommandButton.OnClick) = 0 Then m_RequeryCommandButton.OnClick = g_conEventProcedurePropertyTag

End Property

Private Sub m_RequeryCommandButton_Click()
   RequeryData
End Sub


'-----------------------------------------------------------------------------------------------------------
'
' Hilfsfunktionen
'
Private Function GotoBookmark(ByVal BookmarkID_Name As String, ByVal BookmarkID_Value As Variant)

   Dim rst As Object
   Dim strCriteria As String

   If IsNumeric(BookmarkID_Value) Then
      strCriteria = BookmarkID_Name & "=" & BookmarkID_Value
   Else
      strCriteria = BookmarkID_Name & "='" & BookmarkID_Value & "'"
   End If

   Set rst = m_AccessForm.Recordset.Clone 'Recordset wegen Ac2000, da hier bei ADO kein Recordsetclone möglich ist! => 2008-01-26: .clone
   If TypeOf rst Is DAO.Recordset Then
      With rst
         .FindFirst strCriteria
         If .NoMatch = False Then
            m_AccessForm.Bookmark = .Bookmark
         End If
      End With
   ElseIf TypeOf rst Is ADODB.Recordset Then
      With rst
         .Find strCriteria, , adSearchForward
         If Not .EOF Then
            m_AccessForm.Bookmark = .Bookmark
         End If
      End With
   End If
   Set rst = Nothing

End Function

Private Sub SetAccessApplicationHourglass(ByVal bHourGlass As Boolean)

   On Error Resume Next
   DoCmd.Hourglass bHourGlass

End Sub

Public Property Get EditMode() As FH_EditMode
   EditMode = m_CurrentMode
End Property

Public Property Let EditMode(ByVal NewMode As FH_EditMode)
   ChangeEditMode NewMode
End Property

Public Sub ToggleEditMode()

   Dim bolFormPaintingMode As Boolean

   bolFormPaintingMode = m_AccessForm.Painting
   
   m_AccessForm.Painting = False

   If m_AccessForm.Dirty Then
     m_AccessForm.Dirty = False
   End If

   If m_AccessForm.AllowEdits Then
     ChangeEditMode EM_ReadOnly
   Else
     ChangeEditMode EM_Edit
   End If
   
   m_AccessForm.Painting = bolFormPaintingMode

End Sub

Public Sub ChangeEditMode(ByVal NewMode As FH_EditMode)

   With m_AccessForm
   Select Case NewMode
      Case FH_EditMode.EM_ReadOnly
         .AllowEdits = False
         .AllowAdditions = False
         If Not (m_EditModeCommandButton Is Nothing) Then m_EditModeCommandButton.Enabled = True
         m_CurrentMode = FH_EditMode.EM_ReadOnly
      Case FH_EditMode.EM_Edit
         .AllowEdits = True
         .AllowAdditions = False
         If Not (m_EditModeCommandButton Is Nothing) Then m_EditModeCommandButton.Enabled = True
         m_CurrentMode = FH_EditMode.EM_Edit
      Case FH_EditMode.EM_New
         .AllowEdits = True
         .AllowAdditions = True
         If Not (m_EditModeCommandButton Is Nothing) Then m_EditModeCommandButton.Enabled = True
         m_CurrentMode = FH_EditMode.EM_New
      Case FH_EditMode.EM_Locked
         .AllowEdits = False
         .AllowAdditions = False
         If Not (m_EditModeCommandButton Is Nothing) Then m_EditModeCommandButton.Enabled = False
         m_CurrentMode = FH_EditMode.EM_Locked
      Case Else '=0
         .AllowEdits = False
         .AllowAdditions = False
         If Not (m_EditModeCommandButton Is Nothing) Then m_EditModeCommandButton.Enabled = True
         m_CurrentMode = FH_EditMode.EM_ReadOnly
   End Select
   End With

   RaiseEvent EditModeChanged(m_CurrentMode)

   SetEditModeInfoLabel m_CurrentMode

End Sub

Private Sub SetEditModeInfoLabel(EditMode As FH_EditMode)

   If m_EditModeLabel Is Nothing Then
      Exit Sub
   End If

   With m_EditModeLabel
   Select Case EditMode
      Case FH_EditMode.EM_ReadOnly
         .Caption = "Read Only"
      Case FH_EditMode.EM_Edit
         .Caption = "Edit"
      Case FH_EditMode.EM_New
         .Caption = "New"
      Case FH_EditMode.EM_Locked
         .Caption = "Read Only"
   End Select
   End With

End Sub

Private Property Get IsListForm() As Boolean

   Dim bolIsListForm As Boolean

   With Me.ActiveForm
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

End Property

Private Sub m_AccessForm_Close()
   Dispose
End Sub

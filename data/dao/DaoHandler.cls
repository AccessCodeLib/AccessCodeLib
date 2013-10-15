VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: DaoHandler
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' DAO-Zugriffsmethoden
' </summary>
' <remarks></remarks>
'\ingroup data_dao
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/DaoHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <test>_test/data/dao/DaoHandlerTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)

'DAO-Database
Private m_DaoDb As DAO.Database

'Standardwerte für optionale DAO-Enum-Paramter z. B. in OpenRecordset
' (falls diese nicht gesetzt sind, gilt DAO-Standard)
Private m_RecordsetTypeEnumDefault As DAO.RecordsetTypeEnum
Private m_RecordsetOptionEnumDefault As DAO.RecordsetOptionEnum
Private m_LockTypeEnumDefault As DAO.LockTypeEnum

'Error-Events für mögliche Fehlerbehandlung über die FE-Anwendung
'/**
' <summary>
' Ereignis für fehlende CurrentDb-Referenz
' </summary>
' <param name="Msg">Meldung über Ursache der Ereignisauslösung. (Wird noch nicht genutzt)</param>
' <param name="NewCurrentDbRef">DAO.Database-Referenz, die für CurrentDb verwendet werden soll</param>
' <remarks>über den ByRef-Parameter NewCurrentDbRef kann die passende DAO.Database-Referenz weitergegeben werden.</remarks>
'**/
Public Event ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)

'
' Init / Terminate
' ----------------
'
'Private Sub Class_Initialize()
''
'End Sub

Private Sub Class_Terminate()
   Dispose
End Sub

Public Sub Dispose()
   Set m_DaoDb = Nothing
   m_RecordsetTypeEnumDefault = 0
   m_RecordsetOptionEnumDefault = 0
   m_LockTypeEnumDefault = 0
End Sub

'---------------------------------------------------------------------------------------
' Sub: InitRecordsetEnumDefaultValues
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Standardwerte für Recordset-Enums einstellen, welche verwendet werden, wenn Parameter nicht gesetzt werden können (z. B. in OpenRecordsetParamSQL2)
' </summary>
' <param name="RecordsetType">Standardwert aus DAO.RecordsetTypeEnum</param>
' <param name="Options">Standardwert aus DAO.RecordsetOptionEnum</param>
' <param name="LockEdit">Standardwert aus DAO.LockTypeEnum</param>
' <remarks>Wenn 0 eingestellt wird, kommt Standard von DAO zum Einsatz (0-Parameter werden als "IsMissing"-Variant-Wert an DAO weitergeben)</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub InitRecordsetEnumDefaultValues( _
                        Optional ByRef RecordsetType As DAO.RecordsetTypeEnum, _
                        Optional ByRef Options As DAO.RecordsetOptionEnum, _
                        Optional ByRef LockEdit As DAO.LockTypeEnum)

   m_RecordsetTypeEnumDefault = RecordsetType
   m_RecordsetOptionEnumDefault = Options
   m_LockTypeEnumDefault = LockEdit

End Sub

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'/**
' <summary>
' CurrentDbC-Variante
' </summary>
' <returns>DAO.Database</returns>
' <remarks>
' CurrentDbC-Variante, die zum Unterschied zur Original-Version von Michael Kaplan
' die DAO-Database-Referenz nicht automatisch auf Access.Application.CurrentDb setzt,
' sondern über ein Ereignis anfordert, falls noch keine Referenz vorhanden ist.
' Dadurch kann über die Hauptanwendung eine belibige Database-Referenz übergeben werden.
' (z. B. falls eine Temp-Datenbank als CurrentDb dienen soll)</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentDb() As DAO.Database
    If (m_DaoDb Is Nothing) Then
        'Error-Event auslösen und hoffen, dass neue Referenz geliefert wird
        RaiseEvent ErrorMissingCurrentDb("DaoHandler.CurrentDb: unavailable CurrentDb-Database", m_DaoDb)
        If (m_DaoDb Is Nothing) Then
        'CurrentDb der Anwendung verwenden, falls DaoHandler "selbstinstanzierend" (VB_PredeclaredId = True) verwendet wird
        ' Achtung! ... diese Prüfung hat eventuell den Nachteil,
        '              dass ab hier VB_PredeclaredId wirksam wird,
        '              obwohl es möglicherweise nicht benötigt wird.
        If Me Is DaoHandler Then 'trifft nur zu, wenn Instanz identisch mit der VB_PredeclaredId-Instanz ist
        '                         Jede andere Instanz von DaoHandler hat andere Speicheradresse (ObjPtr(Me) <> ObjPtr(DaoHandler)).
           Set Me.CurrentDb = Application.CurrentDb
        End If
        End If
    End If
    Set CurrentDb = m_DaoDb
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Übergabe der DAO-Database-Referenz
' </summary>
' <param name="NewCurrentDb">Database-Referenz, die als "CurrentDb" verwendet werden soll</param>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set CurrentDb(ByRef NewCurrentDb As DAO.Database)
    Set m_DaoDb = NewCurrentDb
End Property

'---------------------------------------------------------------------------------------
' Function: Execute
'---------------------------------------------------------------------------------------
'/**
' <summary>
' SQL-Anweisung ausführen
' </summary>
' <param name="CommandText">SQL-Anweisung</param>
' <param name="Options"></param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal Query As String, _
               Optional ByVal Options As DAO.RecordsetOptionEnum) As Long

    If (Options And dbSeeChanges) = 0 Then
        Options = Options + dbSeeChanges
    End If
    Me.CurrentDb.Execute Query, Options
    Execute = Me.CurrentDb.RecordsAffected
 
End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteParamSQL
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Sql-Anweisung mit Parametern ausführen
' </summary>
' <param name="SqlText">SQL-Anweisung</param>
' <param name="ParamValues">Parmeterwerte in passender Reihenfolge</param>
' <param name="QueryParams">
'  Werte in passender Reihenfolge als ParamArray oder als Array <br/>
'  oder<br/>
'  2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert) ... erzeugbar mit GetParamDefArray oder GetNamedParamDefArray
' </param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ExecuteParamSQL(ByVal SqlText As String, _
                           ParamArray QueryParams() As Variant) As Long
 
   Dim qdf As DAO.QueryDef

On Error GoTo HandleErr

   If IsMissing(QueryParams) Then
      Set qdf = Me.ParamQueryDefSql(SqlText)
   Else
      Set qdf = Me.ParamQueryDefSql(SqlText, QueryParams)
   End If
   
   ExecuteParamSQL = ExecuteQueryDef(qdf)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing
   
ExitHere:
   Exit Function
   
HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "ExecuteParamSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
 
End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteQueryDefByName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gespeicherte Abfrage (optional mit Parametern) ausführen
' </summary>
' <param name="QueryName">Name der gespeicherten Abfrage (QueryDef)</param>
' <param name="QueryParams">
'  Werte in passender Reihenfolge als ParamArray oder als Array <br/>
'  oder<br/>
'  2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert) ... erzeugbar mit GetParamDefArray oder GetNamedParamDefArray
' </param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ExecuteQueryDefByName(ByVal QueryName As String, _
                                 ParamArray QueryParams() As Variant) As Long
   
   Dim qdf As DAO.QueryDef
 
On Error GoTo HandleErr
 
   If IsMissing(QueryParams) Then
      Set qdf = Me.ParamQueryDefByName(QueryName)
   Else
      Set qdf = Me.ParamQueryDefByName(QueryName, QueryParams)
   End If

   ExecuteQueryDefByName = ExecuteQueryDef(qdf)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing
   
ExitHere:
   Exit Function
   
HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "ExecuteParamQdfByName:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
 
End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteQueryDef
'---------------------------------------------------------------------------------------
'/**
' <summary>
' QueryDef-Objekt mit Parameterwerten befüllen und anschließend ausführen
' </summary>
' <param name="Qdf">QueryDef-Referenz</param>
' <param name="QueryParams">
'  Werte in passender Reihenfolge als ParamArray oder als Array <br/>
'  oder<br/>
'  2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert) ... erzeugbar mit GetParamDefArray oder GetNamedParamDefArray
' </param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ExecuteQueryDef(ByVal qdf As DAO.QueryDef, _
                       ParamArray QueryParams() As Variant) As Long
 
   Dim QdfParamDefs() As Variant
   Dim ParamDefsArraySize As Long
   Dim ParamIndex As Variant
   Dim i As Long
   
   If Not IsMissing(QueryParams) Then
      QdfParamDefs = GetParamDefArray(QueryParams)
      FillQueryDefParameters qdf, QdfParamDefs
   End If
   
   qdf.Execute dbFailOnError
   
   ExecuteQueryDef = qdf.RecordsAffected

End Function

Private Sub FillQueryDefParameters(ByVal qdf As DAO.QueryDef, ByVal QdfParamDefs As Variant)
   
   Dim ParamDefsArraySize As Long
   Dim ParamIndex As Variant
   Dim i As Long

   ParamDefsArraySize = UBound(QdfParamDefs, 1)
   
   With qdf
      For i = 0 To ParamDefsArraySize
         ParamIndex = QdfParamDefs(i, 0)
         If Len(ParamIndex) > 0 Then
            .Parameters(ParamIndex).Value = QdfParamDefs(i, 1)
         End If
      Next i
   End With
   
End Sub

'---------------------------------------------------------------------------------------
' Function: ParamQueryDefByName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gespeicherte Abfrage als QueryDef-Objekt ink. übergebenen Parameterwerten öffnen
' </summary>
' <param name="QueryName">QueryDef-Name</param>
' <param name="QueryParams">
'  Werte in passender Reihenfolge als ParamArray oder als Array <br/>
'  oder<br/>
'  2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert) ... erzeugbar mit GetParamDefArray oder GetNamedParamDefArray
' </param>
' <returns>DAO.QueryDef</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ParamQueryDefByName(ByVal QueryName As String, ParamArray QueryParams() As Variant) As DAO.QueryDef

   Dim qdf As DAO.QueryDef
   Dim QdfParamDefs() As Variant
 
   Set qdf = Me.CurrentDb.QueryDefs(QueryName)
   If Not IsMissing(QueryParams) Then
      QdfParamDefs = GetParamDefArray(QueryParams)
      FillQueryDefParameters qdf, QdfParamDefs
   End If
   
   Set ParamQueryDefByName = qdf

End Function

'---------------------------------------------------------------------------------------
' Function: ParamQueryDefSql
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Temporäres QueryDef-Objekt ink. übergebenen Parameterwerten öffnen
' </summary>
' <param name="ParamSqlText">SQL-Anweisung für das temporäre QueryDef-Objekt</param>
' <param name="QueryParams">
'  Werte in passender Reihenfolge als ParamArray oder als Array <br/>
'  oder<br/>
'  2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert) ... erzeugbar mit GetParamDefArray oder GetNamedParamDefArray
' </param>
' <returns>DAO.QueryDef</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ParamQueryDefSql(ByVal ParamSqlText As String, ParamArray QueryParams() As Variant) As DAO.QueryDef

   Dim qdf As DAO.QueryDef
   Dim QdfParamDefs() As Variant

   Set qdf = Me.CurrentDb.CreateQueryDef("", ParamSqlText)
   If Not IsMissing(QueryParams) Then
      QdfParamDefs = GetParamDefArray(QueryParams)
      FillQueryDefParameters qdf, QdfParamDefs
   End If
   
   Set ParamQueryDefSql = qdf

End Function


'---------------------------------------------------------------------------------------
' Function: OpenRecordset
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Recordset öffnen
' </summary>
' <param name="Source">SQL-Anweisung oder Tabllen- bzw. Abfragename</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordset(ByVal Source As String, _
                     Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenDynaset, _
                     Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges, _
                     Optional ByVal LockEdit As DAO.LockTypeEnum) As DAO.Recordset

   If (RecordsetOptions And dbSeeChanges) = 0 Then
      RecordsetOptions = RecordsetOptions + dbSeeChanges
   End If
   
   ' Replace0WithIsMissing ... damit IsMissing für Variant-Datentyp an Original-DAO-Openrecordset wirkt
   Set OpenRecordset = Me.CurrentDb.OpenRecordset(Source, _
                                                  Replace0WithIsMissing(RecordsetType), _
                                                  Replace0WithIsMissing(RecordsetOptions), _
                                                  Replace0WithIsMissing(LockEdit))

End Function

Private Function Replace0WithIsMissing(ByVal CheckValue As Long) As Variant
   If CheckValue = 0 Then
      Replace0WithIsMissing = GetIsMissing()
   Else
      Replace0WithIsMissing = CheckValue
   End If
End Function

Private Function GetIsMissing(Optional DoNotFillThisParam As Variant) As Variant
   GetIsMissing = DoNotFillThisParam
End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetParamSQL
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Parameter-SQL-Anweisung öffnen (verwendet temporäres QueryDef)
' </summary>
' <param name="ParamSqlText">SQL-Anweisung mit Parameter (Parameters-Deklaration)</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <param name="QueryParams">
'  Werte in passender Reihenfolge als ParamArray oder als Array <br/>
'  oder<br/>
'  2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert) ... erzeugbar mit GetParamDefArray oder GetNamedParamDefArray
' </param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetParamSQL(ByVal ParamSqlText As String, _
                     ByVal RecordsetType As DAO.RecordsetTypeEnum, _
                     ByVal Options As DAO.RecordsetOptionEnum, _
                     ByVal LockEdit As DAO.LockTypeEnum, _
                     ParamArray QueryParams() As Variant) As DAO.Recordset

   Dim qdf As DAO.QueryDef
   Dim QdfParamDefs As Variant '
   
   QdfParamDefs = GetParamDefArray(QueryParams)
   
On Error GoTo HandleErr
   
   Set qdf = Me.CurrentDb.CreateQueryDef("", ParamSqlText)
   Set OpenRecordsetParamSQL = OpenRecordsetQueryDef(qdf, QdfParamDefs, RecordsetType, Options, LockEdit)
   
On Error GoTo 0

   qdf.Close
   Set qdf = Nothing
   
ExitHere:
   Exit Function
   
HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "OpenRecordsetParamSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
 
End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetParamSQL2
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Parameter-SQL-Anweisung öffnen (wie OpenRecordsetParamSQL, nur mit Standardwerten für RecordsetType, RecordsetOptions u. LockEdit)
' </summary>
' <param name="ParamSqlText">SQL-Anweisung mit Parameter (Parameters-Deklaration)</param>
' <param name="QueryParams">
'  Werte in passender Reihenfolge als ParamArray oder als Array <br/>
'  oder<br/>
'  2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert) ... erzeugbar mit GetParamDefArray oder GetNamedParamDefArray
' </param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetParamSQL2(ByVal ParamSqlText As String, _
                                       ParamArray QueryParams() As Variant) As DAO.Recordset

   Set OpenRecordsetParamSQL2 = OpenRecordsetParamSQL(ParamSqlText, _
                                                      m_RecordsetTypeEnumDefault, m_RecordsetOptionEnumDefault, m_LockTypeEnumDefault, _
                                                      QueryParams)

End Function

'---------------------------------------------------------------------------------------
' Function: GetParamDefArray
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erezugt 2-dimensionalen Parameter-Array (z. B. für OpenRecordsetQueryDef)
' </summary>
' <param name="ParamValues">Parmeterwerte in passender Reihenfolge (ohne Index-Kennung!)</param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetParamDefArray(ParamArray ParamValues() As Variant) As Variant
   
   Dim i As Long
   Dim QdfParamDefs() As Variant
   Dim ArraySize As Long
   Dim ParamValueArray As Variant  ' !!! nicht als Array deklarieren, da sonst ParamValueArray = ParamValueArray(0) fehltschlägt, wenn ParamValueArray(0) ein Array ist
   Dim ArrayDimCnt As Long, SubArrayDimCnt As Long
   
   ArraySize = UBound(ParamValues)
   If ArraySize = -1 Then
      GetParamDefArray = QdfParamDefs 'nicht dimensionierten Array zurückgeben
      Exit Function
   End If
   
   ParamValueArray = ParamValues
   ArrayDimCnt = ArrayDimensionCount(ParamValueArray)
   If ArrayDimCnt = 2 Then
      GetParamDefArray = ParamValues
      Exit Function
   End If
   
   SubArrayDimCnt = ArrayDimensionCount(ParamValueArray(0))
   If SubArrayDimCnt <> -1 Then
   
      Do While SubArrayDimCnt <> -1
      
         ParamValueArray = ParamValueArray(0)
      
         If SubArrayDimCnt = 2 Then  ' !!! /** \todo CHECK SubArrayDimCnt = 2176 .. ist das immer richtig? **/
         'Bereits aufbereiteter 2-dim. Array
                  GetParamDefArray = ParamValueArray
                  Exit Function
         End If
         
         SubArrayDimCnt = ArrayDimensionCount(ParamValueArray(0))
         
      Loop
      
      ArraySize = UBound(ParamValueArray)
      
   End If
   
   ReDim QdfParamDefs(ArraySize, 1)

   For i = 0 To ArraySize
      QdfParamDefs(i, 0) = i
      QdfParamDefs(i, 1) = ParamValueArray(i)
   Next
   
   GetParamDefArray = QdfParamDefs
   
End Function

Private Function ArrayDimensionCount(ByRef ArrayToCheck As Variant) As Integer
' siehe http://msdn.microsoft.com/de-de/library/bb978902.aspx .. Funktioniert nicht, wenn Array in Variant eingebettet ist

'   Dim lngCnt As Integer
'   Dim lpSADescriptor As Long, lpSafeArray As Long
'
'On Error Resume Next
'
'   If Not IsArray(ArrayToCheck) Then
'      ArrayDimensionCount = -1
'      Exit Function
'   End If
'
'   ' Adresse des SAFEARRAY-Deskriptors holen
'   CopyMemory lpSADescriptor, ByVal VarPtr(ArrayToCheck) + 8, 4
'   ' Adresse der SAFEARRAY-Struktur holen
'   CopyMemory lpSafeArray, ByVal lpSADescriptor, 4
'
'   If lpSafeArray <> 0 Then
'      ' Die ersten zwei Bytes enthalten die Anzahl der Dimensionen:
'      CopyMemory lngCnt, ByVal lpSafeArray, 2
'   End If
'
'   ArrayDimensionCount = lngCnt
   
   Dim ArraySize As Long
   Dim DimCnt As Long
   
   DimCnt = 0
   
   If Not IsArray(ArrayToCheck) Then
      ArrayDimensionCount = -1
      Exit Function
   End If
   
On Error Resume Next

   Do
      DimCnt = DimCnt + 1
      ArraySize = UBound(ArrayToCheck, DimCnt)
   Loop While Err.Number = 0
   Err.Clear
   
On Error GoTo 0
   
   DimCnt = DimCnt - 1
   
   ArrayDimensionCount = DimCnt
   
End Function

'---------------------------------------------------------------------------------------
' Function: GetNamedParamDefArray
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt 2-dimensionalen Parameter-Array (z. B. für OpenRecordsetQueryDef)
' </summary>
' <param name="ParamValues">Parameter-Paare: abwechselnd Parameter-Index und Parameter-Wert</param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetNamedParamDefArray(ParamArray ParamIndexAndValues() As Variant) As Variant
   
   Dim i As Long
   Dim QdfParamDefs() As Variant
   Dim ArraySize As Long
   
   ArraySize = UBound(ParamIndexAndValues)
   If ArraySize = -1 Then
      GetNamedParamDefArray = QdfParamDefs 'nicht dimensionierten Array zurückgeben
      Exit Function
   End If
   
   'Prüfen, ob immer Paare (Index / Wert) möglich sind
   If (((ArraySize + 1) \ 2) * 2) <> (ArraySize + 1) Then
      Err.Raise vbObjectError, "DaoHandler.GetNamedParamDefArray", "ParamArray muss immer Paare für Parameter-Index und Parameter-Wert enthalten."
   End If
   
   ArraySize = ArraySize \ 2
   
   ReDim QdfParamDefs(ArraySize, 1)

   For i = 0 To ArraySize
      QdfParamDefs(i, 0) = ParamIndexAndValues(i * 2)
      QdfParamDefs(i, 1) = ParamIndexAndValues(i * 2 + 1)
   Next
   
   GetNamedParamDefArray = QdfParamDefs
   
End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetQueryDefByName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Recordset von gespeicherter Abfrage (optional mit Parametern) öffnen
' </summary>
' <param name="QueryName">Name der gespeicherten Abfrage (QueryDef)</param>
' <param name="QueryParams">2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert)</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetQueryDefByName(ByVal QueryName As String, _
                             Optional ByVal QueryParams As Variant, _
                             Optional ByVal RecordsetType As DAO.RecordsetTypeEnum, _
                             Optional ByVal Options As DAO.RecordsetOptionEnum, _
                             Optional ByVal LockEdit As DAO.LockTypeEnum) As DAO.Recordset

   Dim qdf As DAO.QueryDef
   Dim i As Long
 
On Error GoTo HandleErr
 
   Set qdf = Me.CurrentDb.QueryDefs(QueryName)
   Set OpenRecordsetQueryDefByName = OpenRecordsetQueryDef(qdf, QueryParams, RecordsetType, Options, LockEdit)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing
   
ExitHere:
   Exit Function
   
HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "OpenQueryDefRecordsetByName:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
 
End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetQueryDef
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Recordset von gespeicherter Abfrage (optional mit Parametern) öffnen
' </summary>
' <param name="Qdf">QueryDef-Referenz</param>
' <param name="QueryParams">2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert)</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetQueryDef(ByVal qdf As DAO.QueryDef, _
                             Optional ByVal QdfParamDefs As Variant, _
                             Optional ByVal RecordsetType As DAO.RecordsetTypeEnum, _
                             Optional ByVal Options As DAO.RecordsetOptionEnum, _
                             Optional ByVal LockEdit As DAO.LockTypeEnum) As DAO.Recordset
 
   Dim ParamDefsArraySize As Long
   Dim ParamIndex As Variant
   Dim i As Long
   
   If Not IsMissing(QdfParamDefs) Then
      FillQueryDefParameters qdf, QdfParamDefs
   End If

   Set OpenRecordsetQueryDef = qdf.OpenRecordset(Replace0WithIsMissing(RecordsetType), _
                                                 Replace0WithIsMissing(Options), _
                                                 Replace0WithIsMissing(LockEdit))

End Function

'---------------------------------------------------------------------------------------
' Function: LookupSQL
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion für die Übergabe einer vollständigen SQL-Anweisung
' </summary>
' <param name="SqlText">SQL-Anweisung</param>
' <param name="Index">Feldauswahl, falls nicht das erste Feld des Recordsets zurückgegeben werden soll</param>
' <param name="ValueIfNull">Rückgabewert bei fehlendem Datensatz oder NULL-Feldwert (Standard: Null</param>
' <returns>Variant</returns>
' <remarks>Die SQL-Anweisung ist im Jet-SQL-Dialekt zu schreiben.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LookupSQL(ByVal SqlText As String, _
                 Optional ByVal Index As Variant = 0&, _
                 Optional ByVal ValueIfNull As Variant = Null) As Variant

    Dim rst As DAO.Recordset

On Error GoTo HandleErr

    Set rst = Me.OpenRecordset(SqlText, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
    With rst
        If .EOF Then
            LookupSQL = ValueIfNull
        Else
            LookupSQL = Nz(.Fields(Index), ValueIfNull)
        End If
        .Close
    End With
    Set rst = Nothing

ExitHere:
   Exit Function
   
HandleErr:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "LookupSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: Lookup
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Lookup(ByVal Expr As String, ByVal Domain As String, _
              Optional ByVal Criteria As Variant, _
              Optional ByVal ValueIfNull As Variant = Null) As Variant

    Dim strSQL As String

    strSQL = "SELECT " & Expr & " FROM (" & Domain & ")"
    If Not (VarType(Criteria) = vbError) Then
        If Len(Criteria) > 0 Then
            strSQL = strSQL & " WHERE " & Criteria
        End If
    End If
    Lookup = LookupSQL(strSQL, , ValueIfNull)

End Function

'---------------------------------------------------------------------------------------
' Function: Count
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DCount-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Long</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Count(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Long
    Count = Nz(Me.Lookup("Count(" & Expr & ")", Domain, Criteria), 0)
End Function

'---------------------------------------------------------------------------------------
' Function: Max
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DMax-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Max(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Max = Me.Lookup("Max(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: DMin
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DMin-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Min(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Min = Me.Lookup("Min(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: DSum
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DSum-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Sum(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Sum = Me.Lookup("Sum(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: InsertIdentityReturn
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Insert-SQL-Anweisung ausführen und letzten Identity-Wert (Autowert) zurückgeben
' </summary>
' <param name="InsertSQL">Insert-SQL-Anweisung (INSERT ...)</param>
' <returns>Identity-Wert oder NULL falls kein Datensatz angefügt wurde</returns>
' <remarks>
' Funktioniert erst ab Jet 4.0 (Access 2000) und ist vom DBMS abhängig. Bei MySQL-Tabellen wird z. B. 0 zurückgegeben.
' (Die <see cref=AdodbHandler#InsertIdentityReturn>ADODB-Variante</see> dieser Methode ist bezüglich Einsatz für aktive DBMS besser geeignet.)
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function InsertIdentityReturn(ByVal InsertSQL As String) As Variant

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim lngRecordsAffected As Long

On Error GoTo HandleErr

    Set db = Me.CurrentDb
    db.Execute InsertSQL
    lngRecordsAffected = db.RecordsAffected
    
    If lngRecordsAffected = 0 Then
        InsertIdentityReturn = Null 'es wurde kein neuer DS angefügt
        Exit Function
    End If

    Set rst = db.OpenRecordset("SELECT @@Identity") 'zuletzt zugewiesener Autowert ... funktioniert nicht mit MySQL-Tabellen!
    With rst
        If .EOF Then
            InsertIdentityReturn = Null
        Else
            InsertIdentityReturn = .Fields(0)
        End If
        .Close
    End With
    Set rst = Nothing
    
ExitHere:
   Exit Function

HandleErr:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "InsertIdentityReturn:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

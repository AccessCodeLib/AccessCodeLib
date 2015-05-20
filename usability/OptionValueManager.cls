VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionValueManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: OptionValueManager
'---------------------------------------------------------------------------------------
'/**
' \author       Andreas Vogt
' <summary>
' Optionswerte laden und bereitstellen
' </summary>
' <remarks></remarks>
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>usability/OptionValueManager.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'

Option Compare Database
Option Explicit

Private myOption As VBA.Collection
Private m_OptionTable As String
Private CancelFlag As Boolean

Public Property Let OptionTable(ByVal NameofTable As String)
    m_OptionTable = NameofTable
    If TableExists Then
        catchOptionValues
    Else
        m_OptionTable = ""
        CancelFlag = True
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Property

Public Property Get Settings(cOption As String) As Variant
    On Error GoTo Settings_Err
    If CancelFlag Then Exit Property

    Settings = myOption.Item(cOption)
    Exit Property

Settings_Err:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Let Settings(cOption As String, vValue As Variant)
    On Error GoTo Settings_Err
    If CancelFlag Then Exit Property

    myOption.Remove cOption
    myOption.Add vValue, cOption
    vValue = Replace(vValue, "'", "''")
    CurrentDb.Execute "Update " & m_OptionTable & " set strValue = '" & vValue & "' Where strKey = '" & cOption & "'"
    Exit Property

Settings_Err:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Private Sub catchOptionValues()
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset(m_OptionTable, dbOpenSnapshot)

    Do While Not rst.EOF
        myOption.Add CStr(rst!strValue), CStr(rst!strKey)
        rst.MoveNext
    Loop

    rst.Close
    Set rst = Nothing
End Sub

Private Function TableExists() As Boolean
    If m_OptionTable = "" Then m_OptionTable = "tabOptionen"
    TableExists = IsNull(DLookup("[Name]", "MSysObjects", "[Name] = '" & m_OptionTable & "' AND (Type = 1 Or Type = 6)")) = False
End Function

Private Sub Class_Initialize()
    Set myOption = New VBA.Collection
    SampleOptionTableCreate
    Me.OptionTable = "tabOptionen"
End Sub

Private Sub Class_Terminate()
    Set myOption = Nothing
End Sub

Public Function SampleOptionTableCreate()
    Dim strSQL As String
    If Not TableExists Then
        strSQL = "Create Table tabOptionen (id AUTOINCREMENT Primary Key, strKey varchar(50), strValue varchar(255))"
        CurrentDb.Execute strSQL
        CurrentDb.Execute "Insert into tabOptionen (strKey, strValue) values ('Sportjahr','2015')"
        CurrentDb.Execute "Insert into tabOptionen (strKey, strValue) values ('SportjahrAlt','2014')"
        CurrentDb.Execute "Insert into tabOptionen (strKey, strValue) values ('Vereinsname','SC Freiburg')"
        CurrentDb.Execute "Insert into tabOptionen (strKey, strValue) values ('Vereinsnummer','4711')"
        CurrentDb.Execute "Insert into tabOptionen (strKey, strValue) values ('PfadHilfedatei','c:\sbsv\help.txt')"
        CurrentDb.Execute "Insert into tabOptionen (strKey, strValue) values ('PfadImportdatei','c:\sbsv\mitgl.dbf')"
        CurrentDb.Execute "Insert into tabOptionen (strKey, strValue) values ('AnwendungGesperrt',0)"
        Application.RefreshDatabaseWindow
    End If
End Function
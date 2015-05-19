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
'Beispiel für Implementierung in Modul oder Formular:
'  Dim myOptions as OptionValueManager
'  Set myOptions = new OptionValueManager
'  myOption.OptionTable = "tabOptionen"
'  myOptions.Settings("Installationsdatum") = now()
'  MsgBox myOptions.Settings("Installationsdatum")
'  Set myOption = Nothing


Option Compare Database
Option Explicit

Private myOption As VBA.Collection
Private m_OptionTable As String
Private CancelFlag As Boolean

Public Property Let OptionTable(ByVal NameofTable As String)
    m_OptionTable = NameofTable
    If tableExists Then
        catchOptionValues
    Else
        m_OptionTable = ""
        CancelFlag = True
        MsgBox "Tabelle für Optionen nicht gefunden!"
    End If
End Property

Public Property Get Settings(cOption As String) As Variant
    On Error GoTo Settings_Err
    If CancelFlag Then Exit Property

    Settings = myOption.Item(cOption)
    Exit Property

Settings_Err:
    If Err.Number = 5 Then
        MsgBox "Name des Optionsfeldes fehlerhaft!", vbInformation + vbOKOnly
    Else
        MsgBox Err.Description
    End If
End Property

Public Property Let Settings(cOption As String, vValue As Variant)
    On Error GoTo Settings_Err
    If CancelFlag Then Exit Property

    myOption.Remove cOption
    myOption.Add vValue, cOption
    CurrentDb.Execute "Update " & m_OptionTable & " set " & cOption & " = " & myOption.Item(cOption)
    Exit Property

Settings_Err:
    If Err.Number = 5 Then
        MsgBox "Name des Optionsfeldes fehlerhaft!", vbInformation + vbOKOnly
    Else
        MsgBox Err.Description
    End If
End Property

Private Sub catchOptionValues()
    Dim i As Long
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset(m_OptionTable, dbOpenSnapshot)

    With myOption
        For i = 0 To rst.Fields.Count - 1
            If Not isPrimaryKey(rst.Fields(i).Name) Then
                .Add DLookup(CStr(rst.Fields(i).Name), m_OptionTable), LCase(CStr(rst.Fields(i).Name))
            End If
        Next i
    End With

    rst.Close
    Set rst = Nothing
End Sub

Private Function isPrimaryKey(cField As String) As Boolean
    Dim tdf As DAO.TableDef
    Dim db As DAO.Database
    Dim idx As DAO.Index

    Set db = CurrentDb
    Set tdf = db.TableDefs(m_OptionTable)

    For Each idx In tdf.Indexes
        If idx.Primary Then
            If idx.Fields(0).Name = cField Then
                isPrimaryKey = True
                Exit For
            End If
        End If
    Next idx

    Set tdf = Nothing
    Set db = Nothing
End Function

Private Function tableExists() As Boolean
    tableExists = IsNull(DLookup("[Name]", "MSysObjects", "[Name] = '" & m_OptionTable & "' AND (Type = 1 Or Type = 6)")) = False
End Function

Private Sub Class_Initialize()
    Set myOption = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Set myOption = Nothing
End Sub

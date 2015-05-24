VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: OptionManager
'---------------------------------------------------------------------------------------
'/**
' \author       Andreas Vogt
' \brief        Klasse zur Verwendung als OptionManager
' \ingroup utils
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>usability/OptionManager.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private myOptions As Variant
Private FieldNameArr As Variant
Private m_OptionTable As String

'***************************************************************************
'************* Initialisierung und Terminierung der Klasse *****************
'***************************************************************************
Private Sub Class_Initialize()
    'nothing to do
End Sub

Private Sub Class_Terminate()
    'nothing to do
End Sub

'***************************************************************************
'************ Public Sub Update zum Generieren der Enum-Werte **************
'***************************************************************************
Public Sub Update()
    CreateEnum
End Sub

'***************************************************************************
'****** Zentrale Property Settings zum Holen und Setzen eines Wertes *******
'***************************************************************************
Public Property Get Settings(cOption As ltOptions) As Variant
    On Error GoTo Settings_Err

    Settings = myOptions(cOption)
    Exit Property

Settings_Err:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Let Settings(cOption As ltOptions, vValue As Variant)
    On Error GoTo Settings_Err

    vValue = Replace(vValue, "'", "''")
    myOptions(cOption) = vValue
    CurrentDb.Execute "Update " & OptionTable & " set strValue = '" & vValue & "' Where strKey = '" & cOption & "'"
    Exit Property

Settings_Err:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

'***************************************************************************
' Property Optiontable zum setzen des Tabellennamens und auslesen der Werte
'***************************************************************************
Public Property Get OptionTable() As String
    OptionTable = m_OptionTable
End Property

Public Property Let OptionTable(ByVal cTable As String)
    m_OptionTable = cTable
    catchOptionValues
End Property

'***************************************************************************
'********************* Prozedur zum Befüllen des Array *********************
'***************************************************************************
Private Sub catchOptionValues()
    Dim i As Long
    Dim rst As DAO.Recordset
    
    Set rst = CurrentDb.OpenRecordset(OptionTable, dbOpenSnapshot)
    rst.MoveLast
    ReDim myOptions(rst.RecordCount - 1)
    rst.MoveFirst
    
    Do While Not rst.EOF
        myOptions(i) = rst!strValue
        i = i + 1
        rst.MoveNext
    Loop

    rst.Close
    Set rst = Nothing
End Sub

'***************************************************************************
'*************************** Enum-Creator-Prozedur *************************
'***************************************************************************
Private Sub CreateEnum()
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim strKeyArr As Variant
    Dim i As Long
    Dim SL As Long, SC As Long, EL As Long, EC As Long

    strKeyArr = ""                                               ''''''''
    Set rst = CurrentDb.OpenRecordset(OptionTable, dbOpenSnapshot)    '
    Do While Not rst.EOF                                                '
        strKeyArr = strKeyArr & ";" & rst!strKey                        '
        rst.MoveNext                                                    ' Ausdruck für Konstante zusammenbauen
    Loop                                                                '
    strKeyArr = Mid(strKeyArr, 2)                                       '
    rst.Close                                                           '
    Set rst = Nothing                                            ''''''''

    If strKeyArr = "" Then Exit Sub

    With Application.VBE.ActiveVBProject.VBComponents("OptionManagerhelper").CodeModule
        'Deklarationsblock löschen
        .DeleteLines 1, .CountOfDeclarationLines
        
        'Array aus Ausdruck erstellen
        strKeyArr = Split(strKeyArr, ";")

        'Enum einfügen
        .InsertLines 1, "Option Explicit"
        .InsertLines 2, ""
        .InsertLines 3, "Public Enum ltOptions"

        For i = LBound(strKeyArr) To UBound(strKeyArr)
            .InsertLines i + 4, "    " & strKeyArr(i) & " = " & i
        Next i
        
        .InsertLines i + 5, "End Enum"
    End With
End Sub

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
'Option Explicit

Private myOptions As Variant
Private FieldNameArr As Variant
Private m_OptionTable As String

'***************************************************************************
'************* Initialisierung und Terminierung der Klasse *****************
'***************************************************************************
Private Sub Class_Initialize()
    m_OptionTable = "tabOptionen"
    If Not TableExists Then
        CreateOptionTable
        If Not helperModuleExists Then CreateHelperModule
    Else
        If helperModuleExists Then
            'check for new options
            If hasNewOptions Then
                CreateEnum
            End If
            catchOptionValues
        Else
            Err.Raise 999, "OptionManager.Initialize", "HelperModul nicht vorhanden"
        End If
    End If
End Sub

Private Sub Class_Terminate()
    'nothing to do
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
    CurrentDb.Execute "Update " & m_OptionTable & " set strValue = '" & vValue & "' Where strKey = '" & cOption & "'"
    Exit Property

Settings_Err:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

'***************************************************************************
'********************* Prozedur zum Befüllen des Array *********************
'***************************************************************************
Private Sub catchOptionValues()
    Dim i As Long
    Dim rst As DAO.Recordset
    Set rst = CurrentDb.OpenRecordset(m_OptionTable, dbOpenSnapshot)
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
'****************************** Checker-Prozeduren *************************
'***************************************************************************
Private Function hasNewOptions() As Boolean
    Dim helperArr As Variant
    Dim rst As DAO.Recordset

    helperArr = Split(OptionManagerhelper_FieldArr, ";")

    If UBound(helperArr) < 0 Then
        hasNewOptions = True
        Exit Function
    End If

    Set rst = CurrentDb.OpenRecordset(m_OptionTable, dbOpenSnapshot)
    Do While Not rst.EOF
        If Not isInArray(helperArr, rst!strKey) Then
            hasNewOptions = True
            Exit Do
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Function

Private Function isInArray(ByVal vArr As Variant, ByVal cAusdruck As String) As Boolean
    Dim i As Long
    For i = LBound(vArr) To UBound(vArr)
        If vArr(i) = cAusdruck Then
            isInArray = True
        End If
    Next i
End Function

Private Function TableExists() As Boolean
    If m_OptionTable = "" Then m_OptionTable = "tabOptionen"
    TableExists = IsNull(DLookup("[Name]", "MSysObjects", "[Name] = '" & m_OptionTable & "' AND (Type = 1 Or Type = 6)")) = False
End Function

Private Function helperModuleExists() As Boolean
    helperModuleExists = IsNull(DLookup("[Name]", "MSysObjects", "[Name] = 'OptionManagerhelper' AND (Type = -32761)")) = False
End Function

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
    Set rst = CurrentDb.OpenRecordset(m_OptionTable, dbOpenSnapshot)    '
    Do While Not rst.EOF                                                '
        strKeyArr = strKeyArr & ";" & rst!strKey                        '
        rst.MoveNext                                                    ' Ausdruck für Konstante zusammenbauen
    Loop                                                                '
    strKeyArr = Mid(strKeyArr, 2)                                       '
    rst.Close                                                           '
    Set rst = Nothing                                            ''''''''
    
    If strKeyArr = "" Then Exit Sub
    
    With Application.VBE.ActiveVBProject.VBComponents("OptionManagerhelper").CodeModule
        'Codeblock finden und löschen
        .Find "Public Const OptionManagerhelper_FieldArr", SL, SC, EL, EC
        .DeleteLines SL, .CountOfDeclarationLines - SL + 1

        'Konstante einfügen
        .InsertLines SL, "Public Const OptionManagerhelper_FieldArr = """ & strKeyArr & """"

        'Array aus Ausdruck erstellen
        strKeyArr = Split(strKeyArr, ";")

        'Enum einfügen
        .InsertLines SL + 1, ""
        .InsertLines SL + 2, "Public Enum ltOptions"
        SL = SL + 3
        For i = LBound(strKeyArr) To UBound(strKeyArr)
            .InsertLines SL + i, "  " & strKeyArr(i) & " = " & i
        Next i
        .InsertLines SL + i, "End Enum"
    End With
End Sub

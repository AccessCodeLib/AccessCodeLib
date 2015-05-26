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
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private myOptions() As Variant
Private m_OptionTable As String

Private Const TableFieldEnumKey As String = "strKey"
Private Const TableFieldValue As String = "strValue"

'***************************************************************************
'************ Public Sub Update zum Generieren der Enum-Werte **************
'***************************************************************************
Friend Sub UpdateEnum()
    CreateEnum
End Sub

'***************************************************************************
'****** Zentrale Property Settings zum Holen und Setzen eines Wertes *******
'***************************************************************************
Public Property Get Settings(ByVal cOption As ltOptions) As Variant

On Error GoTo Settings_Err
    
    If (0 / 1) + (Not Not myOptions) = 0 Then 'Array noch nicht initialisiert
        CatchOptionValues
    End If
    
    Settings = myOptions(cOption)
    
Exit Property

Settings_Err:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Property

Public Property Let Settings(ByVal cOption As ltOptions, ByVal vValue As Variant)
    
On Error GoTo Settings_Err

    vValue = Replace(vValue, "'", "''")
    myOptions(cOption) = vValue
    CurrentDb.Execute "Update " & OptionTable & " set " & TableFieldValue & " = '" & vValue & "' Where " & TableFieldEnumKey & " = '" & cOption & "'"
    Exit Property

Settings_Err:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Property

'***************************************************************************
' Property Optiontable zum Setzen des Tabellennamens und Auslesen der Werte
'***************************************************************************
Public Property Get OptionTable() As String
    If Len(m_OptionTable) = 0 Then
        m_OptionTable = OptionManagerDefaultTableName
    End If
    OptionTable = m_OptionTable
End Property

Public Property Let OptionTable(ByVal cTable As String)
    m_OptionTable = cTable
End Property

'***************************************************************************
'********************* Prozedur zum Befüllen des Array *********************
'***************************************************************************
Private Sub CatchOptionValues()

    Dim i As Long
	dim strSouce as String
	
	strSouce = OptionTable
	If Len(strSouce) = 0 Then
		Err.Raise vbObjectError + 100, "OptionManager", "OptionTable not defined"
	End If
    
    With CurrentDb.OpenRecordset(strSouce, dbOpenSnapshot)
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            ReDim myOptions(.RecordCount - 1)
            
            Do While Not .EOF
                myOptions(i) = .Fields(TableFieldValue).Value
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With

End Sub

'***************************************************************************
'*************************** Enum-Creator-Prozedur *************************
'***************************************************************************
Private Sub CreateEnum()
    
    Dim EnumKeys() As String
    
    Dim i As Long
    Dim SL As Long, SC As Long, EL As Long, EC As Long

    Dim EnumDeclarationLine As Long
    Dim EnumEndLine As Long
    
    EnumKeys = GetEnumKeyArray

    With Application.VBE.ActiveVBProject.VBComponents("OptionManagerhelper")
        EnumDeclarationLine = GetEnumDeclarationLine(.CodeModule)
        If EnumDeclarationLine > 0 Then
            EnumEndLine = GetEnumCodeEndLine(.CodeModule, EnumDeclarationLine)
        End If
        
        With .CodeModule
        
            If EnumDeclarationLine > 0 Then
                .DeleteLines EnumDeclarationLine + 1, EnumEndLine - EnumDeclarationLine - 1
            Else
                EnumDeclarationLine = .CountOfLines
                .InsertLines EnumDeclarationLine, "Public Enum ltOptions"
                .InsertLines EnumDeclarationLine + 1, "End Enum"
            End If
        
            For i = LBound(EnumKeys) To UBound(EnumKeys)
                .InsertLines i + EnumDeclarationLine + 1, "    " & EnumKeys(i) & " = " & i
            Next i
            
        End With
    End With
End Sub

Private Function GetEnumKeyArray() As String()

    Dim strSQL As String
    Dim EnumKeys() As String
    Dim i As Long

    With CurrentDb.OpenRecordset(OptionTable, dbOpenSnapshot)
        If .RecordCount = 0 Then
            ReDim EnumKeys(0)
            EnumKeys(0) = "[_undefined]"
        Else
            .MoveLast
            .MoveFirst
            ReDim EnumKeys(.RecordCount - 1)
            Do While Not .EOF
                EnumKeys(i) = .Fields(TableFieldEnumKey).Value
                .MoveNext
                i = i + 1
            Loop
        End If
        .Close
    End With
    
    GetEnumKeyArray = EnumKeys
    
End Function

Private Function GetEnumDeclarationLine(ByVal cm As CodeModule) As Long

    Dim EnumStartLine As Long

    If cm.Find("Public Enum ltOptions", EnumStartLine, 0, 0, 0, False, False, False) Then
        GetEnumDeclarationLine = EnumStartLine
    Else
        GetEnumDeclarationLine = 0
    End If
    
End Function

Private Function GetEnumCodeEndLine(ByVal cm As CodeModule, ByVal EnumStartLine As Long) As Long

    Dim EnumEndLine As Long

    If cm.Find("End Enum", EnumStartLine, 0, EnumEndLine, 0, False, False, False) Then
        GetEnumCodeEndLine = EnumEndLine
    Else
        GetEnumCodeEndLine = 0
    End If
    
End Function

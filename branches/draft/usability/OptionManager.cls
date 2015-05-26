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

Private m_Settings() As Variant
Private m_DataSource As String

Private Const TableFieldEnumKey As String = "strKey"
Private Const TableFieldValue As String = "strValue"

'***************************************************************************
'****** Zentrale Property Setting zum Holen und Setzen eines Wertes *******
'***************************************************************************
Public Property Get Setting(ByVal cOption As ltOptions) As Variant
Attribute Setting.VB_UserMemId = 0
    If (0 / 1) + (Not Not m_Settings) = 0 Then 'Array noch nicht initialisiert
        CatchOptionValues
    End If
    
    Setting = m_Settings(cOption)
    
End Property

Public Property Let Setting(ByVal cOption As ltOptions, ByVal vValue As Variant)
    
    vValue = Replace(vValue, "'", "''")
    m_Settings(cOption) = vValue
    CurrentDb.Execute "Update " & DataSource & " set " & TableFieldValue & " = '" & vValue & "' Where " & TableFieldEnumKey & " = '" & cOption & "'"
    
End Property

'***************************************************************************************
' Property DataSource zum Setzen des Tabellennamens und Auslesen und Speichern der Werte
'***************************************************************************************
Public Property Get DataSource() As String
    If Len(m_DataSource) = 0 Then
        m_DataSource = OptionManagerDefaultDataSource
    End If
    DataSource = m_DataSource
End Property

Public Property Let DataSource(ByVal sDataSource As String)
    m_DataSource = sDataSource
End Property

'***************************************************************************
'********************* Prozedur zum Befüllen des Array *********************
'***************************************************************************
Private Sub CatchOptionValues()

    Dim i As Long
    Dim strSource As String
        
    strSource = DataSource
    If Len(strSource) = 0 Then
        Err.Raise vbObjectError + 100, "OptionManager", "DataSource not defined"
    End If
    
    With CurrentDb.OpenRecordset(strSource, dbOpenSnapshot)
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            ReDim m_Settings(.RecordCount - 1)
            
            Do While Not .EOF
                m_Settings(i) = .Fields(TableFieldValue).Value
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With

End Sub

'***************************************************************************
'************ Public Sub Update zum Generieren der Enum-Werte **************
'***************************************************************************
Public Sub UpdateEnum()
    
    Dim EnumKeys() As String
    
    Dim i As Long
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

    Dim RecordSource As String
    Dim EnumKeys() As String
    Dim i As Long
    Dim strSource As String
        
    strSource = DataSource
    If Len(strSource) = 0 Then
        Err.Raise vbObjectError + 100, "OptionManager", "DataSource not defined"
    End If

    With CurrentDb.OpenRecordset(strSource, dbOpenSnapshot)
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

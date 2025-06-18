VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Class Module: OptionManager
'---------------------------------------------------------------------------------------
'/**
' \author       Andreas Vogt; Josef Pötzel
' \brief        Klasse zur Verwendung als OptionManager
' \ingroup usability
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>usability/OptionManager.cls</file>
'  <license>_codelib/license.bas</license>
'  <test>_test/usability/OptionManagerTests.cls</test>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_Keys() As String
Private m_Values() As Variant
Private m_DataSource As String

Private Const EnumKeyFieldName As String = "strKey"
Private Const EnumValueFieldName As String = "strValue"
Private Const EnumName As String = "SettingKeys"

Private Declare PtrSafe Sub GetSafeArrayPointer Lib "msvbvm60.dll" Alias "GetMem4" (pArray() As Any, Ptr As Long)

'***************************************************************************
'****** Zentrale Property Setting zum Holen und Setzen eines Wertes ********
'***************************************************************************
Public Property Get Setting(ByVal SettingIndex As SettingKeys) As Variant
Attribute Setting.VB_UserMemId = 0

    CheckAndCatchOptionValues

    Setting = m_Values(SettingIndex)

End Property

Public Property Let Setting(ByVal SettingIndex As SettingKeys, ByVal vValue As Variant)

    CheckAndCatchOptionValues

    m_Values(SettingIndex) = vValue

    With CurrentDb
        .Execute "Update " & DataSource & " set " & EnumValueFieldName & " = '" & Replace(vValue, "'", "''") & "' Where " & EnumKeyFieldName & " = '" & m_Keys(SettingIndex) & "'"
        If .RecordsAffected = 0 Then    'derzeit nicht möglich, da in so einem Fall zuvor schon ein Fehler bei m_Values(cOption) ausgelöst würde
            .Execute "Insert into " & DataSource & " (" & EnumKeyFieldName & ", " & EnumValueFieldName & ")" & _
                   " Values ('" & m_Keys(SettingIndex) & "', '" & Replace(vValue, "'", "''") & "')"
        End If
    End With

End Property

Public Property Get SettingByName(ByVal KeyName As String) As Variant

    Dim i As Long
    
    CheckAndCatchOptionValues
    For i = LBound(m_Keys) To UBound(m_Keys)
        If m_Keys(i) = KeyName Then
            SettingByName = m_Values(i)
        End If
    Next i

End Property

Public Property Let SettingByName(ByVal KeyName As String, ByVal vValue As Variant)

    CheckAndCatchOptionValues

    With CurrentDb
        If KeyNameExists(KeyName) Then
            .Execute "Update " & DataSource & " set " & EnumValueFieldName & " = '" & Replace(vValue, "'", "''") & "' Where " & EnumKeyFieldName & " = '" & KeyName & "'"
        Else
            .Execute "Insert into " & DataSource & " (" & EnumKeyFieldName & ", " & EnumValueFieldName & ")" & _
                   " Values ('" & KeyName & "', '" & Replace(vValue, "'", "''") & "')"
        End If
        CatchOptionValues
    End With

End Property

'****************************************************************************
'************** Get-Properties fuer Key-Name und Anzahl Keys ****************
'****************************************************************************
Public Property Get KeyName(ByVal SettingIndex As SettingKeys) As String

    CheckAndCatchOptionValues
    
    KeyName = m_Keys(SettingIndex) ' Wenn SettingIndex nicht enthalten ist, wird von VBA Fehler ausgelöst, dass sollte ausreichen

End Property

Public Property Get Count() As Long

    If (0 / 1) + (Not Not m_Values) = 0 Then
        CatchOptionValues
    End If

    If (0 / 1) + (Not Not m_Values) = 0 Then
        Err.Raise vbObjectError + 102, "OptionManager", "no existing Options"
    Else
        Count = UBound(m_Keys)
    End If

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
'***************************** Public Subs *********************************
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
                .InsertLines EnumDeclarationLine, "Public Enum " & EnumName
                .InsertLines EnumDeclarationLine + 1, "End Enum"
            End If

            If LBound(EnumKeys) = 0 Then
               EnumDeclarationLine = EnumDeclarationLine + 1
            End If
            For i = LBound(EnumKeys) To UBound(EnumKeys)
                .InsertLines i + EnumDeclarationLine, "    " & EnumKeys(i) & " = " & i
            Next i

        End With
    End With
End Sub

Friend Sub DeleteByName(ByVal KeyName As String)

    With CurrentDb
        If KeyNameExists(KeyName) Then
            .Execute "Delete from " & DataSource & " Where " & EnumKeyFieldName & " = '" & KeyName & "'"
            CatchOptionValues
        Else
            Err.Raise vbObjectError + 101, "OptionManager", "Keyname does not exist"
        End If
    End With

End Sub

'***************************************************************************
'********************** Private Class Functions ****************************
'***************************************************************************

Private Sub CheckAndCatchOptionValues()
    If Not ValueArrayIsInitialized Then
        CatchOptionValues
    End If
End Sub

Private Property Get ValueArrayIsInitialized() As Boolean
    
    Dim SafeArrayPointer As Long
    GetSafeArrayPointer m_Values, SafeArrayPointer
    
    ValueArrayIsInitialized = (SafeArrayPointer > 0)

End Property


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

            ReDim m_Keys(1 To .RecordCount)
            ReDim m_Values(1 To .RecordCount)

            Do While Not .EOF
                i = i + 1
                m_Keys(i) = .Fields(EnumKeyFieldName).Value
                m_Values(i) = .Fields(EnumValueFieldName).Value
                .MoveNext
            Loop
        Else
           '0-Array erzeugen, da Enums nur mit 0 beginnen, wenn [_undefined] = 0 enthalten ist.
           ReDim m_Keys(0)
           ReDim m_Values(0)
           m_Keys(0) = "[_undefined]"
           m_Values(0) = Null
        End If
        .Close
    End With

End Sub

Private Function GetEnumKeyArray() As String()

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
            ReDim EnumKeys(1 To .RecordCount)
            Do While Not .EOF
                i = i + 1
                EnumKeys(i) = .Fields(EnumKeyFieldName).Value
                .MoveNext
            Loop
        End If
        .Close
    End With

    GetEnumKeyArray = EnumKeys

End Function

Private Function GetEnumDeclarationLine(ByVal cm As Object) As Long

    Dim EnumStartLine As Long

    If cm.Find("Public Enum " & EnumName, EnumStartLine, 0, 0, 0, False, False, False) Then
        GetEnumDeclarationLine = EnumStartLine
    Else
        GetEnumDeclarationLine = 0
    End If

End Function

Private Function GetEnumCodeEndLine(ByVal cm As Object, ByVal EnumStartLine As Long) As Long

    Dim EnumEndLine As Long

    If cm.Find("End Enum", EnumStartLine, 0, EnumEndLine, 0, False, False, False) Then
        GetEnumCodeEndLine = EnumEndLine
    Else
        GetEnumCodeEndLine = 0
    End If

End Function

Private Function KeyNameExists(ByVal KeyName As String) As Boolean
    Dim rst As DAO.Recordset
    
    With CurrentDb
        Set rst = .OpenRecordset("Select * From " & DataSource & " Where " & EnumKeyFieldName & " = '" & KeyName & "'", dbOpenSnapshot)
        KeyNameExists = Not rst.EOF
        rst.Close
        Set rst = Nothing
    End With
    
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddInConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: _codelib.addins.shared.AddInConfiguration
'---------------------------------------------------------------------------------------
'
' Manage Access Add-In configuration settings
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/AddInConfiguration.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_AppTitle As String
Private m_Title As String
Private m_Author As String
Private m_Company As String
Private m_Comments As String

Private m_AddInRegPathName As String
Private m_AddInRegFilePath As String
Private m_AddInStartFunction As String

Private m_InsertRecordset As DAO.Recordset

Public Property Get FileName() As String
   If Len(m_AddInRegFilePath) > 0 Then
      FileName = Mid(m_AddInRegFilePath, InStrRev(m_AddInRegFilePath, "\", , vbTextCompare) + 1)
   End If
End Property

Public Property Let FileName(ByVal NewValue As String)
   If Len(NewValue) > 0 Then
      AddInRegFilePath = "|ACCDIR\" & NewValue
   End If
End Property

Public Property Get AppTitle() As String
   AppTitle = m_AppTitle
End Property

Public Property Let AppTitle(ByVal NewValue As String)
   m_AppTitle = NewValue
End Property

Public Property Get Title() As String
   Title = m_Title
End Property

Public Property Let Title(ByVal NewValue As String)
   m_Title = NewValue
End Property

Public Property Get Author() As String
   Author = m_Author
End Property

Public Property Let Author(ByVal NewValue As String)
   m_Author = NewValue
End Property

Public Property Get Company() As String
   Company = m_Company
End Property

Public Property Let Company(ByVal NewValue As String)
   m_Company = NewValue
End Property

Public Property Get Comments() As String
   Comments = m_Comments
End Property

Public Property Let Comments(ByVal NewValue As String)
   m_Comments = NewValue
End Property

Public Property Get AddInRegPathName() As String
   AddInRegPathName = m_AddInRegPathName
End Property

Public Property Let AddInRegPathName(ByVal NewValue As String)
   m_AddInRegPathName = NewValue
End Property

Public Property Get AddInStartFunction() As String
   AddInStartFunction = m_AddInStartFunction
End Property

Public Property Let AddInStartFunction(ByVal NewValue As String)
   m_AddInStartFunction = NewValue
End Property

Public Property Get AddInRegFilePath() As String
   AddInRegFilePath = m_AddInRegFilePath
End Property

Public Property Let AddInRegFilePath(ByVal NewValue As String)
   m_AddInRegFilePath = NewValue
End Property

Public Sub LoadFromCurrentFile()

   If CurrentDb Is Nothing Then
      Exit Sub
   End If

   LoadDataFromUSysRegInfo
   LoadDataFromDatabase

End Sub

Private Sub LoadDataFromDatabase()

   Dim db As DAO.Database
   Set db = CurrentDb

   m_AppTitle = GetDbProperty(db, "AppTitle", vbNullString)

   If Len(FileName) = 0 Then
      FileName = Replace(CurrentProject.Name, ".accdb", ".accda", , , vbTextCompare)
   End If

   m_Title = GetDocProperty(db, "Title", vbNullString)
   m_Author = GetDocProperty(db, "Author", vbNullString)
   m_Company = GetDocProperty(db, "Company", vbNullString)
   m_Comments = GetDocProperty(db, "Comments", vbNullString)

End Sub

Private Function GetDbProperty(ByVal PropDb As DAO.Database, PropertyName As String, Optional ByVal ValueIfNotExists As Variant = Null) As Variant

   On Error Resume Next

   GetDbProperty = PropDb.Properties(PropertyName)
   If Err.Number <> 0 Then
      Err.Clear
      GetDbProperty = ValueIfNotExists
   End If

End Function

Private Function GetDocProperty(ByVal PropDb As DAO.Database, PropertyName As String, Optional ByVal ValueIfNotExists As Variant = Null) As String

On Error Resume Next

   Dim PropValue As String

   PropValue = PropDb.Containers("Databases").Documents("SummaryInfo").Properties(PropertyName)
   If Err.Number <> 0 Then
      Err.Clear
      PropValue = ValueIfNotExists
   End If

   GetDocProperty = PropValue

End Function

Private Sub LoadDataFromUSysRegInfo()
'Subkey                                                              Type  ValName     Value
'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\ACLib Access Add-In Builder 0
'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\ACLib Access Add-In Builder 1    Expression   =LoadAddIn()
'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\ACLib Access Add-In Builder 1    Library     |ACCDIR\ACLibAccessAddInBuilder.accda

   Dim AddInRegPathName As String
   Dim AddInStartFunction As String

   Dim ValNameField As DAO.Field
   Dim SubkeyField As DAO.Field
   Dim ValueField As DAO.Field

   If Not Nz(DLookup("True", "MSysObjects", "[Name]='USysRegInfo'"), False) Then
      Exit Sub
   End If

   With CurrentDb.OpenRecordset("select * from USysRegInfo Order By Subkey, ValName")
      If Not .EOF Then

         Set SubkeyField = .Fields("Subkey")
         Set ValNameField = .Fields("ValName")
         Set ValueField = .Fields("Value")

         Do While Not .EOF
            If IsNull(ValNameField.Value) Then
               With .Fields("Subkey")
                  m_AddInRegPathName = Mid(.Value, InStrRev(.Value, "\", , vbTextCompare) + 1)
               End With

            ElseIf ValNameField.Value = "Expression" Then
               m_AddInStartFunction = Nz(ValueField.Value, vbNullString)

            ElseIf ValNameField.Value = "Library" Then
               m_AddInRegFilePath = Nz(ValueField.Value, vbNullString)

            End If
            .MoveNext
         Loop
      End If
      .Close
   End With

End Sub

Public Sub SaveToCurrentFile()

   If CurrentDb Is Nothing Then
      Exit Sub
   End If

   SaveDataToUSysRegInfo
   SaveDataToDatabase

End Sub

Private Sub SaveDataToDatabase()

   Dim db As DAO.Database
   Set db = CurrentDb

   SetDbProperty db, "AppTitle", m_AppTitle

   SetDocProperty db, "Title", m_Title
   SetDocProperty db, "Author", m_Author
   SetDocProperty db, "Company", m_Company
   SetDocProperty db, "Comments", m_Comments

End Sub

Private Sub SetDbProperty(ByVal PropDb As DAO.Database, PropertyName As String, ByVal NewValue As Variant)

On Error GoTo HandleErr

   PropDb.Properties(PropertyName).Value = Nz(NewValue, vbNullString)

ExitHere:
   Exit Sub

HandleErr:
   If Err.Number = 3270 Then
      PropDb.Properties.Append PropDb.CreateProperty(PropertyName, dbText, NewValue)
      Resume ExitHere
   End If

   HandleError Err.Number, Err.Source, Err.Description, aclibErrRaise

End Sub

Private Sub SetDocProperty(ByVal PropDb As DAO.Database, PropertyName As String, ByVal NewValue As Variant)

On Error GoTo HandleErr

   PropDb.Containers("Databases").Documents("SummaryInfo").Properties(PropertyName).Value = Nz(NewValue, vbNullString)

ExitHere:
   Exit Sub

HandleErr:
   If Err.Number = 3270 Then
      If Len(NewValue) > 0 Then
         PropDb.Containers("Databases").Documents("SummaryInfo").Properties.Append PropDb.CreateProperty(PropertyName, dbText, NewValue)
      End If
      Resume ExitHere
   End If

   HandleError Err.Number, Err.Source, Err.Description, aclibErrRaise

End Sub

Private Sub SaveDataToUSysRegInfo()
'Subkey   Type  ValName  Value
'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\ACLib Access Add-In Builder 0
'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\ACLib Access Add-In Builder 1  Expression  =LoadAddIn()
'HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\ACLib Access Add-In Builder 1  Library  |ACCDIR\ACLibAccessAddInBuilder.accda
'
   Dim SubKey As String
   Dim Expression As String
   Dim Library As String

   If Len(m_AddInRegPathName) = 0 Then
      MsgBox "Enter USysRegInfo-Name first!"
      Exit Sub
   End If
   SubKey = "HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\" & m_AddInRegPathName


   If Len(m_AddInStartFunction) = 0 Then
      m_AddInStartFunction = "=LoadAddIn()"
   End If
   Expression = m_AddInStartFunction
   If Left(Expression, 1) <> "=" Then
      Expression = "=" & Expression
   End If
   If Right(Expression, 1) <> ")" Then
      Expression = Expression & "()"
   End If

   Library = m_AddInRegFilePath

   If Not Nz(DLookup("True", "MSysObjects", "[Name]='USysRegInfo'"), False) Then
      CreateTableUSysRegInfo
   End If

   Dim db As DAO.Database
   Set db = CurrentDb

   db.Execute "delete from USysRegInfo where Type = 0 or Type = 1 and ValName in ('Expression', 'Library')"

   Dim rs As DAO.Recordset
   Set rs = db.OpenRecordset("select * from USysRegInfo")

   InsertUSysRegInfoRecord rs, SubKey, 0, Null, Null
   InsertUSysRegInfoRecord rs, SubKey, 1, "Expression", Expression
   InsertUSysRegInfoRecord rs, SubKey, 1, "Library", Library

   rs.Close

End Sub

Private Sub InsertUSysRegInfoRecord(ByVal RegInfoRs As DAO.Recordset, ByVal SubKey As String, ByVal TypeNo As Long, ByVal ValName As Variant, ByVal NewValue As Variant)

   With RegInfoRs
      .AddNew
      .Fields("Subkey").Value = SubKey
      .Fields("Type").Value = TypeNo
      .Fields("ValName").Value = ValName
      .Fields("Value").Value = NewValue
      .Update
   End With

End Sub

Private Sub CreateTableUSysRegInfo()

   Dim CreateTableDDL As String

   CreateTableDDL = "create Table USysRegInfo (" & _
                    " Subkey nvarchar(255), Type int, ValName nvarchar(255), [Value] nvarchar(255))"

   CurrentProject.Connection.Execute CreateTableDDL

End Sub

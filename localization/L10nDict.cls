VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "L10nDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Attribute VB_PredeclaredId = True
'---------------------------------------------------------------------------------------
' Klasse: L10nDict
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Localization dictionary - Auxiliary class for language setting
' </summary>
' <remarks>Set attributes VB_PredeclaredId = True so that class can be used without instantiation</remarks>
''
' \ingroup localization
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>localization/L10nDict.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/SqlTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const LngDictTableName As String = "L10n_Dict"
Private Const DefaultLangSeparatorChar As String = "|"

Private Const DefaultBaseLongCode As String = "EN"
Private m_BaseLangCode As String

Private m_LangCode As String
Private m_LangSeparatorChar As String

#Const ScriptingDictEarlyBinding = 0
#If ScriptingDictEarlyBinding Then
   Private m_Dict As Scripting.Dictionary
#Else
   Private m_Dict As Object
#End If

Public Event LanguageChanged()
Public Event DictionaryRefreshed()

#If VBA7 Then
    Private Declare PtrSafe Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
#Else
    Private Declare Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
#End If

#If DictEarlyBinding Then
Private Property Get Dict() As Scripting.Dictionary
#Else
Private Property Get Dict() As Object
#End If
   CheckDictionary
   Set Dict = m_Dict
End Property

Private Sub CheckDictionary()
   If m_Dict Is Nothing Then
      FillDict
   End If
End Sub

Private Sub FillDict()

   Dim rst As DAO.Recordset
   Dim Key As DAO.Field
   Dim Msg As DAO.Field
   
   Set m_Dict = CreateObject("Scripting.Dictionary")
      
On Error GoTo HandleErr
     
   Set rst = CodeDb.OpenRecordset("select KeyText, LngText from " & LngDictTableName & " where " & SqlTools.BuildCriteria("LangCode", SQL_Text, SQL_Equal, LangCode), dbOpenForwardOnly)
   With rst
      Set Key = .Fields("KeyText")
      Set Msg = .Fields("LngText")
   End With

   With m_Dict
      Do While Not rst.EOF
         .Add Key.Value, Msg.Value
         rst.MoveNext
      Loop
   End With
   
   rst.Close
   
   RaiseEvent DictionaryRefreshed
   
Exit Sub

HandleErr:
   If Err.Number = 3078 Then
      If CreateL10nDictTable Then
         Resume
      End If
   End If
   Err.Raise vbObjectError + Err.Number, "L10nDict.FillDict", Err.Description

End Sub

Friend Function CreateL10nDictTable() As Boolean

   Dim CreateTableDDL As String
   
On Error GoTo HandleErr
   
   CreateTableDDL = "create table " & LngDictTableName & "(" & _
                    " LangCode varchar(2), KeyText varchar(255), LngText varchar(255)" & _
                    ", CONSTRAINT PK_" & LngDictTableName & " PRIMARY KEY (LangCode, KeyText)" & _
                    ")"
   
   CodeProject.Connection.Execute CreateTableDDL
   Application.RefreshDatabaseWindow
   
   CreateL10nDictTable = True
   
   Exit Function
   
HandleErr:
   CreateL10nDictTable = False

End Function

Public Property Get LangCode() As String
   If Len(m_LangCode) = 0 Then
      If GetUserLanguageID = 1031 Then ' msoLanguageIDUI = 2
         m_LangCode = "DE"
      Else
         m_LangCode = "EN"
      End If
   End If
   LangCode = m_LangCode
End Property

Public Property Let LangCode(ByVal NewValue As String)
   m_LangCode = NewValue
   FillDict
   RaiseEvent LanguageChanged
End Property

Public Property Get BaseLangCode() As String
   If Len(m_BaseLangCode) = 0 Then
      m_BaseLangCode = DefaultBaseLongCode
   End If
   BaseLangCode = m_BaseLangCode
End Property

Public Property Let BaseLangCode(ByVal NewValue As String)
   m_BaseLangCode = NewValue
End Property

Public Property Get LangSeparatorChar() As String
   If Len(m_LangSeparatorChar) = 0 Then
      m_LangSeparatorChar = DefaultLangSeparatorChar
   End If
   LangSeparatorChar = m_LangSeparatorChar
End Property

Public Property Let LangSeparatorChar(ByVal NewValue As String)
   m_LangSeparatorChar = NewValue
End Property

Public Function GetUserLanguageID() As Long

   Dim AccApp As Object
   
   Set AccApp = Application
   
On Error GoTo AccAppInterFaceErr
   GetUserLanguageID = AccApp.LanguageSettings.LanguageID(2)

Exit Function

AccAppInterFaceErr:
   GetUserLanguageID = GetUserDefaultUILanguage

End Function

Public Function Text(ByVal KeyText As String) As String

   With Dict
      If Not .Exists(KeyText) Then
         AddTextToTable KeyText
         .Add KeyText, KeyText
      End If
      Text = .Item(KeyText)
   End With

End Function

Public Function ControlTag(ByVal TagText As String) As String
   ControlTag = RTrim(Split(TagText & " ", LangSeparatorChar)(0))
End Function

Private Sub AddTextToTable(ByVal BaseText As String)
   
   Dim LngText As String
   
   LngText = BaseText
   If m_LangCode <> BaseLangCode Then
      LngText = m_LangCode & ":" & LngText
   End If
   
   With CodeDb.OpenRecordset(LngDictTableName, dbOpenTable, dbAppendOnly)
      .AddNew
      .Fields("LangCode").Value = LangCode
      .Fields("KeyText").Value = BaseText
      .Fields("LngText").Value = LngText
      .Update
      .Close
   End With
   
End Sub

Public Sub TranslateControls(ByVal ControlList As Controls)

   Dim ctl As Control
   
   For Each ctl In ControlList
      TranslateControl ctl
   Next

End Sub

Private Sub TranslateControl(ByVal CtlRef As Control)

   Dim LangTextKeyString As String
   Dim ControlTextKeys() As String
   Dim TagText As String
   Dim TagTextPrefix As String
   Dim Pos As Long

   Dim TextKeySeparatorString As String
   TextKeySeparatorString = Chr(1)

   TagText = CtlRef.Tag
   
   If Len(TagText) = 0 Then
      Exit Sub
   End If
   
   Pos = InStr(1, TagText, LangSeparatorChar & "LANG:")
   If Pos > 0 Then
      TagTextPrefix = Left(TagText, Pos)
      TagText = Mid(TagText, Pos + Len(LangSeparatorChar))
   End If

   If Left(TagText, 5) <> "LANG:" Then
      Exit Sub
   End If

   LangTextKeyString = Mid(TagText, 6)
   With CtlRef
      If Len(LangTextKeyString) = 0 Then
         ControlTextKeys = GetControlTextKeys(CtlRef)
         LangTextKeyString = Join(ControlTextKeys, TextKeySeparatorString)
         .Tag = TagTextPrefix & "LANG:" & LangTextKeyString
      Else
         ControlTextKeys = Split(LangTextKeyString, TextKeySeparatorString)
         If UBound(ControlTextKeys) < 3 Then
            ReDim Preserve ControlTextKeys(3)
         End If
      End If
   End With
   
   SetControlText CtlRef, ControlTextKeys
      
End Sub

Private Function GetControlTextKeys(ByVal CtlRef As Control) As String()

   Dim ControlTextKeys(3) As String
   Dim CaptionText As String
   Dim ControlTipText As String
   Dim StatusBarText As String
   Dim FormatText As String

   With CtlRef
      ControlTipText = .ControlTipText
      
      Select Case .ControlType
         Case AcControlType.acLabel
            CaptionText = .Caption
         Case AcControlType.acCommandButton, AcControlType.acToggleButton
            CaptionText = .Caption
            StatusBarText = .StatusBarText
         Case AcControlType.acTextBox
            FormatText = .Format
            StatusBarText = .StatusBarText
         Case Else
            StatusBarText = .StatusBarText
      End Select
   End With
   
   ControlTextKeys(0) = CaptionText
   ControlTextKeys(1) = ControlTipText
   ControlTextKeys(2) = StatusBarText
   ControlTextKeys(3) = FormatText
   
   GetControlTextKeys = ControlTextKeys

End Function

Private Sub SetControlText(ByVal CtlRef As Control, ByRef ControlTextKeys() As String)

   With CtlRef
      If Len(ControlTextKeys(0)) > 0 Then
         .Caption = Text(ControlTextKeys(0))
      End If
   
      If Len(ControlTextKeys(1)) > 0 Then
         .ControlTipText = Text(ControlTextKeys(1))
      End If
   
      If Len(ControlTextKeys(2)) > 0 Then
         .StatusBarText = Text(ControlTextKeys(2))
      End If
      
      If Len(ControlTextKeys(3)) > 0 Then
         .Format = Text(ControlTextKeys(3))
      End If
      
   End With

End Sub

Public Function MsgBox(ByVal Prompt As Variant, _
              Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
              Optional ByVal Title As Variant, _
              Optional ByVal HelpFile As Variant, _
              Optional ByVal Context As Variant) As VbMsgBoxResult
   
   If Not IsMissing(Prompt) Then
      Prompt = Me.Text(Prompt)
   End If
   
   If Not IsMissing(Title) Then
      Title = Me.Text(Title)
   End If
   
   MsgBox = VBA.MsgBox(Prompt, Buttons, Title, HelpFile, Context)

End Function

Public Function InputBox(ByVal Prompt As Variant, _
                Optional ByVal Title As Variant, _
                Optional ByVal Default As Variant, _
                Optional ByVal XPos As Variant, Optional ByVal YPos As Variant, _
                Optional ByVal HelpFile As Variant, _
                Optional ByVal Context As Variant) As String

   If Not IsMissing(Prompt) Then
      Prompt = Me.Text(Prompt)
   End If
   
   If Not IsMissing(Title) Then
      Title = Me.Text(Title)
   End If
   
   InputBox = VBA.InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)

End Function

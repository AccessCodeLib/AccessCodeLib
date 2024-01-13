VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACLibGitHubImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: _codelib.addins.shared.ACLibGitHubImporter
'---------------------------------------------------------------------------------------
'
' Import GitHub files
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/ACLibGitHubImporter.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const GitHubContentBaseUrl As String = "https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{path}"
Private Const GitHubApiBaseUrl As String = "https://api.github.com/repos/{owner}/{repo}/"

Private m_GitHubApiAuthorizationToken As String
Private m_LastCommit As Date

Private m_RepositoryOwner As String
Private m_RepositoryName As String
Private m_BranchName As String

#If VBA7 Then
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

'---------------------------------------------------------------------------------------
' Property: GitHubAuthorizationubAuthToken
'---------------------------------------------------------------------------------------
Public Property Get GitHubApiAuthorizationToken() As String
   GitHubApiAuthorizationToken = m_GitHubApiAuthorizationToken
End Property

Public Property Let GitHubApiAuthorizationToken(ByVal NewValue As String)
   m_GitHubApiAuthorizationToken = NewValue
End Property

'---------------------------------------------------------------------------------------
' Property: RepositoryOwner
'---------------------------------------------------------------------------------------
Public Property Get RepositoryOwner() As String
   If Len(m_RepositoryOwner) > 0 Then
      RepositoryOwner = m_RepositoryOwner
   Else ' Default: AccessCodeLib
      RepositoryOwner = "AccessCodeLib"
   End If
End Property

Public Property Let RepositoryOwner(ByVal NewValue As String)
   m_RepositoryOwner = NewValue
End Property

'---------------------------------------------------------------------------------------
' Property: RepositoryName
'---------------------------------------------------------------------------------------
Public Property Get RepositoryName() As String
   If Len(m_RepositoryName) > 0 Then
      RepositoryName = m_RepositoryName
   Else ' Default: AccessCodeLib
      RepositoryName = "AccessCodeLib"
   End If
End Property

Public Property Let RepositoryName(ByVal NewValue As String)
   m_RepositoryName = NewValue
End Property

'---------------------------------------------------------------------------------------
' Property: BranchName
'---------------------------------------------------------------------------------------
Public Property Get BranchName() As String
   If Len(m_BranchName) > 0 Then
      BranchName = m_BranchName
   Else ' Default: master
      BranchName = "master"
   End If
End Property

Public Property Let BranchName(ByVal NewValue As String)
   m_BranchName = NewValue
End Property

'---------------------------------------------------------------------------------------
' Property: RevisionString
'---------------------------------------------------------------------------------------
Public Property Get RevisionString(Optional ByVal Requery As Boolean = False) As String
   RevisionString = Format(LastCommit, "yyyymmddhhnnss")
   If BranchName <> "master" Then
      RevisionString = RevisionString & "-" & BranchName
   End If
End Property

'---------------------------------------------------------------------------------------
' Property: LastCommit
'---------------------------------------------------------------------------------------
Public Property Get LastCommit(Optional ByVal Requery As Boolean = False) As String
   If m_LastCommit = 0 Or Requery Then
      m_LastCommit = GetLastCommitFromWeb()
   End If
   LastCommit = m_LastCommit
End Property

'---------------------------------------------------------------------------------------
' Sub: UpdateCodeModules
'---------------------------------------------------------------------------------------
Public Sub UpdateCodeModules()
   
   Dim SelectSql As String
   Dim IsFirstRecord As Boolean
   
   SelectSql = "select id, url from usys_Appfiles where url > ''"
   
   With CreateObject("ADODB.Recordset")
      .CursorLocation = 3 'adUseClient
      .Open SelectSql, CodeProject.Connection, 1, 1 ' 1 = adOpenKeyset, 1 = adLockReadOnly
      Set .ActiveConnection = Nothing
   
      IsFirstRecord = True
      Do While Not .EOF
         UpdateCodeModuleInTable .Fields(0).Value, .Fields(1).Value, IsFirstRecord
         If IsFirstRecord Then IsFirstRecord = False
         .MoveNext
      Loop
      
      .Close

   End With

End Sub

Private Sub UpdateCodeModuleInTable(ByVal ModuleName As String, ByVal ACLibPath As String, Optional ByVal Requery As Boolean = False)

   Dim TempFile As String
   
   
   TempFile = FileTools.TempPath & ModuleName & FileTools.GetFileExtension(ACLibPath, True)
   DownloadACLibFileFromWeb ACLibPath, TempFile
   
   CurrentApplication.SaveAppFile ModuleName, TempFile, False, "SccRev", Me.RevisionString(Requery)
   Kill TempFile

End Sub

Friend Sub DownloadACLibFileFromWeb(ByVal ACLibPath As String, ByVal TargetFilePath As String)

   Dim DownLoadUrl As String

   DownLoadUrl = FillRepositoryData(GitHubContentBaseUrl)
   DownLoadUrl = Replace(DownLoadUrl, "{path}", ACLibPath)
   
   DownloadFileFromWeb DownLoadUrl, TargetFilePath

End Sub

Private Function FillRepositoryData(ByVal StringWithPlaceHolder As String) As String
   
   Dim TempValue As String
   
   TempValue = Replace(StringWithPlaceHolder, "{owner}", RepositoryOwner)
   TempValue = Replace(TempValue, "{repo}", RepositoryName)
   TempValue = Replace(TempValue, "{branch}", BranchName)
   
   FillRepositoryData = TempValue
   
End Function

Private Function GetLastCommitFromWeb() As Date

'alternative: git rev-list HEAD --count

' https://api.github.com/repos/AccessCodeLib/AccessCodeLib/commits/master

   Dim CommitUrl As String
   Dim LastCommitInfo As String
   
   CommitUrl = FillRepositoryData(GitHubApiBaseUrl) & "commits/" & BranchName

   Const RevisionTag As String = "Revision "

   Dim JsonString As String
   JsonString = GetJsonString(CommitUrl)
   
   Dim LastCommitPos As Long
   LastCommitPos = InStr(1, JsonString, """committer"":")
   LastCommitPos = InStr(LastCommitPos, JsonString, """date"":") + Len("date"": """)
   '"date": "2023-05-14T09:34:04Z"
   LastCommitInfo = Mid(JsonString, LastCommitPos, Len("2023-05-14T09:34:04"))
   
   GetLastCommitFromWeb = CDate(Replace(LastCommitInfo, "T", " "))

End Function

Friend Function GetJsonString(ByVal ApiUrl As String) As String

   Dim ApiResponse As String
   Dim ApiAuthToken As String
   Dim json As Object
   Dim xml As Object 'MSXML2.XMLHTTP6
   
   ApiUrl = FillRepositoryData(ApiUrl)
   
   ApiAuthToken = GitHubApiAuthorizationToken

   Set xml = CreateObject("MSXML2.XMLHTTP.6.0")
   
   xml.Open "GET", ApiUrl, False
   If Len(ApiAuthToken) > 0 Then
     xml.setRequestHeader "Authorization", ApiAuthToken
   End If
   xml.setRequestHeader "Content-type", "application/json"
   xml.send
   While xml.ReadyState <> 4
       DoEvents
   Wend
   ApiResponse = xml.responseText
   
   GetJsonString = ApiResponse

End Function

Private Sub OpenIEandLoadHtmlDoc(ByVal Url As String, ByRef IE As Object, ByRef HtmlDoc As Object)

   Dim TimeOut As Long
   Dim RunInTimeOut As Boolean
   Dim ErrHdlCnt As Long
   
   Dim ErrNumber As Long
   Dim ErrDescription As String

On Error Resume Next
   Set IE = CreateObject("InternetExplorer.Application")
   Do While Err.Number = -2147023706 And ErrHdlCnt < 10
      Err.Clear
      ErrHdlCnt = ErrHdlCnt + 1
      Set IE = CreateObject("InternetExplorer.Application")
   Loop

   If Err.Number <> 0 Then
      ErrNumber = Err.Number
      ErrDescription = Err.Description
   On Error GoTo 0
      Err.Raise ErrNumber, "ACLibWebImporter.OpenIEandLoadHtmlDoc", ErrDescription
   End If
   
On Error GoTo 0
   
   With IE
      TimeOut = Timer + 10
      Do While .Busy And (Not RunInTimeOut)
         DoEvents
         If Timer > TimeOut Then RunInTimeOut = True
      Loop
      
      If Not RunInTimeOut Then
         .Visible = 0
         .navigate Url
         TimeOut = Timer + 10
         Do Until .ReadyState = 4 Or RunInTimeOut
            DoEvents
            If Timer > TimeOut Then RunInTimeOut = True
         Loop
      End If
      
      If RunInTimeOut Then
         On Error Resume Next
         IE.Quit
         Set IE = Nothing
         On Error GoTo 0
         Err.Raise vbObjectError, "OpenIEandLoadHtmlDoc", "Time-Out beim Laden von '" & Url & "'"
      End If
      
      Set HtmlDoc = IE.Document
             
   End With

End Sub

Private Sub DownloadFileFromWeb(ByVal Url As String, ByVal TargetPath As String)
   If FileExists(TargetPath) Then Kill TargetPath
   DeleteUrlCacheEntry Url
   URLDownloadToFile 0, Url, TargetPath, 0, 0
End Sub

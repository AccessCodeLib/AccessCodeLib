VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecursiveFileSearcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: RecursiveFileSearcher
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Verzeichnisstruktur nach einer oder mehreren Dateien durchsuchen
' </summary>
' <remarks>
' <code>
' Dim Itm As Variant
' With New RecursiveFileSearcher
'    .AddSearchItem "msaccess.*"
'    .AddSearchItem "winword.exe"
'
'    If Not .Search("C:\Program Files (x86)\Microsoft Office", FSMode_SearchAllFolders) Then
'       Debug.Print "Kein Treffer"
'    Else
'       For Each Itm In .Matches
'          Debug.Print Itm
'       Next
'    End If
' End With
' </code>
' </remarks>
' \ingroup file
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>file/RecursiveFileSearcher.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Enum: FileSearchMode
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Suchvarianten
' </summary>
' <list type="table">
'   <item><term>FSMode_SearchOneOfEveryItem (0)</term><description>Sobald Suchbegriff einmal gefunden wurde, diesen Begriff nicht weiter suchen. Die Suche läuft, bis das Verzeichnis durchsucht wurde oder jeder Suchbegriff 1x gefunden wurde.</description></item>
'   <item><term>FSMode_SearchOneFile (1)</term><description>Suche beenden, sobald einer der Suchbegriffe einen Treffer ergab</description></item>
'   <item><term>FSMode_SearchAllFolders (2)</term><description>Die komplette Verzeichnisstruktur durchsuchen -> je Suchbegriff kann es mehrere Treffer geben</description></item>
' </list>
'**/
Public Enum FileSearchMode
   FSMode_SearchOneOfEveryItem = 0  ' Sobald Suchbegriff einmal gefunden wurde, diesen Begriff nicht weiter suchen.
                                    '   Die Suche läuft, bis das Verzeichnis durchsucht wurde oder jeder Suchbegriff 1x gefunden wurde.
   FSMode_SearchOneFile = 1         ' Suche beenden, sobald einer der Suchbegriffe einen Treffer ergab
   FSMode_SearchAllFolders = 2      ' Die komplette Verzeichnisstruktur durchsuchen -> je Suchbegriff kann es mehrere Treffer geben
End Enum

Private Type SeachItem
   FileName As String
   Matches As Collection
End Type

Private m_SearchMode As FileSearchMode
Private m_SeachItems() As SeachItem
Private m_SearchItemCount As Long
Private m_OpenSeachItems As Collection
Private m_OpenSearchItemCount As Long

Private m_FSO As Scripting.FileSystemObject

'---------------------------------------------------------------------------------------
' Sub: AddSearchItem
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Suchbegriff anfügen
' </summary>
' <param name="FileName">Suchbegriff für den zu suchenen Dateinamen</param>
' <remarks>Wildcards sind möglich, da Textvergleich mit Like durchgeführt wird.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub AddSearchItem(ByVal FileName As String)

   Dim Itm As SeachItem
   
   With Itm
      .FileName = FileName
      Set .Matches = New Collection
   End With
   
   m_SearchItemCount = m_SearchItemCount + 1
   ReDim Preserve m_SeachItems(m_SearchItemCount - 1)
   m_SeachItems(m_SearchItemCount - 1) = Itm
   
End Sub

'---------------------------------------------------------------------------------------
' Function: Search
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Dateisuche starten
' </summary>
' <param name="StartPath">Start-Verzeichnis (String)/param>
' <param name="SearchMode">(Optional) Suchmodus ... Standard = FSMode_SearchOneOfEveryItem (sobald Suchbegriff einmal gefunden wurde, diesen Begriff nicht weiter suchen</param>
' <returns>True wenn mind. 1 Datei gefunden wurde</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function Search(ByVal StartPath As String, Optional ByVal SearchMode As FileSearchMode = FSMode_SearchOneOfEveryItem) As Boolean
   
   Dim StartFolder As Scripting.Folder
   
   m_SearchMode = SearchMode
   
   m_OpenSearchItemCount = m_SearchItemCount

   If m_OpenSearchItemCount = 0 Then
      Err.Raise vbObjectError, "FileSearch.Search", "File search cancelled - 0 items to search"
   End If
   
   Set m_FSO = CreateObject("Scripting.FileSystemObject")
   Set StartFolder = m_FSO.GetFolder(StartPath)
   
   Set m_OpenSeachItems = New Collection
   Dim i As Long
   For i = LBound(m_SeachItems) To UBound(m_SeachItems)
      m_OpenSeachItems.Add i, "Idx" & i
   Next
   
   RunThroughFolders StartFolder
   
   Search = MatchExists

End Function

'---------------------------------------------------------------------------------------
' Property: Matches
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Matches-Collection
' </summary>
' <returns>Collection (String-Item) der gefundenen Dateien</returns>
' <remarks>Mit Matches.Count kann die Anzahl der gefundenen Dateien abgefragt werden.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Matches() As Collection

   Dim Mtchs As Collection
   Dim Itm As SeachItem
   Dim i As Long
   Dim m As Variant
   
   Set Mtchs = New Collection
   
   For i = 0 To UBound(m_SeachItems)
      Itm = m_SeachItems(i)
      If Itm.Matches.Count > 0 Then
         For Each m In Itm.Matches
            Mtchs.Add m
         Next
      End If
   Next
   
   Set Matches = Mtchs

End Property

Private Property Get MatchExists() As Boolean
   
   Dim Itm As SeachItem
   Dim i As Long
   
   For i = 0 To UBound(m_SeachItems)
      If m_SeachItems(i).Matches.Count > 0 Then
         MatchExists = True
         Exit Property
      End If
   Next
   
   MatchExists = False
   
End Property

Private Sub RunThroughFolders(ByVal StartFolder As Object)
   
   Dim SubFolder As Scripting.Folder
   
   RunThroughFolder StartFolder
   If m_OpenSearchItemCount <= 0 Then
      Exit Sub
   End If
   
On Error GoTo HandleErr

   For Each SubFolder In StartFolder.SubFolders
      RunThroughFolders SubFolder
      If m_OpenSearchItemCount <= 0 Then
         Exit Sub
      End If
   Next
   
   Exit Sub
   
HandleErr:
   If Err.Number = 70 Then ' Zugriff verweigert => ignorieren
      Err.Clear
      Exit Sub
   Else
      Err.Raise Err.Number, "RunThroughFolder:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   End If

End Sub

Private Sub RunThroughFolder(ByVal StartFolder As Object)
   
   Dim FileRef As Scripting.File

On Error GoTo HandleErr

   For Each FileRef In StartFolder.Files
      CheckSearchItems FileRef
      If m_OpenSearchItemCount <= 0 Then
         Exit Sub
      End If
   Next
   
   Exit Sub
   
HandleErr:
   If Err.Number = 70 Then ' Zugriff verweigert => ignorieren
      Err.Clear
      Exit Sub
   Else
      Err.Raise Err.Number, "RunThroughFolder:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   End If
   
End Sub

Private Sub CheckSearchItems(ByVal FileRef As Scripting.File)
   
   Dim ItmIndex As Variant
   
   For Each ItmIndex In m_OpenSeachItems
   
      If CheckFileName(m_SeachItems(ItmIndex), FileRef) Then
         If m_SearchMode = FSMode_SearchOneFile Then
            m_OpenSearchItemCount = 0
            Exit Sub
         ElseIf m_SearchMode = FSMode_SearchOneOfEveryItem Then
            m_OpenSeachItems.Remove "Idx" & ItmIndex
            m_OpenSearchItemCount = m_OpenSearchItemCount - 1
            If m_OpenSearchItemCount <= 0 Then
               Exit Sub
            End If
         End If
      End If
   Next
   
End Sub

Private Function CheckFileName(ByRef SearchItm As SeachItem, ByVal FileRef As Scripting.File) As Boolean

   If FileRef.Name Like SearchItm.FileName Then
      With SearchItm.Matches
         .Add FileRef.Path
      End With
      CheckFileName = True
   End If
   
End Function

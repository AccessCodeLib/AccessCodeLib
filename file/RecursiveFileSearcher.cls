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
' Class: file.RecursiveFileSearcher
'---------------------------------------------------------------------------------------
'
' Search a folder structure for one or more files
'
' Remarks:
'     (start code)
'
'     Dim Itm As Variant
'     With New RecursiveFileSearcher
'        .AddSearchItem "msaccess.*"
'        .AddSearchItem "winword.exe"
'
'        If Not .Search("C:\Program Files (x86)\Microsoft Office", FSMode_SearchAllFolders) Then
'           Debug.Print "No match"
'        Else
'           For Each Itm In .Matches
'              Debug.Print Itm
'           Next
'        End If
'     End With
'
'     (end)
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>file/RecursiveFileSearcher.cls</file>
'  <license>_codelib/license.bas</license>
'  <description>Searches a folder tree recursively for one or more file names or patterns (FileSystemObject); modes: stop at first match, one match per item, all matches</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Enum: FileSearchMode
'---------------------------------------------------------------------------------------
'
' Search modes
'
' FSMode_SearchOneOfEveryItem (0) - As soon as a search item has been found, this item is not searched any further. The search runs until the folder has been searched or each search item has been found once.
' FSMode_SearchOneFile (1) - Stop the search as soon as one of the search items returned a match
' FSMode_SearchAllFolders (2) - Search the complete folder structure -> several matches per search item possible
'
'---------------------------------------------------------------------------------------
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
'
' Add search item
'
' Parameters:
'     FileName - Search item for the file name to be searched
'
' Remarks:
'     Wildcards are possible, as the text comparison is done with Like.
'
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
'
' Start file search
'
' Parameters:
'     StartPath  - Start folder (String)
'     SearchMode - (Optional) Search mode ... Default = FSMode_SearchOneOfEveryItem
'
' Returns:
'     True if at least 1 file was found
'
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
'
' Matches collection
'
' Returns:
'     Collection (String items) of the files found
'
' Remarks:
'     Matches.Count returns the number of files found.
'
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

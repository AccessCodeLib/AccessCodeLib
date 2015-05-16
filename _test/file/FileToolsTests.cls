VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileToolsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
'\short     AccUnit-Testklasse für FileTools
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/file/FileToolsTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>file/FileTools.bas</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'--------------------------------------------------------------------
' AccUnit Infrastructure
'--------------------------------------------------------------------

Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row("C:\Windows\A1\A2\A3\", "C:\Windows", ".\A1\A2\A3\").Name = "Forward_EndBackslash"
'AccUnit:Row("C:\Windows\A1\A2\A3", "C:\Windows", ".\A1\A2\A3").Name = "Forward_EndNoBackslash"
'AccUnit:Row("C:\Windows\A1\A2\A3", "C:\Windows\", ".\A1\A2\A3").Name = "Forward_BaseBackslash"
'AccUnit:Row("C:\Windows\A1", "C:\Windows\A1\A2\A3\", ".\..\..\..\A1").Name = "Back_BaseBackslash"
'AccUnit:Row("C:\Windows\A1", "C:\Windows\A1\A2\A3", ".\..\..\..\A1").Name = "Back_BaseNoBackslash"
'AccUnit:Row("C:\Windows\A1\", "C:\Windows\A1\A2\A3\", ".\..\..\").Name = "Back_Backslash"
'AccUnit:Row("C:\Windows\A1", "C:\Windows\A1\A2\A3", ".\..\..\..\A1").Name = "Back_NoBackslash"
'AccUnit:Row("C:\Windows\B1\", "C:\Windows\A1\A2\A3", ".\..\..\..\B1\").Name = "BackAndForward_Backslash"
Public Sub GetRelativPathFromFullPath_Tests(ByVal FullPath As String, ByVal BaseDir As String, ByVal Expected As String)
   Dim Actual As String
   Actual = FileTools.GetRelativPathFromFullPath(FullPath, BaseDir, True)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(".\..\..\..\B1", "C:\Windows\A1\A2\A3\", "C:\Windows\B1").Name = "BackAndForward"
Public Sub GetFullPathFromRelativPath(ByVal RelativePath As String, ByVal BaseDir As String, ByVal Expected As String)
   Dim Actual As String
   Actual = FileTools.GetFullPathFromRelativPath(RelativePath, BaseDir)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("C:\Windows\abc.txt", "abc.txt").Name = "no spaces"
'AccUnit:Row("C:\Windows\ab cd.txt", "ab cd.txt").Name = "spaces"
'AccUnit:Row("C:\Windows\ab\", "").Name = "Dir"
'AccUnit:Row("C:\Windows\ab", "ab").Name = "Dir or File (no ext)"
Public Sub FileNameWithoutPath(ByVal FileAndPath As String, ByVal Expected As String)
   Dim Actual As String
   Actual = FileTools.FileNameWithoutPath(FileAndPath)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("C:\Windows\abc.txt", "C:\Windows\").Name = "no spaces"
'AccUnit:Row("C:\Windows\xyz abc\ab cd.txt", "C:\Windows\xyz abc\").Name = "spaces"
'AccUnit:Row("C:\Windows\ab\", "C:\Windows\ab\").Name = "Dir"
'AccUnit:Row("C:\Windows\ab", "C:\Windows\").Name = "Dir or File (no ext)"
Public Sub GetDirFromFullFileName_FullFileName_Dir(ByVal FileName As String, ByVal Expected As String)
   Dim Actual As String
   Actual = FileTools.GetDirFromFullFileName(FileName)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("C:\Windows\System32\abc.txt", 20, "C:\W...em32\abc.txt").Name = "len 20"
Public Sub ShortenFileName_FullPath_ReturnsShortenedPath(ByVal FileName As String, _
                                 ByVal MaxLen As Long, ByVal Expected As String)
   Dim Actual As String
   Actual = FileTools.ShortenFileName(FileName, MaxLen)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("test.pdf", ".pdf").Name = ".pdf"
'AccUnit:Row("test.PDF", ".PDF").Name = ".PDF"
'AccUnit:Row("c:\test.accdb", ".accdb").Name = "LocalPath"
'AccUnit:Row("\\myserver\myshare\test.ACCDB", ".ACCDB").Name = "UNC"
Public Sub GetFileExtension_FilePath_ReturnsFileExtension(ByVal FilePath As String, ByVal Expected As String)
    Assert.That FileTools.GetFileExtension(FilePath), Iz.EqualTo(Expected)
End Sub

Public Sub FileExists_EmptyPath_ReturnsFalse()
    Assert.IsFalse FileTools.FileExists(vbNullString)
End Sub

Public Sub CreateAndDeleteFolder()
'Issue #109

   Dim TestPath As String
   Dim FSO As Object 'New Scripting.FileSystemObject
   
   Set FSO = CreateObject("Scripting.FileSystemObject")
   
   TestPath = CurrentProject.Path & "\Test_Issue109"

   FileTools.CreateDirectory TestPath
   
   FSO.DeleteFolder TestPath, True
   
   Assert.IsEmpty VBA.Dir(TestPath, vbDirectory)

End Sub

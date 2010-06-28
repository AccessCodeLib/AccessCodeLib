VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "modFiles_FileFolderTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' <TestClass>
'
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/file/modFiles_FileFolderTest.cls</file>
'  <use>file/modFiles.bas</use>
'  <use>test/simplyvbunit/VBATestSuite.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'Aktivierung für TestSuite.RunAll
'<VBATestSuite><testclass>true</testclass></VBATestSuite>

Implements SimplyVBUnit.ITestCaseData

Private Sub ITestCaseData_GetTestData(ByVal Test As SimplyVBUnit.TestDataBuilder)

   Select Case Test.MethodName
      Case "GetRelativPathFromFullPath"
         ' Args: Full, Base, Expected
         Call Test.Use("C:\Windows\A1\A2\A3\", "C:\Windows", ".\A1\A2\A3\").TestName("Forward_EndBackslash")
         Call Test.Use("C:\Windows\A1\A2\A3", "C:\Windows", ".\A1\A2\A3").TestName("Forward_EndNoBackslash")
         Call Test.Use("C:\Windows\A1\A2\A3", "C:\Windows\", ".\A1\A2\A3").TestName("Forward_BaseBackslash")
         
         Call Test.Use("C:\Windows\A1", "C:\Windows\A1\A2\A3\", ".\..\..\..\A1").TestName("Back_BaseBackslash")
         Call Test.Use("C:\Windows\A1", "C:\Windows\A1\A2\A3", ".\..\..\..\A1").TestName("Back_BaseNoBackslash")
         Call Test.Use("C:\Windows\A1\", "C:\Windows\A1\A2\A3\", ".\..\..\").TestName("Back_Backslash")
         Call Test.Use("C:\Windows\A1", "C:\Windows\A1\A2\A3", ".\..\..\..\A1").TestName("Back_NoBackslash")
         
         Call Test.Use("C:\Windows\B1\", "C:\Windows\A1\A2\A3", ".\..\..\..\B1\").TestName("BackAndForward_Backslash")
         
      Case "GetFullPathFromRelativPath"
         ' Args: Relative, Base, Expected
         Call Test.Use(".\..\..\..\B1", "C:\Windows\A1\A2\A3\", "C:\Windows\B1").TestName("BackAndForward")
      
      Case "FileNameWithoutPath"
         ' Args: FullPath, Expected
         Call Test.Use("C:\Windows\abc.txt", "abc.txt").TestName("no spaces")
         Call Test.Use("C:\Windows\ab cd.txt", "ab cd.txt").TestName("spaces")
         Call Test.Use("C:\Windows\ab\", "").TestName("Dir")
         Call Test.Use("C:\Windows\ab", "ab").TestName("Dir or File (no ext)")
      
      Case "GetDirFromFilePath"
         ' Args: File (Fullpath), Expected
         Call Test.Use("C:\Windows\abc.txt", "C:\Windows\").TestName("no spaces")
         Call Test.Use("C:\Windows\xyz abc\ab cd.txt", "C:\Windows\xyz abc\").TestName("spaces")
         Call Test.Use("C:\Windows\ab\", "C:\Windows\ab\").TestName("Dir")
         Call Test.Use("C:\Windows\ab", "C:\Windows\").TestName("Dir or File (no ext)")
      
      Case "ShortFileName"
         ' Args: File (Fullpath), Expected
         Call Test.Use("C:\Windows\System32\abc.txt", 20, "C:\W...em32\abc.txt").TestName("len 20")
      
   End Select

End Sub

Public Sub GetRelativPathFromFullPath(ByVal strFull As String, ByVal strBase As String, ByVal strRelativeExpected As String)
   Dim strRelative As String
   strRelative = modFiles.GetRelativPathFromFullPath(strFull, strBase, True)
   Assert.That strRelative, Iz.EqualTo(strRelativeExpected)
End Sub


Public Sub GetFullPathFromRelativPath(ByVal strRelative As String, ByVal strBase As String, ByVal strFullExpected As String)
   Dim strFull As String
   strFull = modFiles.GetFullPathFromRelativPath(strRelative, strBase)
   Assert.That strFull, Iz.EqualTo(strFullExpected)
End Sub

Public Sub FileNameWithoutPath(ByVal strFileAndPath As String, ByVal strFileNameExpected As String)
   Dim strFileName As String
   strFileName = modFiles.FileNameWithoutPath(strFileAndPath)
   Assert.That strFileName, Iz.EqualTo(strFileNameExpected)
End Sub

Public Sub GetDirFromFilePath(ByVal strFileName As String, ByVal strDirExpected As String)
   Dim strDir As String
   strDir = modFiles.GetDirFromFilePath(strFileName)
   Assert.That strDir, Iz.EqualTo(strDirExpected)
End Sub


Public Sub ShortFileName(ByVal strFileName As String, ByVal lngMaxLen As Long, ByVal strExpected As String)
   Dim strResult As String
   strResult = modFiles.ShortFileName(strFileName, lngMaxLen)
   Assert.That strResult, Iz.EqualTo(strExpected)
End Sub

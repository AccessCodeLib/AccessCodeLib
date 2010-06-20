VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBATestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: VBATestSuite (Josef Pötzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Anpassung (Kapselung) von SimplyVBUnit für VBA
' </summary>
' <remarks>
' Abgestimmt auf die Version 3.0 von SimplyVBUnit
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/_simplyvbunit/VBATestSuite.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_test/_simplyvbunit/_VBATestSuite.bas</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'</codelib>
' Anm.: use(_VBATestSuite.bas) ist eigentlich nicht notwendig, hilft aber beim Import
'                              (falls nur VBATestSuite.cls ausgewählt wurde)
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const EARLYBINDING_VBIDE = 0

#If EARLYBINDING_VBIDE = 1 Then
   Private m_CurrentVbProject As VBIDE.VBProject
#Else
   Private m_CurrentVbProject As Object

   Public Enum vbext_ProcKind
      vbext_pk_Proc = 0
      vbext_pk_Let = 1
      vbext_pk_Set = 2
      vbext_pk_Get = 3
   End Enum

   Public Enum vbext_ComponentType
      vbext_ct_StdModule = 1
      vbext_ct_ClassModule = 2
      vbext_ct_MSForm = 3
      vbext_ct_ActiveXDesigner = 11
      vbext_ct_Document = 100
   End Enum
#End If

'TestSuite (Testerstellung und Durchführung):
Private suite As SimplyVBUnit.TestSuite

'Listener (für Ausgabe)
Private WithEvents listener As SimplyVBUnit.EventCastListener
Attribute listener.VB_VarHelpID = -1

Private m_Passed As Long
Private m_Failed As Long
Private m_Ignored As Long

Private Sub Class_Initialize()
   Set suite = New SimplyVBUnit.TestSuite
   Set listener = New SimplyVBUnit.EventCastListener
End Sub

Public Property Get SimplyVbUnitTestSuite() As SimplyVBUnit.TestSuite
   Set SimplyVbUnitTestSuite = suite
End Property

'---------------------------------------------------------------------------------------
' Function: Run (Josef Pötzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Tests starten
' </summary>
' <param name="bResetSuite">Suite zurücksetzen (SimplyVBUnit.TestSuite.Reset)</param>
' <param name="Filter">ITestFilter</param>
' <returns>Aktuelle Instanz von VBATestSuite</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Run(Optional ByVal bResetSuite As Boolean = True, _
                    Optional ByVal Filter As ITestFilter) As VBATestSuite

   If bResetSuite Then suite.Reset

   Dim Result As SimplyVBUnit.TestResult

   Set Result = suite.Run(listener, Filter)

   Call printSummary(Result)
   
   Set Run = Me
   
End Function


'---------------------------------------------------------------------------------------
' Function: RunAll (Josef Pötzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Tests starten, erzeugt Tests aus Klassenelementen des VB-Projektes
' </summary>
' <param name="bResetAllTests">Tests neue Einlesen (Default = True)</param>
' <returns>Aktuelle Instanz von VBATestSuite</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function RunAll(Optional ByVal bResetAllTests As Boolean = True) As VBATestSuite

   Dim bolResetSuite As Boolean

   If bResetAllTests Then
      Me.Reset True
      Me.AddTestsFromVbProject
   Else
      bolResetSuite = True
   End If

   Set RunAll = Me.Run(bolResetSuite)

End Function

'---------------------------------------------------------------------------------------
' Function: Add (Josef Pötzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Tests (Instanzen von Testklassen) anfügen
' </summary>
' <param name="TestToAdd">ParamArray: durchzuführende Tests (Instanzen von Testklassen)</param>
' <returns>Aktuelle Instanz von VBATestSuite</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Add(ParamArray TestToAdd()) As VBATestSuite

   Dim itm As Variant

   For Each itm In TestToAdd
      suite.Add itm
   Next

   Set Add = Me
End Function

'---------------------------------------------------------------------------------------
' Function: AddTestsFromVbProject (Josef Pötzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Tests aus VB-Projekt erstellen
' </summary>
' <returns>Aktuelle Instanz von VBATestSuite</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function AddTestsFromVbProject() As VBATestSuite
   addTestsFromVbProjectClasses
   Set AddTestsFromVbProject = Me
End Function

'---------------------------------------------------------------------------------------
' Function: Reset (Josef Pötzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' TestSuite zurücksetzen
' </summary>
' <param name="removeTests">True (default): alle eingefügten Tests werden entfernt</param>
' <returns>Aktuelle Instanz von VBATestSuite</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Reset(Optional removeTests As Boolean = False) As VBATestSuite

   If removeTests Then
      Set suite = New SimplyVBUnit.TestSuite
   Else
      suite.Reset
   End If

   Set Reset = Me

End Function


'-----------------------------------------------------
' Ausgabe
'
Private Sub listener_TestSuiteStarted(ByVal testSuiteRef As SimplyVBUnit.TestSuite)
   Debug.Print testSuiteRef.Name '& " Start: " & Now()
End Sub

Private Sub listener_TestFixtureStarted(ByVal Fixture As SimplyVBUnit.TestFixture)
   Debug.Print String(50, "-")
   Debug.Print Fixture.Name
   Debug.Print String(Len(Fixture.Name), "-")
   Debug.Print , "Started: " & Now()
   Debug.Print
End Sub

Private Sub listener_TestCaseFinished(ByVal Result As SimplyVBUnit.TestResult)
   Debug.Print Result.Test.Name,
   If Result.IsSuccess Then
      m_Passed = m_Passed + 1
      Debug.Print "Success",
   ElseIf Result.IsFailure Then
      m_Failed = m_Failed + 1
      Debug.Print "Failure"
   ElseIf Result.IsIgnored Then
      m_Ignored = m_Ignored + 1
      Debug.Print "Ignored",
   ElseIf Result.IsError Then
      m_Failed = m_Failed + 1
      Debug.Print "Error"
      Debug.Print String(2, " ");
   End If
   Debug.Print Result.Message

End Sub

Private Sub listener_TestFixtureFinished(ByVal Result As SimplyVBUnit.TestResult)
   Debug.Print String(5, "-")
   Debug.Print , "Finished: " & Now()
End Sub

Private Sub printSummary(ByVal Result As SimplyVBUnit.TestResult)

   Debug.Print String$(50, "-")
   Debug.Print "Total  : " & Result.Test.TestCount
   Debug.Print String$(10, "-")
   Debug.Print "Passed : " & m_Passed
   Debug.Print "Failed : " & m_Failed
   Debug.Print "Ignored: " & m_Ignored
   Debug.Print String$(9 + Len(CStr(Result.Test.TestCount)), "-")
   Debug.Print "Time   : " & Round(Result.Time, 0) & " ms"
   Debug.Print String$(50, "-")

End Sub

'-----------------------------------------------------
' Test-Instanzen aus Klassennamen erzeugen
'
Private Sub addTestsFromVbProjectClasses()

#If EARLYBINDING_VBIDE Then
   Dim testFactoryVbProject As VBIDE.VBProject
   Dim testFactoryVbComponent As VBComponent
   Dim testFactoryCodeModule As CodeModule

   Dim checkVbComponent As VBComponent
   Dim checkCodeModule As CodeModule
#Else
   Dim testFactoryVbProject As Object
   Dim testFactoryVbComponent As Object
   Dim testFactoryCodeModule As Object

   Dim checkVbComponent As Object
   Dim checkCodeModule As Object
#End If

   Set testFactoryVbProject = CurrentVbProject
   Set testFactoryVbComponent = testFactoryVbProject.VBComponents.Add(vbext_ct_StdModule)
   Set testFactoryCodeModule = testFactoryVbComponent.CodeModule

   Dim checkString As String
   For Each checkVbComponent In testFactoryVbProject.VBComponents
      If checkVbComponent.Type = vbext_ct_ClassModule Then
         ' find:
         Set checkCodeModule = checkVbComponent.CodeModule
         checkString = findSubString(checkCodeModule.Lines(1, checkCodeModule.CountOfLines), "<VBATestSuite>", "</VBATestSuite>")
         If Len(checkString) > 0 Then
            If StrComp("true", findSubString(checkString, "<testclass>", "</testclass>"), vbTextCompare) = 0 Then
               addTestClassFromClassName testFactoryCodeModule, checkVbComponent.Name
            End If
         End If
      End If
   Next

   testFactoryVbProject.VBComponents.Remove testFactoryVbComponent

End Sub

#If EARLYBINDING_VBIDE Then
Private Sub addTestClassFromClassName(testFactoryCodeModule As CodeModule, className As String)
#Else
Private Sub addTestClassFromClassName(testFactoryCodeModule As Object, className As String)
#End If

   Dim procString As String
   procString = "public sub VBATestSuite_" & className & "(byref testObj as object)" & vbNewLine & _
                "   set testObj = new " & className & vbNewLine & _
                "end sub" & vbNewLine

   testFactoryCodeModule.InsertLines testFactoryCodeModule.CountOfLines + 1, procString

   Dim newTestFromClassName As Object
   Application.Run "VBATestSuite_" & className, newTestFromClassName

   Me.Add newTestFromClassName

   testFactoryCodeModule.DeleteLines 1, testFactoryCodeModule.CountOfLines

End Sub

Private Function findSubString(ByRef SourceString As String, _
                               ByRef StartTag As String, ByRef EndTag As String, _
                               Optional ByRef LocatedStartPos As Long, Optional ByRef LocatedEndPos As Long) As String

   Dim startPos As Long, endPos As Long

   startPos = InStr(1, SourceString, StartTag, vbTextCompare)
   If startPos <= 0 Then
      Exit Function
   End If

   startPos = startPos + Len(StartTag)
   endPos = InStr(startPos, SourceString, EndTag, vbTextCompare)
   If endPos <= 0 Then
      Exit Function
   End If

   LocatedStartPos = startPos
   LocatedEndPos = endPos
   findSubString = Trim$(Mid$(SourceString, startPos, endPos - startPos))

End Function

#If EARLYBINDING_VBIDE Then
Public Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Public Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING_VBIDE Then
   Dim proj As VBProject
#Else
   Dim proj As Object
#End If
   Dim strCurrentDbName As String

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = VBE.ActiveVBProject
      'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
      strCurrentDbName = CurrentDb.Name
      If m_CurrentVbProject.FileName <> strCurrentDbName Then
         Set m_CurrentVbProject = Nothing
         For Each proj In VBE.VBProjects
            If proj.FileName = strCurrentDbName Then
               Set m_CurrentVbProject = proj
               Exit For
            End If
         Next
      End If
   End If

   Set CurrentVbProject = m_CurrentVbProject

End Property

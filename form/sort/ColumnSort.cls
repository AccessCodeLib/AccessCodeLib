VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: ColumnSort
'---------------------------------------------------------------------------------------
'/**
' \author       Andreas Vogt - nach einer Vorlage von Armin Hess (http://www.gaeuwetter.de/access/home/)
' \brief        Klasse zur Sortierung von Endlosformularen mittels CommandButtons
' \ingroup usability
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/sort/ColumnSort.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Enum SortMode
    Unsort = 0
    Asc = 1
    DESC = 2
End Enum

Private m_Form As Form
Private m_cmdBtn As CommandButton
Private m_cmdBtnOld As CommandButton
Private m_Modus As SortMode
Private m_FieldName As String
Private m_AscIcon As String
Private m_DescIcon As String

Public Property Let SortFieldName(ByVal cFieldName As String)
    m_FieldName = cFieldName
End Property

Public Property Get SortFieldName() As String
    SortFieldName = m_FieldName
End Property

Public Sub Sort()
    Set m_Form = Screen.ActiveForm
    If m_Form.RecordsetClone.RecordCount = 0 Then
        Set m_Form = Nothing
        Err.Raise vbObjectError + 99, "clsSort.Sort", "no records to sort"
    End If

    Set m_cmdBtn = Screen.ActiveControl
    updateModus m_Modus
    
    'bei Wechsel des Buttons den vorherigen Button zurücksetzen und Standard-Modus ASC für neuen Button setzen
    If Not m_cmdBtnOld Is Nothing Then
        If m_cmdBtn.Name <> m_cmdBtnOld.Name Then
            setIcon m_cmdBtnOld, 0
            m_Modus = Asc
        End If
    End If

    'aktuellen Button konfigurieren
    setIcon m_cmdBtn, m_Modus
    Set m_cmdBtnOld = m_cmdBtn

    'Wenn kein Sortierfeld angegeben Fehler auswerfen
    If SortFieldName = vbNullString Then
        Err.Raise vbObjectError + 100, "clsSort.Sort", "no selected field to sort"
    End If

    m_Form.OrderBy = CreateSortString(m_Modus)
    m_Form.OrderByOn = True
End Sub

Private Function CreateSortString(ByVal Mode As SortMode) As String
    Dim cSortMode As String
    
    If Mode <> Unsort Then
        If Mode = DESC Then cSortMode = " DESC"
        CreateSortString = "[" & SortFieldName & "]" & cSortMode
    End If
End Function

Private Sub setIcon(ByRef cmdBtn As CommandButton, ByVal Mode As SortMode)
    If AscW(Right$(cmdBtn.Caption, 1)) > &H2500 Then
        cmdBtn.Caption = Left$(cmdBtn.Caption, Len(cmdBtn.Caption) - 2)
    End If

    Select Case Mode
    Case DESC
        cmdBtn.Caption = cmdBtn.Caption & " " & m_DescIcon
    Case Asc
        cmdBtn.Caption = cmdBtn.Caption & " " & m_AscIcon
    End Select
End Sub

Private Sub updateModus(ByRef SortModus As Long)
    Select Case SortModus
    Case 0, 1
        SortModus = SortModus + 1
    Case 2
        SortModus = 0
    End Select
End Sub

Private Sub Class_Initialize()
    m_AscIcon = ChrW(&H25B2)
    m_DescIcon = ChrW(&H25BC)
End Sub

Private Sub Class_Terminate()
    If Not m_Form Is Nothing Then m_Form.OrderByOn = False
End Sub

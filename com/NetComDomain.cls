VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NetComDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: com.NetComDomain
'---------------------------------------------------------------------------------------
'
' Use a .NET COM dll without registration
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>com/NetComDomain.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>mscoree</name><major>2</major><minor>0</minor><guid>{5477469E-83B1-11D2-8B49-00A0C9B7C9C4}</guid></ref>
'  <ref><name>mscorlib</name><major>2</major><minor>0</minor><guid>{BED7F4EA-1A96-11D2-8F08-00A0C9A6186D}</guid></ref>
'  <example>_examples/com/NetComDomain_Example_ShowWinForm_Minimal.bas</example>
'  <description>Creates objects from a .NET COM library without registration (loads the dll into an AppDomain via mscoree)</description>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private m_LibraryFile As String
Private m_DefaultNameSpace As String

'---------------------------------------------------------------------------------------
' Function: CreateObject
'---------------------------------------------------------------------------------------
'
' Creates a new instance of a class from a .NET COM dll
'
' Parameters:
'     ClassIdentifier  - Name of the class without namespace
'     DefaultNameSpace - Namespace containing the class
'     LibraryFile      - Absolute path of the .NET COM dll
'
' Returns:
'     Object
'
'---------------------------------------------------------------------------------------
Public Function CreateObject(ByVal ClassIdentifier As String, _
                    Optional ByVal DefaultNameSpace As String = vbNullString, _
                    Optional ByVal LibraryFile As String = vbNullString) As Object

   If VBA.Len(DefaultNameSpace) > 0 Then Me.DefaultNameSpace = DefaultNameSpace
   If VBA.Len(LibraryFile) > 0 Then Me.LibraryFile = LibraryFile
   
   Dim NetRTH As mscoree.CorRuntimeHost
   Dim NetAppDomain As mscorlib.AppDomain

   If VBA.InStr(1, ClassIdentifier, ".") = 0 Then
      If VBA.Len(Me.DefaultNameSpace) = 0 Then
         Err.Raise vbObjectError, "NetComDomain.CreateObject", "Namespace muss angegeben werden!"
      End If
      ClassIdentifier = Me.DefaultNameSpace & "." & ClassIdentifier
   End If

   Set NetRTH = New mscoree.CorRuntimeHost
   NetRTH.Start
   NetRTH.GetDefaultDomain NetAppDomain
   
   Set CreateObject = NetAppDomain.CreateInstanceFrom(Me.LibraryFile, ClassIdentifier).Unwrap

   NetRTH.Stop
   Set NetAppDomain = Nothing
   Set NetRTH = Nothing

End Function

Public Property Get LibraryFile() As String
    LibraryFile = m_LibraryFile
End Property

Public Property Let LibraryFile(ByVal FilePath As String)
    m_LibraryFile = FilePath
End Property

Public Property Get DefaultNameSpace() As String
    DefaultNameSpace = m_DefaultNameSpace
End Property

Public Property Let DefaultNameSpace(ByVal NameSpace As String)
    m_DefaultNameSpace = NameSpace
End Property

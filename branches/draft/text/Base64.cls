VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Base64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: Base64
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Base64 Kodierung und Dekodierung
' \ingroup text
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/Base64.cls</file>
'  <test>_test\text\Base64Tests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Private m_xmlDocument As Object
Private m_xmlNode As Object

'---------------------------------------------------------------------------------------
' Enum: DecodeOption
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Verfügbare Optionen für Decode
' </summary>
' <list type="table">
'   <item><term>None (1)</term><description>Keine Konvertierung</description></item>
'   <item><term>FromUnicode (2)</term><description>Von Unicode zum Systemzeichensatz</description></item>
'   <item><term>ToUnicode (3)</term><description>Vom Systemzeichensatz nach Unicode</description></item>
' </list>
'**/                                            '<-- Ende Doxygen-Block
'---------------------------------------------------------------------------------------
Public Enum Base64ConvertOption
    None
    FromUnicode
    ToUnicode
End Enum

Private Property Get xmlNode() As Object

    If m_xmlNode Is Nothing Then
        Set m_xmlDocument = CreateObject("MSXML2.DOMDocument")
        Set m_xmlNode = m_xmlDocument.createElement("b64")
            m_xmlNode.DataType = "bin.base64"
    End If
    
    Set xmlNode = m_xmlNode
End Property

'---------------------------------------------------------------------------------------
' Function: Encode
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Kodiert den Eingabestring mittels Base64
' </summary>
' <param name="Value">String der kodiert werden soll</param>
' <returns>kodierter Base64 String</returns>
' <remarks>
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
'---------------------------------------------------------------------------------------
Public Function Encode(ByVal Value As String, Optional ByVal ConvertOption As Base64ConvertOption = Base64ConvertOption.None) As String
    
    Dim byteArray() As Byte
    
    Select Case ConvertOption
        Case Base64ConvertOption.None
            byteArray = Value
        Case Base64ConvertOption.FromUnicode
            byteArray = VBA.Strings.StrConv(Value, vbFromUnicode)
        Case Base64ConvertOption.ToUnicode
            byteArray = VBA.Strings.StrConv(Value, vbUnicode)
    End Select
    
    xmlNode().nodeTypedValue = byteArray
    Encode = xmlNode().Text

End Function

'---------------------------------------------------------------------------------------
' Function: DecodeToByteArray
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Dekodiert einen Base64 kodierten String
' </summary>
' <param name="Value">String der dekodiert werden soll</param>
' <returns>dekodierter String als Byte()</returns>
' <remarks>
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
'---------------------------------------------------------------------------------------
Public Function DecodeToByteArray(ByVal Value As String) As Byte()

    xmlNode().Text = Value
    DecodeToByteArray = xmlNode().nodeTypedValue

End Function

'---------------------------------------------------------------------------------------
' Function: Decode
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Dekodiert einen Base64 kodierten String
' </summary>
' <param name="Value">String der dekodiert werden soll</param>
' <param name="ConvertOption">Konvertierungsoptionen</param>
' <returns>dekodierter String</returns>
' <remarks>
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
'---------------------------------------------------------------------------------------
Public Function Decode(ByVal Value As String, Optional ByVal ConvertOption As Base64ConvertOption = Base64ConvertOption.None) As String
    
    If Value = vbNullString Then
        Decode = vbNullString
        Exit Function
    End If
    
    Select Case ConvertOption
        Case Base64ConvertOption.None
            Decode = Base64.DecodeToByteArray(Value)
            Exit Function
        Case Base64ConvertOption.FromUnicode
            Decode = VBA.Strings.StrConv(Base64.DecodeToByteArray(Value), vbFromUnicode)
            Exit Function
        Case Base64ConvertOption.ToUnicode
            Decode = VBA.Strings.StrConv(Base64.DecodeToByteArray(Value), vbUnicode)
            Exit Function
    End Select
    
End Function

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
' Class: text.Base64
'---------------------------------------------------------------------------------------
'
' Base64 encoding and decoding
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/Base64.cls</file>
'  <test>_test/text/Base64Tests.cls</test>
'  <license>_codelib/license.bas</license>
'  <description>Encodes strings to Base64 and decodes Base64 to strings or byte arrays (via MSXML2.DOMDocument), optional Unicode conversion</description>
'</codelib>
'---------------------------------------------------------------------------------------

Private m_xmlDocument As Object
Private m_xmlNode As Object

'---------------------------------------------------------------------------------------
' Enum: Base64ConvertOption
'---------------------------------------------------------------------------------------
'
' Available conversion options
'
' None (1) - No conversion
' FromUnicode (2) - From Unicode to the system character set
' ToUnicode (3) - From the system character set to Unicode
'
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
'
' Encodes the input string with Base64
'
' Parameters:
'     Value         - String to be encoded
'     ConvertOption - Conversion option
'
' Returns:
'     Base64 encoded string
'
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
'
' Decodes a Base64 encoded string
'
' Parameters:
'     Value - String to be decoded
'
' Returns:
'     Decoded string as Byte()
'
'---------------------------------------------------------------------------------------
Public Function DecodeToByteArray(ByVal Value As String) As Byte()

    xmlNode().Text = Value
    DecodeToByteArray = xmlNode().nodeTypedValue

End Function

'---------------------------------------------------------------------------------------
' Function: Decode
'---------------------------------------------------------------------------------------
'
' Decodes a Base64 encoded string
'
' Parameters:
'     Value         - String to be decoded
'     ConvertOption - Conversion option
'
' Returns:
'     Decoded string
'
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

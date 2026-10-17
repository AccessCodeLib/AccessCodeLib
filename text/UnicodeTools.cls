VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnicodeTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: text.UnicodeTools
'---------------------------------------------------------------------------------------
'
' (EXPERIMENTAL)
' Tools for detecting and processing Unicode strings
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/UnicodeTools.cls</file>
'  <test>_test/text/UnicodeToolsTests.cls</test>
'  <license>_codelib/license.bas</license>
'  <description>Experimental: checks whether a string contains Unicode text (WinAPI IsTextUnicode)</description>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Enum: IS_TEXT_UNICODE
'---------------------------------------------------------------------------------------
'
' Unicode test options to pass to the WinAPI call IsTextUnicode()
'
' ASCII16 (1) -  The text is Unicode, and contains only zero-extended ASCII values/characters.
' STATISTICS (2) -  The text is probably Unicode, with the determination made by applying statistical analysis. Absolute certainty is not guaranteed. See the Remarks section.
' CONTROLS (4) -  The text contains Unicode representations of one or more of these nonprinting characters: RETURN, LINEFEED, SPACE, CJK_SPACE, TAB.
' SIGNATURE (8) -  The text contains the Unicode byte-order mark (BOM) 0xFEFF as its first character.
' UNICODE_MASK (15) -  The value is a combination of ASCII16, STATISTICS, CONTROLS, SIGNATURE.
' REVERSE_ASCII16 (16) -  Same as ASCII16, except that the Unicode text is byte-reversed.
' REVERSE_STATISTICS (32) -  Same as STATISTICS, except that the text that is probably Unicode is byte-reversed.
' REVERSE_CONTROLS (64) -  Same as CONTROLS, except that the Unicode characters are byte-reversed.
' REVERSE_SIGNATURE (128) -  The text contains the Unicode byte-reversed byte-order mark (Reverse BOM) 0xFFFE as its first character.
' REVERSE_MASK (240) -  The value is a combination of REVERSE_ASCII16, REVERSE_STATISTICS, REVERSE_CONTROLS, REVERSE_SIGNATURE.
' ILLEGAL_CHARS (256) -  The text contains one of these Unicode-illegal characters: embedded Reverse BOM, UNICODE_NUL, CRLF (packed into one word), or 0xFFFF.
' ODD_LENGTH (512) -  The number of characters in the string is odd. A string of odd length cannot (by definition) be Unicode text.
' DBCS_LEADBYTE (1024) - (not documented yet)
' NOT_UNICODE_MASK (3840) -  The value is a combination of ILLEGAL_CHARS, ODD_LENGTH, and two currently unused bit flags.
' NULL_BYTES (4096) -  The text contains null bytes, which indicate non-ASCII text.
' NOT_ASCII_MASK (61440) -  The value is a combination of NULL_BYTES and three currently unused bit flags.
'
'---------------------------------------------------------------------------------------
Public Enum IS_TEXT_UNICODE
    ASCII16 = 1 '&H1
    STATISTICS = 2 '&H2
    Controls = 4 '&H4
    SIGNATURE = 8 '&H8
    UNICODE_MASK = 15 '&HF
    REVERSE_ASCII16 = 16 '&H10
    REVERSE_STATISTICS = 32 '&H20
    REVERSE_CONTROLS = 64 '&H40
    REVERSE_SIGNATURE = 128 '&H80
    REVERSE_MASK = 240 '&HF0
    ILLEGAL_CHARS = 256 '&H100
    ODD_LENGTH = 512 '&H200
    DBCS_LEADBYTE = 1024 '&H400
    NOT_UNICODE_MASK = 3840 '&HF00
    NULL_BYTES = 4096 '&H1000
    NOT_ASCII_MASK = 61440  '&HF000 '-4096 '???
End Enum

'---------------------------------------------------------------------------------------
' Function: IsTextUnicode
'---------------------------------------------------------------------------------------
'
' Determines if a buffer is likely to contain a form of Unicode text.
'
' Parameters:
'     lpv       - [in] Pointer to the input buffer to examine.
'     iSize     - [in] Size, in bytes, of the input buffer indicated by lpv.
'     lpiResult - [in, out, optional] On input, pointer to the tests to apply to the input buffer text.
'
' Returns:
'     Boolean
'
' Remarks:
'     http://msdn.microsoft.com/en-us/library/windows/desktop/dd318672%28v=vs.85%29.aspx
'
'---------------------------------------------------------------------------------------
Private Declare PtrSafe Function IsTextUnicode Lib "advapi32" (ByRef lpv As Any, ByVal iSize As Long, ByRef lpiResult As Long) As Boolean

'---------------------------------------------------------------------------------------
' Function: IsUnicode
'---------------------------------------------------------------------------------------
'
' Determines by means of various test options whether a string is in the Unicode character set.
'
' Parameters:
'     Value                  - String to test
'     TestOption             - Test option
'     RaiseErrorOnFailedTest - Determines whether an error is raised if a test fails and
'                              the return value is therefore ambiguous.
'
' Returns:
'     Boolean
'
' Remarks:
'     Based on the WinAPI call 'IsTextUnicode()', see
'     http://msdn.microsoft.com/en-us/library/windows/desktop/dd318672%28v=vs.85%29.aspx
'
'---------------------------------------------------------------------------------------
Public Property Get IsUnicode(ByRef Value As String, _
                              Optional ByVal TestOption As IS_TEXT_UNICODE = IS_TEXT_UNICODE.UNICODE_MASK, _
                              Optional ByVal RaiseErrorOnFailedTest As Boolean = False) As Boolean
        
    Dim lpiResult As Long
        lpiResult = VBA.Conversion.CLng(TestOption)
    
    IsUnicode = (IsTextUnicode(Value, VBA.Strings.LenB(Value) * 2, lpiResult) And lpiResult <> 0)
    
    If RaiseErrorOnFailedTest And lpiResult = 0 Then Err.Raise vbObjectError, "IsUnicode", "IsTextUnicode Test failed"
        
End Property

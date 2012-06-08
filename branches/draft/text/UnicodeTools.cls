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
' Class Module: UnicodeTools (EXPERIMENTAL)
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Tools zur Erkennung und Verarbeitung von Unicode-Zeichenketten
' \ingroup text
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/UnicodeTools.cls</file>
'  <test>_test\text\UnicodeToolsTests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Enum: IS_TEXT_UNICODE
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Unicode Testverfahren zur Übergabe an den WinAPI Call IsTextUnicode()
' </summary>
' <list type="table">
'   <item><term>ASCII16 (1)</term><description>The text is Unicode, and contains only zero-extended ASCII values/characters.</description></item>
'   <item><term>STATISTICS (2)</term><description>The text is probably Unicode, with the determination made by applying statistical analysis. Absolute certainty is not guaranteed. See the Remarks section.</description></item>
'   <item><term>CONTROLS (4)</term><description>The text contains Unicode representations of one or more of these nonprinting characters: RETURN, LINEFEED, SPACE, CJK_SPACE, TAB.</description></item>
'   <item><term>SIGNATURE (8)</term><description>The text contains the Unicode byte-order mark (BOM) 0xFEFF as its first character.</description></item>
'   <item><term>UNICODE_MASK (15)</term><description>The value is a combination of IS_TEXT_UNICODE_ASCII16, IS_TEXT_UNICODE_STATISTICS, IS_TEXT_UNICODE_CONTROLS, IS_TEXT_UNICODE_SIGNATURE.</description></item>
'   <item><term>REVERSE_ASCII16 (16)</term><description>Same as ASCII16, except that the Unicode text is byte-reversed.</description></item>
'   <item><term>REVERSE_STATISTICS (32)</term><description>Same as STATISTICS, except that the text that is probably Unicode is byte-reversed.</description></item>
'   <item><term>REVERSE_CONTROLS (64)</term><description>Same as CONTROLS, except that the Unicode characters are byte-reversed.</description></item>
'   <item><term>REVERSE_SIGNATURE (128)</term><description>The text contains the Unicode byte-reversed byte-order mark (Reverse BOM) 0xFFFE as its first character.</description></item>
'   <item><term>REVERSE_MASK (240)</term><description>The value is a combination of IS_TEXT_UNICODE_REVERSE_ASCII16, IS_TEXT_UNICODE_REVERSE_STATISTICS, IS_TEXT_UNICODE_REVERSE_CONTROLS, IS_TEXT_UNICODE_REVERSE_SIGNATURE.</description></item>
'   <item><term>ILLEGAL_CHARS (256)</term><description>The text contains one of these Unicode-illegal characters: embedded Reverse BOM, UNICODE_NUL, CRLF (packed into one word), or 0xFFFF.</description></item>
'   <item><term>ODD_LENGTH (512)</term><description>The number of characters in the string is odd. A string of odd length cannot (by definition) be Unicode text.</description></item>
'   <item><term>DBCS_LEADBYTE (1024)</term><description>??? TODO ???</description></item>
'   <item><term>NOT_UNICODE_MASK (3840)</term><description>The value is a combination of IS_TEXT_UNICODE_ILLEGAL_CHARS, IS_TEXT_UNICODE_ODD_LENGTH, and two currently unused bit flags.</description></item>
'   <item><term>NULL_BYTES (4096)</term><description>The text contains null bytes, which indicate non-ASCII text.</description></item>
'   <item><term>NOT_ASCII_MASK (61440)</term><description>The value is a combination of IS_TEXT_UNICODE_NULL_BYTES and three currently unused bit flags.</description></item>
' </list>
'**/                                            '<-- Ende Doxygen-Block
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
'/**
' <summary>
' Determines if a buffer is likely to contain a form of Unicode text.
' </summary>
' <param name="lpv">[in] Pointer to the input buffer to examine.</param>
' <param name="iSize">[in] Size, in bytes, of the input buffer indicated by lpv.</param>
' <param name="iSize">[in, out, optional] On input, pointer to the tests to apply to the input buffer text.</param>
' <returns>Boolean</returns>
' <remarks>
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd318672%28v=vs.85%29.aspx
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Private Declare Function IsTextUnicode Lib "advapi32" (ByRef lpv As Any, ByVal iSize As Long, ByRef lpiResult As Long) As Boolean

'---------------------------------------------------------------------------------------
' Function: IsUnicode
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt anhand verschiedener Testoptionen, ob ein String im Unicode-Zeichensatz vorliegt.
' </summary>
' <param name="Value">Prüfstring</param>
' <param name="TestOption">Prüf- bzw. Testverfahren</param>
' <param name="RaiseErrorOnFailedTest">Legt fest ob ein Fehler generiert werden soll, wenn ein Test fehlschlägt und der Return-Wert darurch nicht eindeutig ist.</param>
' <returns>Boolean</returns>
' <remarks>
' Basiert auf dem WinAPI-Call 'IsTextUnicode()', siehe
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd318672%28v=vs.85%29.aspx
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get IsUnicode(ByRef Value As String, _
                              Optional ByVal TestOption As IS_TEXT_UNICODE = IS_TEXT_UNICODE.UNICODE_MASK, _
                              Optional ByVal RaiseErrorOnFailedTest As Boolean = False) As Boolean
        
    Dim lpiResult As Long
        lpiResult = VBA.Conversion.CLng(TestOption)
    
    IsUnicode = (IsTextUnicode(Value, VBA.Strings.LenB(Value) * 2, lpiResult) And lpiResult <> 0)
    
    If RaiseErrorOnFailedTest And lpiResult = 0 Then Err.Raise vbObjectError, "IsUnicode", "IsTextUnicode Test failed"
        
End Property

Attribute VB_Name = "modDirectCOM"
'---------------------------------------------------------------------------------------
' Module: modDirectCOM
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Instanz aus unregistrierter COM-dll erzeugen
' </summary>
' <remarks>
' Die Funktion GetInstance ist hilfreich, wenn auf eine COM-dll nur mittel Late Binding
' zugegriffen wird. Die dll muss nicht mittels regsvr32 registriert worden sein.
' </remarks>
' \ingroup COM
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>com/modDirectCOM.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Aufruf siehe Function GetInstance(ByVal LibraryString As String, ByVal ProgIdString As String) As stdole.IUnknown
'
Option Compare Text
Option Explicit

'Zuordnung der Prozeduren zur Doxygen-Gruppe:
'/** \addtogroup COM
'@{ **/

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal adr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal szLib As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal szFnc As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal szModule As String) As Long
Private Declare Function LoadTypeLib Lib "oleaut32" (ByVal szFile As Long, pptlib As Any) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (tGUID As Any, ByVal lpszString As String, ByVal lMax As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function ProgIDFromCLSID Lib "ole32" (pCLSID As Any, lpszProgID As Long) As Long
Private Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal dlen As Long)

Private Type IUnknown2
    QueryInterface As Long
    AddRef As Long
    Release As Long
End Type

Private Type IClassFactory
    IUnk As IUnknown2
    CreateInstance As Long
    Lock As Long
End Type

Private Type ITypeInfo
    IUnk As IUnknown2
    GetTypeAttr As Long
    GetTypeComp As Long
    GetFuncDesc As Long
    GetVarDesc As Long
    GetNames As Long
    GetRefTypeOfImplType As Long
    GetImplTypeFlags As Long
    GetIDsOfNames As Long
    Invoke As Long
    GetDocumentation As Long
    GetDllEntry As Long
    GetRefTypeInfo As Long
    AddressOfMember As Long
    CreateInstance As Long
    GetMops As Long
    GetContainingTypeLib As Long
    ReleaseTypeAttr As Long
    ReleaseFuncDesc As Long
    ReleaseVarDesc As Long
End Type

Private Type ITypeLib
    IUnk As IUnknown2
    GetTypeInfoCount As Long
    GetTypeInfo As Long
    GetTypeInfoType As Long
    GetTypeInfoOfGuid As Long
    GetLibAttr As Long
    GetTypeComp As Long
    GetDocumentation As Long
    IsName As Long
    FindName As Long
    ReleaseTLibAttr As Long
End Type

Private Type TYPEATTR
    GUID(15) As Byte
    lcid As Long
    dwReserved As Long
    memidConstructor As Long
    memidDestructor As Long
    pstrSchema As Long
    cbSizeInstance As Long
    TYPEKIND As Long
    cFuncs As Integer
    cVars As Integer
    cImplTypes As Integer
    cbSizeVft As Integer
    cbAlignment As Integer
    wTypeFlags As Integer
    wMajorVerNum As Integer
    wMinorVerNum As Integer
    tdescAlias As Long
    idldescType As Long
End Type

Private Enum TYPEKIND
    TKIND_ENUM
    TKIND_RECORD
    TKIND_MODULE
    TKIND_INTERFACE
    TKIND_DISPATCH
    TKIND_COCLASS
    TKIND_ALIAS
    TKIND_UNION
    TKIND_MAX
End Enum

Private Enum HRESULT
    S_OK = 0
End Enum

Private Type CoClass
    Name As String
    prgid As String
    GUID() As Byte
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

'---------------------------------------------------------------------------------------
' Function: GetInstance (2009-12-01)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt eine Instanz aus der Klasse in einer COM-dll. Die dll muss nicht registriert sein.
' </summary>
' <param name="LibraryString">Vollstaendiger Pfad zur COM-dll</param>
' <param name="ProgIdString">Name der Klasse in der COM-dll</param>
' <returns>Instanz der Klasse</returns>
' <remarks>
' Beispiel: \n
' set obj = GetInstance(CurrentProject.Path & "\EineCombibliothek.dll", "XYZ") \n
' Damit wird eine Instanz der Klasse "XYZ" aus der dll (die sich im mdb/accdb-Verzeichnis) befindet erzeugt.
' \n\n
' Prozedur im Internet gefunden: http://www.ppreview.net/Forum/post.asp?method=ReplyQuote&REPLY_ID=846&TOPIC_ID=570&FORUM_ID=2 \n
' Code angepasst
' </remarks>
'  <ref><name>stdole</name><major>2</major><minor>0</minor><guid>{00020430-0000-0000-C000-000000000046}</guid></ref>
'**/
'---------------------------------------------------------------------------------------
Public Function getInstance(ByVal LibraryString As String, ByVal ProgIdString As String) As stdole.IUnknown
    Dim newobj As stdole.IUnknown
    Dim TFactory As IClassFactory

    Dim classid As GUID
    Dim IID_ClassFactory As GUID
    Dim IID_IUnknown As GUID
    Dim lib As String

    Dim obj As Long
    Dim vtbl As Long

    Dim hModule As Long
    Dim pFunc As Long
    Dim arrTCoClass() As CoClass

    Dim i As Long, n As Long
    Dim flag As Boolean
    Dim strClassname As String
    

    n = InStr(1, ProgIdString, ".")
    If n > 0 Then
        strClassname = Mid(ProgIdString, n + 1)
        n = InStr(1, strClassname, ".")
        If n > 0 Then strClassname = Left(ProgIdString, n - 1)
    Else
        strClassname = ProgIdString
    End If

    With IID_ClassFactory
        .Data1 = &H1
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With IID_IUnknown
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    ' get all CoClasses from the type lib of the file, and find the GUID of the Prog ID
    If Not getCoClasses(LibraryString, arrTCoClass) Then
        Exit Function
    End If

    For i = 0 To UBound(arrTCoClass)
#If DEBUGMODE = 1 Then
   Debug.Print arrTCoClass(i).prgid, arrTCoClass(i).Name, StringFromGUID(arrTCoClass(i).GUID)
#End If
        If Len(arrTCoClass(i).prgid) > 0 Then
            If StrComp(arrTCoClass(i).prgid, ProgIdString, vbTextCompare) = 0 Then
                flag = True
            Else
                If StrComp(arrTCoClass(i).Name, strClassname, vbTextCompare) = 0 Then flag = True
            End If
        Else
            If StrComp(arrTCoClass(i).Name, strClassname, vbTextCompare) = 0 Then flag = True
        End If
        If flag Then
            CpyMem classid, arrTCoClass(i).GUID(0), Len(classid)
            Exit For
        End If
    Next i

    If i = UBound(arrTCoClass) + 1 Then Exit Function

    ' load the file, if it isn't yet
    hModule = GetModuleHandle(LibraryString)
    If hModule = 0 Then
        hModule = LoadLibrary(LibraryString)
        If hModule = 0 Then Exit Function
    End If

    pFunc = GetProcAddress(hModule, "DllGetClassObject")
    If pFunc = 0 Then Exit Function

    ' call DllGetClassObject to get a class factory for the class ID
    If 0 <> callPointer(pFunc, VarPtr(classid), VarPtr(IID_ClassFactory), VarPtr(obj)) Then Exit Function

    ' IClassFactory VTable
    CpyMem vtbl, ByVal obj, 4
    CpyMem TFactory, ByVal vtbl, Len(TFactory)

    ' create an instance of the object
    If 0 <> callPointer(TFactory.CreateInstance, obj, 0, VarPtr(IID_IUnknown), VarPtr(newobj)) Then
        ' Set IClassFactory = Nothing
        callPointer TFactory.IUnk.Release, obj
        Exit Function
    End If

    ' Set IClassFactory = Nothing
    callPointer TFactory.IUnk.Release, obj

    Set getInstance = newobj

   
End Function

Private Function getCoClasses(ByVal strFile As String, udtCoClasses() As CoClass) As Boolean

    Dim hRes As HRESULT

    Dim TTypeLib As ITypeLib
    Dim TTypeInfo As ITypeInfo
    Dim TTypeAttr As TYPEATTR

    Dim ptrTypeLib As Long
    Dim ptrTypeInfo As Long
    Dim pVTbl As Long
    Dim pAttr As Long

    Dim lngTypeInfos As Long
    Dim lngCoCls As Long
    Dim strTypeName As String

    Dim i As Long

    ' load the type lib of the file
    hRes = LoadTypeLib(StrPtr(strFile), ptrTypeLib)
    If hRes <> S_OK Then Exit Function

    ' ITypeLib's VTable
    CpyMem pVTbl, ByVal ptrTypeLib, 4
    CpyMem TTypeLib, ByVal pVTbl, Len(TTypeLib)

    lngTypeInfos = callPointer(TTypeLib.GetTypeInfoCount, ptrTypeLib)

    For i = 0 To lngTypeInfos - 1

        hRes = callPointer(TTypeLib.GetTypeInfo, ptrTypeLib, i, VarPtr(ptrTypeInfo))

        If hRes <> S_OK Then
        
        Else
        
   
           ' ITypeInfo's VTable
           CpyMem pVTbl, ByVal ptrTypeInfo, 4
           CpyMem TTypeInfo, ByVal pVTbl, Len(TTypeInfo)
   
           ' TYPEATTR struct, which describes the type
           callPointer TTypeInfo.GetTypeAttr, ptrTypeInfo, VarPtr(pAttr)
           CpyMem TTypeAttr, ByVal pAttr, Len(TTypeAttr)
           callPointer TTypeInfo.ReleaseTypeAttr, ptrTypeInfo, pAttr
   
           ' name of the type
           callPointer TTypeLib.GetDocumentation, ptrTypeLib, i, VarPtr(strTypeName), 0, 0, 0
   
           If TTypeAttr.TYPEKIND = TKIND_COCLASS Then
               ReDim Preserve udtCoClasses(lngCoCls) As CoClass
   
               With udtCoClasses(lngCoCls)
                   .GUID = TTypeAttr.GUID
                   .Name = strTypeName
                   .prgid = clsidToProgID(guid2Str(.GUID))
               End With
   
               lngCoCls = lngCoCls + 1
           End If
   
           ' Set ITypeInfo = Nothing
           callPointer TTypeInfo.IUnk.Release, ptrTypeInfo
           ptrTypeInfo = 0
        End If
    Next

    ' Set ITypeLib = Nothing
    callPointer TTypeLib.IUnk.Release, ptrTypeLib
    '
    getCoClasses = True

End Function

Private Function callPointer(ByVal fnc As Long, ParamArray params()) As Long

    Dim btASM(&HEC00& - 1) As Byte
    Dim pASM As Long
    Dim i As Integer

    pASM = VarPtr(btASM(0))

    addByte pASM, &H58                  ' POP EAX
    addByte pASM, &H59                  ' POP ECX
    addByte pASM, &H59                  ' POP ECX
    addByte pASM, &H59                  ' POP ECX
    addByte pASM, &H59                  ' POP ECX
    addByte pASM, &H50                  ' PUSH EAX

    For i = UBound(params) To 0 Step -1
        addPush pASM, CLng(params(i))   ' PUSH dword
    Next

    addCall pASM, fnc                   ' CALL rel addr
    addByte pASM, &HC3                  ' RET

    callPointer = CallWindowProc(VarPtr(btASM(0)), 0, 0, 0, 0)

End Function

Private Sub addPush(pASM As Long, lng As Long)

    addByte pASM, &H68
    addLong pASM, lng


End Sub

Private Sub addCall(pASM As Long, addr As Long)

    addByte pASM, &HE8
    addLong pASM, addr - pASM - 4

End Sub

Private Sub addLong(pASM As Long, lng As Long)

    CpyMem ByVal pASM, lng, 4
    pASM = pASM + 4

End Sub

Private Sub addByte(pASM As Long, bt As Byte)

    CpyMem ByVal pASM, bt, 1
    pASM = pASM + 1

End Sub

Private Function guid2Str(GUIDBytes() As Byte) As String

    Dim nTemp As String
    Dim nGUID(15) As Byte
    Dim nLength As Long

    nTemp = Space$(78)
    CpyMem nGUID(0), GUIDBytes(0), 16
    nLength = StringFromGUID2(nGUID(0), nTemp, Len(nTemp))
    guid2Str = Left$(StrConv(nTemp, vbFromUnicode), nLength - 1)

End Function

Private Function clsidToProgID(ByVal CLSID As String) As String

    Dim pResult As Long, pChar As Long
    Dim char As Integer, Length As Long
    Dim GUID(15) As Byte

    CLSIDFromString StrPtr(CLSID), GUID(0)
    ProgIDFromCLSID GUID(0), pResult
    If pResult = 0 Then Exit Function

    pChar = pResult - 2
    Do
        pChar = pChar + 2
        CpyMem char, ByVal pChar, 2
    Loop While char

    Length = pChar - pResult
    clsidToProgID = Space$(Length \ 2)
    CpyMem ByVal StrPtr(clsidToProgID), ByVal pResult, Length

End Function

'/** @} **/ '<-- Ende der Doxygen-Gruppen-Zuordnung

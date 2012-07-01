Version =20
VersionRequired =20
Checksum =-1108149879
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    RecordLocks =2
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4478
    DatasheetFontHeight =11
    ItemSuffix =1
    Right =9660
    Bottom =7815
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xb154a1450e02e440
    End
    GUID = Begin
        0xc6dc3a8be3108f47b0476cfd4b6c84cc
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    PrtDevMode = Begin
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x010402059c00c40253ef800101007a009a0b3408640001000f00580202000100 ,
        0x5802030001004134000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000010000000000000001000000 ,
        0x0200000001000000000000000000000000000000000000000000000050524956 ,
        0xe230000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000180000000000102710271027 ,
        0x0000102700000000000000008800c40200000000000000000100000000000000 ,
        0x0000000000000000030000000000000000001000503403002888040000000000 ,
        0x000000000000010000000000000000000000000000000000e7b14b4c03000000 ,
        0x05000a00ff000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000
    End
    PrtDevNames = Begin
        0x080013001e000100000000000000000000000000000000000000000000007064 ,
        0x66636d6f6e00
    End
    OnTimer ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Section
            Height =1644
            Name ="Detailbereich"
            GUID = Begin
                0x118022fb34ada84781d54f5fb884cc84
            End
            AlternateBackColor =15921906
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =795
                    Top =675
                    Width =2955
                    Height =285
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bezeichnungsfeld0"
                    Caption ="Anwendung wird initialisiert ..."
                    GUID = Begin
                        0xe5bed945b0777042bffd174bf1d84450
                    End
                    GridlineColor =10921638
                    LayoutCachedLeft =795
                    LayoutCachedTop =675
                    LayoutCachedWidth =3750
                    LayoutCachedHeight =960
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: frmLibRepair
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Startformular zur initialisierung der ACLib .NET Integration
' </summary>
' <remarks></remarks>
'\ingroup DotNetLib
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_dotnetlib/integration/DotNetLibRepair.frm</file>
'  <use>_dotnetlib/integration/DotNetLibs.cls</use>
'  <use>_dotnetlib/integration/DotNetLibsSetup.bas</use>
'  <use>file/LibFiles.cls</use>
'  <license>_codelib/license.bas</license>
'  <description>Startformular zur initialisierung der ACLib .NET Integration</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Sub Form_Open(Cancel As Integer)
   Me.TimerInterval = 10
End Sub

Private Sub Form_Timer()
   Me.TimerInterval = 0
   LibFiles.ReInitialize
   DoCmd.Close acForm, Me.Name
End Sub
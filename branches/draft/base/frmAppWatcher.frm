Version =19
VersionRequired =19
Checksum =992900355
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5102
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =4755
    Top =3510
    Right =18645
    Bottom =8940
    DatasheetGridlinesColor =12632256
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0xf6285ccfde3ee340
    End
    Caption ="Anwendungsstart"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    PrtMip = Begin
        0xae050000ae050000ae050000ae05000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    PrtDevMode = Begin
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x010400069c00ec0253ef8001010009009a0b3408640001000f00580202000100 ,
        0x5802030001004134000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000010000000000000001000000 ,
        0x0200000001000000ffffffff0000000000000000000000000000000044494e55 ,
        0x2200b000ec020000c1951cfb0000000000000000000000000000000000000000 ,
        0x0000000000000000060000000100000000000000020001000000000000000000 ,
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
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x000000000000000000000000010000000000000000000000b0000000534d544a ,
        0x000000001000a00053006e006100670069007400200039002000500072006900 ,
        0x6e007400650072000000496e70757442696e004d414e55414c00524553444c4c ,
        0x00556e69726573444c4c004f7269656e746174696f6e00504f52545241495400 ,
        0x506170657253697a65004c4554544552005265736f6c7574696f6e004f707469 ,
        0x6f6e3300436f6c6f724d6f646500323462707000000000000000000000000000 ,
        0x0000000000000000
    End
    PrtDevNames = Begin
        0x0800190022000100000000000000000000000000000000000000000000000000 ,
        0x0000483a5c50726f6772616d446174615c54656368536d6974685c536e616769 ,
        0x7420395c5072696e746572506f727446696c6500
    End
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =850
            BackColor =5263440
            Name ="Detailbereich"
            GUID = Begin
                0x15795c3f729178469bb66e378f9acc20
            End
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =50
                    Top =226
                    Width =4995
                    Height =420
                    FontSize =16
                    BackColor =5263440
                    ForeColor =12632256
                    Name ="labInfo"
                    Caption ="Anwendung wird initialisiert ..."
                    GUID = Begin
                        0xafd80fe146c5184fa8e27ee33b92a82e
                    End
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
' Form: frmAppWatcher (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsformular für Anwendungsstart (inkl. DisposeCurrentApplicationHandler beim Beenden)
' </summary>
' <remarks>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/frmAppWatcher.frm</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_AppLoaded As Boolean

Private Sub Form_Current()
On Error Resume Next
   Me.TimerInterval = 100
   ' => damit Formular gezeichnet wird, falls der Anwendungsstart länger dauert
End Sub

Private Sub Form_Load()
   Dim lngBackColor As Long
On Error Resume Next
   Me.Caption = CurrentApplication.ApplicationTitle
   lngBackColor = CurrentApplication.MdiBackColor
   If lngBackColor <> 0 Then
      setFormColor lngBackColor
   End If
End Sub

Private Sub Form_Timer()
On Error Resume Next
   Me.TimerInterval = 0
   If Not m_AppLoaded Then
      m_AppLoaded = CurrentApplication.Start
   End If
   Me.Visible = False
   If Not m_AppLoaded Then DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   DisposeCurrentApplicationHandler
End Sub

Private Sub setFormColor(ByVal lBackColor As Long)
On Error Resume Next
   Me.Section(0).BackColor = lBackColor
   Me.labInfo.ForeColor = &HFFFFFF Xor (lBackColor / 2)
End Sub
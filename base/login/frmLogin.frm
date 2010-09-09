Version =19
VersionRequired =19
Checksum =-122265359
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =124
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =3911
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =3255
    Top =2160
    Right =19395
    Bottom =14235
    DatasheetGridlinesColor =12632256
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x68568eb19ee2e240
    End
    Caption ="Login"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xae050000ae050000ae050000ae05000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    PrtDevMode = Begin
        0x00000000ba28d92f0400000001000000c04c130018000000344f1300e84c1300 ,
        0x010400069c00440353ef8001010009009a0b3408640001000f00580202000100 ,
        0x580203000100413400000000c8fc8a09204f1300c4fc8a09184d1300e993db2f ,
        0xe0fc8a09f0930000000000000000000000000000010000000000000001000000 ,
        0x0200000001000000000000000000000000000000000000000000000050524956 ,
        0xe230000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000180000000000102710271027 ,
        0x0000102700000000000000007000440300000000000000000000000000000000 ,
        0x0000000000000000030000000000000000001000503403002888040000000000 ,
        0x0000000000000100000000000000000000000000000000002401df8c03000000 ,
        0x05000400ff000000000000000000000000000000000000000000000000000000 ,
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
        0x0100000000000000000000000000000070000000534d544a0000000010006000 ,
        0x46007200650065005000440046005f005800500000005265736f6c7574696f6e ,
        0x00363030647069005061676553697a65004c6574746572005061676552656769 ,
        0x6f6e000000000000000000000000000000000000000000000000000000000000
    End
    PrtDevNames = Begin
        0x080013001b000100000000000000000000000000000000000000004672656550 ,
        0x44465850313a00
    End
    OnLoad ="[Event Procedure]"
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =1701
            BackColor =-2147483633
            Name ="Detailbereich"
            GUID = Begin
                0x1c21f281a7059d40ae5d54243054118e
            End
            Begin
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1583
                    Top =225
                    Width =2154
                    Height =255
                    FontSize =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtUID"
                    GUID = Begin
                        0x9a0c401e28084442b6474a35fe0eb00d
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =165
                            Top =225
                            Width =1293
                            Height =255
                            FontSize =9
                            Name ="Bezeichnungsfeld1"
                            Caption ="Benutzername"
                            GUID = Begin
                                0x0e7d67976e5b104499341fcd6cb86ce6
                            End
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1583
                    Top =621
                    Width =2154
                    Height =255
                    FontSize =9
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtPwd"
                    InputMask ="Password"
                    GUID = Begin
                        0x38f22128a6703e4fad4101d5897d146c
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =165
                            Top =621
                            Width =1293
                            Height =255
                            FontSize =9
                            Name ="Bezeichnungsfeld3"
                            Caption ="Kennwort"
                            GUID = Begin
                                0x0ad270581b4cf74e8a996e77d30b83be
                            End
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =79
                    Left =170
                    Top =1133
                    Width =1418
                    Height =418
                    TabIndex =2
                    Name ="cmdLogin"
                    Caption ="&OK"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0xc962fbc74b72ab478f25843d24a1c4d1
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =65
                    Left =2324
                    Top =1133
                    Width =1418
                    Height =418
                    TabIndex =3
                    Name ="cmdCancel"
                    Caption ="&Abbrechen"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x0545355ba5721040a7f36862e5274afe
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
' Form: frmLogin (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsformular für Anwendungslogin
' </summary>
' <remarks>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/login/frmLogin.frm</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'  <use>base/modApplication.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Event Logon(ByVal LoginName As String, ByVal Password As String, ByRef Cancel As Boolean)
Public Event Cancelled()

Public m_CloseFormOK As Boolean

Private Sub cmdCancel_Click()
On Error Resume Next
   RaiseEvent Cancelled
   m_CloseFormOK = True
   DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdLogin_Click()

   Dim strLoginName As String
   Dim strLoginPassword As String
   Dim bolCancel As Boolean

#If USE_GLOBAL_ERRORHANDLER Then
On Error GoTo HandleErr
#End If

   strLoginName = Me.txtUID & vbNullString
   strLoginPassword = Me.txtPwd & vbNullString
   
   If Len(strLoginName) * Len(strLoginPassword) = 0 Then
      MsgBox "Bitte Loginnamen und Kennwort angeben"
      If Len(strLoginName) = 0 Then
         Me.txtUID.SetFocus
      Else
         Me.txtPwd.SetFocus
      End If
      Exit Sub
   End If
   
   RaiseEvent Logon(strLoginName, strLoginPassword, bolCancel)
   If bolCancel Then
      Me.txtPwd.SetFocus
      Exit Sub
   End If
      
   m_CloseFormOK = True
   DoCmd.Close acForm, Me.Name

#If USE_GLOBAL_ERRORHANDLER Then
ExitHere:
On Error Resume Next
   Exit Sub
HandleErr:
   Select Case HandleError(Err.Number, "cmdLogin_Click", Err.Description)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
#End If

End Sub

Private Sub Form_Load()
   
   Dim lngPos As Long
  
#If USE_GLOBAL_ERRORHANDLER Then
On Error GoTo HandleErr
#End If

   If Len(Me.OpenArgs) > 0 Then 'OpenArgs= "Caption@StandardUser"
      lngPos = InStr(1, Me.OpenArgs, "@")
      Me.Caption = Left$(Me.OpenArgs, lngPos - 1)
      Me.txtUID = Mid$(Me.OpenArgs, lngPos + 1)
      If Len(Me.txtUID.Value) > 0 Then
         Me.txtPwd.SetFocus
      End If
   End If

'   If Not (TempLoginFormLoginHandlerRef Is Nothing) Then
'      Set TempLoginFormLoginHandlerRef.LoginForm = Me 'damit werden die Events trotz Dialog-Modus möglich
'   End If
   Dim tempObj As Object
   Set tempObj = CurrentApplication.Extensions("AppLogin")
   If Not (tempObj Is Nothing) Then
      Set tempObj.LoginForm = Me 'damit werden die Events trotz Dialog-Modus möglich
   End If
   
#If USE_GLOBAL_ERRORHANDLER Then
ExitHere:
On Error Resume Next
   Exit Sub
HandleErr:
   Select Case HandleError(Err.Number, "Form_Load", Err.Description)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
#End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   If Not m_CloseFormOK Then RaiseEvent Cancelled
End Sub
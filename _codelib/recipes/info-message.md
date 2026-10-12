# Non-modal info message with auto close

Keywords: info message, toast, notification, non-modal message, auto close, fade out, transparency, status message,
Infomeldung, Hinweis einblenden, nicht modal, automatisch ausblenden, Benachrichtigung

## Import

- `form/InfoMessageBox.cls` (uses the helper form `form/InfoMessageBoxForm.frm`)

## Usage

```vba
' Show for 5 seconds with 25 % transparency
With New InfoMessageBox
   .Show "Data saved.", 5000, 25
End With

' Position, colors, fade-in and fade-out (times in ms); simple HTML formatting in the text
With New InfoMessageBox
   .Move 250, 250, 2600                       ' left, top, width (twips)
   .BackColor = vbBlue
   .ForeColor = vbWhite
   .Show "<b>Import</b> finished", 500, 25, 2000, 2500
End With

' Visible until the user clicks on it
With New InfoMessageBox
   .Show "Please check the log.", 0, 50
End With
```

## Notes

- `Show(Text, CloseTimer, TransparencyFactor, FadeOutTime, FadeInTime)`: `CloseTimer = 0` keeps the message open
  until it is clicked.
- The instance keeps a reference to itself while visible, so `With New InfoMessageBox` is sufficient.
- Another form can be assigned with `Set .Form = New Form_MyInfoForm` (e.g. `form/InfoMessageBoxForm_mit_Buttons.frm`).
- `form/frmInfoMessage.frm` is an older standalone form version; prefer the class.

## See also

- `_examples/form/InfoMessageBox_Examples.frm`

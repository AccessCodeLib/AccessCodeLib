# Sort a continuous form by clicking on the column header

Keywords: sort, order by, column header, continuous form, list form, sort symbol, ascending, descending,
sortieren, Spaltenkopf, Endlosformular, Sortiersymbol, aufsteigend, absteigend

## Import

- `form/sort/FormSortingExtension.cls`

## Usage

Code in the continuous form (header labels or buttons as sorting controls):

```vba
Private WithEvents m_Sorting As FormSortingExtension

Private Sub Form_Load()
   Set m_Sorting = New FormSortingExtension
   With m_Sorting
      .DefaultSortSymbolString(False) = " ^"          ' optional: symbol appended to the caption (ascending)
      .DefaultSortSymbolString(True) = " v"           ' ... (descending)
      .Add Me.lblLastName, "LastName, FirstName"
      .Add Me.lblOrderDate, "OrderDate DESC, OrderID"  ' fixed DESC/ASC is not toggled
      .Add Me.cmdAmount, "Amount"
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Sorting = Nothing
End Sub

Private Sub m_Sorting_OrderByChanged(ByVal OrderByString As String)
   Me.OrderBy = OrderByString
   Me.OrderByOn = True
End Sub
```

## Notes

- Supported sorting controls: labels and command buttons. The click event is connected automatically.
- A second click on the same control toggles ascending/descending for all fields without fixed `ASC`/`DESC`.
- `InitSortSymbolLabel` positions one separate label as sort symbol next to the active control.
- `form/sort/FormSortingExtension.cls` declares its WinAPI functions without `PtrSafe`: it does not compile in 64-bit Office
  without adjustment.
- `form/sort/ColumnSort.cls` is a simpler alternative with command buttons
  (`_examples/form/sort/frmColumnSort.frm`).

## See also

- [filter-form.md](filter-form.md) - filter the same form

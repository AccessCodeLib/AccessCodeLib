# Search form with filter controls

Keywords: search form, filter form, unbound filter controls, form header filter, auto filter, continuous form, subform filter,
Suchformular, Filterformular, ungebundene Steuerelemente, Formularkopf, Endlosformular, Unterformular filtern

## Import

- `form/filter/FilterControlManager.cls` (dependencies are added by the import: FilterControlCollection, FilterControl,
  FilterControlEventBridge, FilterStringBuilder, SqlTools, StringCollection)

The ACLib FilterFormWizard add-in generates the code below for a form.

## Usage

Code in the form module (filter controls e.g. in the form header, data in the form itself or in a subform):

```vba
Private WithEvents m_FilterControlManager As FilterControlManager

Private Sub Form_Load()
   InitFilterControlManager
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_FilterControlManager = Nothing
End Sub

Private Sub InitFilterControlManager()
   Set m_FilterControlManager = New FilterControlManager
   m_FilterControlManager.ConfigSqlFormat "\#yyyy-mm-dd hh:nn:ss\#", "True", "*"   ' DAO/Jet/ACE

   With m_FilterControlManager.FilterControls
      .Add "LastName", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Me.txtLastName
      .Add "StatusID", SQL_Numeric, SQL_Equal, Me.cboStatus
      .Add "OrderDate", SQL_Date, SQL_Between + SQL_Add_WildCardSuffix, Me.txtDateFrom, Me.txtDateTo
   End With

   m_FilterControlManager.AutoFilterOn = True   ' filter on every change; False: only with ApplyFilter
End Sub

Private Sub m_FilterControlManager_FilterStringChanged(ByVal NewFilterString As String)
   With Me            ' or: With Me.sfrList.Form
      .Filter = NewFilterString
      .FilterOn = (Len(.Filter) > 0)
   End With
End Sub

Private Sub cmdApplyFilter_Click()
   m_FilterControlManager.ApplyFilter
End Sub

Private Sub cmdRemoveFilter_Click()
   m_FilterControlManager.RemoveFilter          ' clears the filter controls, shows all records
   ' m_FilterControlManager.RemoveFilter "1=0"  ' ... or show no records
End Sub
```

## Notes

- Empty filter controls create no condition.
- `.Add` takes an optional second control (e.g. date from/to for `SQL_Between`) and an `IgnoreValue`.
- `.AddGroup(SQL_Or)` returns a nested collection for Or groups; `.AddMultiControlCriteria`, `.AddSubSelectCriteria` and
  `.AddExistsCriteria` cover more complex filters.
- The event `FilterValuesChanged` can be used to signal that the filter must be applied (e.g. bold apply button).
- Variants with the same object names exist for the ACLib DotNetLib (`*_DotNetLib.cls`); import only one variant.

## See also

- [filter-string.md](filter-string.md) - filter string without form controls
- [sql-criteria.md](sql-criteria.md) - relational operators and data types
- tests `_test/form/filter/FilterControlCollectionTests.cls` with `_test/form/filter/FilterControlTestForm.frm`

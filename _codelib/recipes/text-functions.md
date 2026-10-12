# Text functions (format, pad, trim, placeholders)

Keywords: string functions, format text, placeholder {0}, pad left, pad right, trim, contains, starts with, ends with,
hours over 24, string collection, join, Textfunktionen, Platzhalter, auffüllen, Stunden über 24, Zeichenkette verbinden

## Import

- `text/StringTools.cls` (preferred)
- `text/StringCollection.cls` (collect strings and join them)

## Usage

```vba
' StringTools.cls is a predeclared class: call with prefix, no New
Debug.Print StringTools.FormatText("Order {0} for {1}", 4711, "Miller")   ' Order 4711 for Miller
Debug.Print StringTools.Format(#12/31/1899 1:30:45 AM#, "[h]:nn:ss")      ' 25:30:45 (hours over 24)
Debug.Print StringTools.PadLeft("Test", 6, "_")                           ' __Test
Debug.Print StringTools.IsNullOrEmpty(Me.txtName.Value, True)            ' True for Null, "" and "   "
Debug.Print StringTools.Contains("abcdef", "cd"), StringTools.StartsWith("abc", "a")
Debug.Print StringTools.Trim("  abc  ", TrimStart)
Debug.Print StringTools.Replicate("-", 20)

' Collect strings and join them
With New StringCollection
   .Add "a"
   .Add "b"
   .AddFromArray Array(1, 2), "0.00"         ' optional format per item
   Debug.Print .ToString(", ")
End With
```

## Notes

- Two variants with the same object name exist; import only one:
  - `text/StringTools.cls` (predeclared class): calls with prefix `StringTools.`, no name conflicts with VBA functions.
  - `text/StringTools.bas` (standard module): its public `Format`, `Trim` etc. replace the VBA functions in the whole
    project (calls without prefix). Use it only if this is intended.
- `text/String2.cls` offers the same functions as chainable string type (`String2.NewValue("abc").PadLeft(6).ToString`).

## See also

- tests `_test/text/StringToolsTests.cls`, `_test/text/StringCollectionTests.cls`
- `_examples/text/StringCollection_Examples.bas`, `_examples/text/String2_Example.bas`

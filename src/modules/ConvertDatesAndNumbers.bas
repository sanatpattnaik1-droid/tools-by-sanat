Attribute VB_Name = "ConvertDatesAndNumbers"
Option Explicit

'==============================================================
' Main entry point - run this macro
'==============================================================
Sub ConvertSelectedCells()
    Dim cell As Range
    Dim parts As Variant

    If Selection Is Nothing Then
        MsgBox "Please select cells first.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            If VarType(cell.Value) = vbDate Then
                ' NEW: cell is ALREADY a real Excel date/datetime (e.g. auto-parsed
                ' by Excel on entry/import, which is why filters already showed the
                ' Year > Month > Day tree). Strip any time portion and apply the
                ' standard DD-MM-YY display format.
                cell.Value = Int(cell.Value)
                cell.NumberFormat = "DD-MM-YY"
            Else
                parts = TryExtractDateParts(CStr(cell.Value))
                If IsArray(parts) And UBound(parts) = 2 Then
                    ' Write as a real Excel date so filters show Year > Month > Day hierarchy
                    cell.Value = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
                    cell.NumberFormat = "DD-MM-YY"
                ElseIf IsLargeNumber(cell) Then
                    ConvertToText cell
                ElseIf IsNumericTextValue(cell) Then
                    ' NEW: fixes columns like "Total Amount" where numbers are
                    ' stored as text (left-aligned) and therefore don't SUM.
                    ConvertToNumber cell
                End If
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "Conversion complete!", vbInformation
End Sub

'==============================================================
' Parses all known date formats.
' Returns a 3-element array: (day, month, year) as strings.
' Returns empty array if not recognised.
'==============================================================
Function TryExtractDateParts(raw As String) As Variant
    Dim s As String
    Dim yr As String, mo As String, dy As String

    s = Trim(raw)

    ' --- YYYY?MM?DD ...  (covers YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD)
    '     e.g. 2026-05-18 06:20:44+00:00 | 2026-05-28T18:29:59Z | 2026.05.29 09:10:07
    If Len(s) >= 10 Then
        Dim sep1 As String, sep2 As String
        sep1 = Mid(s, 5, 1)
        sep2 = Mid(s, 8, 1)

        If (sep1 = "-" Or sep1 = "." Or sep1 = "/") And _
           (sep2 = "-" Or sep2 = "." Or sep2 = "/") And _
           IsNumeric(Left(s, 4)) And Val(Left(s, 4)) > 1900 Then

            yr = Left(s, 4)
            mo = Mid(s, 6, 2)
            dy = Mid(s, 9, 2)

            If IsValidDateParts(dy, mo, yr) Then
                TryExtractDateParts = Array(dy, mo, yr)
                Exit Function
            End If
        End If
    End If

    ' --- DD-MM-YYYY / DD/MM/YYYY / DD.MM.YYYY ... (e.g. 13-05-2026 9:00, 01/07/2026)
    ' NEW: generalized to accept "-", "/", or "." as the separator so that
    ' text dates like "01/07/2026" also convert to real Excel dates
    ' (this makes Excel's AutoFilter show the Year > Month > Day tree
    '  instead of a flat text list).
    If Len(s) >= 10 Then
        Dim p1 As Integer, p2 As Integer
        Dim seps As Variant, si As Integer

        seps = Array("/", "-", ".")
        For si = LBound(seps) To UBound(seps)
            p1 = InStr(s, CStr(seps(si)))
            p2 = InStr(p1 + 1, s, CStr(seps(si)))

            If p1 > 0 And p2 > 0 Then
                Dim part1 As String, part2 As String, part3 As String
                part1 = Left(s, p1 - 1)
                part2 = Mid(s, p1 + 1, p2 - p1 - 1)
                part3 = Left(Mid(s, p2 + 1), 4)

                If Len(part1) <= 2 And Len(part3) = 4 And _
                   IsNumeric(part1) And IsNumeric(part2) And IsNumeric(part3) Then
                    dy = Format(Val(part1), "00")
                    mo = Format(Val(part2), "00")
                    yr = part3
                    If IsValidDateParts(dy, mo, yr) Then
                        TryExtractDateParts = Array(dy, mo, yr)
                        Exit Function
                    End If
                End If
            End If
        Next si
    End If

    TryExtractDateParts = Array()
End Function

'==============================================================
' Validates day/month/year strings
'==============================================================
Function IsValidDateParts(dy As String, mo As String, yr As String) As Boolean
    If Not IsNumeric(dy) Or Not IsNumeric(mo) Or Not IsNumeric(yr) Then
        IsValidDateParts = False: Exit Function
    End If
    Dim d As Integer, m As Integer, y As Integer
    d = CInt(dy): m = CInt(mo): y = CInt(yr)
    IsValidDateParts = (d >= 1 And d <= 31) And (m >= 1 And m <= 12) And (y >= 1900 And y <= 2100)
End Function

'==============================================================
' Returns True for numeric cells with >= 10 digits (large numbers)
' e.g. IDs, phone numbers, account numbers - these should be TEXT.
'==============================================================
Function IsLargeNumber(cell As Range) As Boolean
    If IsNumeric(cell.Value) Then
        IsLargeNumber = (Len(Format(Abs(CDbl(cell.Value)), "0")) >= 10)
    Else
        IsLargeNumber = False
    End If
End Function

'==============================================================
' NEW: Returns True for cells that LOOK like normal numbers/amounts
' (e.g. "3375.00", "6,600.00", " 7493.12 ") but are actually stored
' as TEXT, which is why SUM/AVERAGE on that column returns 0 or
' skips them. Large-number IDs are excluded (handled separately).
'==============================================================
Function IsNumericTextValue(cell As Range) As Boolean
    Dim v As String
    Dim cleaned As String

    ' Only cells whose underlying stored value is text qualify.
    ' (A cell already holding a true number/date doesn't need conversion.)
    If VarType(cell.Value) <> vbString Then
        IsNumericTextValue = False
        Exit Function
    End If

    v = Trim(cell.Value)
    If Len(v) = 0 Then
        IsNumericTextValue = False
        Exit Function
    End If

    ' Strip thousands separators (commas) so "6,600.00" still validates
    cleaned = Replace(v, ",", "")

    If Not IsNumeric(cleaned) Then
        IsNumericTextValue = False
        Exit Function
    End If

    ' Skip long numeric strings (IDs, account numbers, etc.) -
    ' those are intentionally kept/converted as text elsewhere.
    Dim digitsOnly As String
    digitsOnly = Replace(Replace(Replace(cleaned, "-", ""), ".", ""), "+", "")
    If Len(digitsOnly) >= 10 Then
        IsNumericTextValue = False
        Exit Function
    End If

    IsNumericTextValue = True
End Function

'==============================================================
' NEW: Converts a text-stored numeric value into a real Excel
' number, preserving 2 decimal places (adjust format as needed).
'==============================================================
Sub ConvertToNumber(cell As Range)
    Dim cleaned As String
    Dim numVal As Double

    cleaned = Replace(Trim(cell.Value), ",", "")
    numVal = CDbl(cleaned)

    cell.NumberFormat = "General"
    cell.Value = numVal
    cell.NumberFormat = "0.00"
End Sub

'==============================================================
' Converts a numeric cell to text, preserving all digits
'==============================================================
Sub ConvertToText(cell As Range)
    Dim numStr As String
    numStr = cell.Text
    If InStr(numStr, "E") > 0 Or InStr(numStr, "e") > 0 Then
        numStr = Format(CDec(cell.Value), "0")
    End If
    cell.NumberFormat = "@"
    cell.Value = numStr
End Sub



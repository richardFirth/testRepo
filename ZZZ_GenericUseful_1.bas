Attribute VB_Name = "ZZZ_GenericUseful_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.0
'$*DATE*19Jan18
'$*ID*GenericUseful



Option Explicit



Public Sub saveWorkbookToDesktop(theWKBK As Workbook, theName As String)
    theWKBK.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & theName)
End Sub


Public Sub addSheetWithName(theName As String, theBook As Workbook)
   theBook.Sheets.Add(After:=theBook.Worksheets(theBook.Worksheets.Count)).Name = theName
End Sub



Public Sub complexRoutineStart(notUsed As String)
    With Application
        .ShowWindowsInTaskbar = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Sub

Public Sub endComplex()
    Call complexRoutineEnd("")
End Sub


Public Sub complexRoutineEnd(notUsed As String)
    With Application
        .ShowWindowsInTaskbar = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Public Function GenerateIDcode() As String

Randomize

Dim a As Integer:   a = Int(9 * Rnd) + 1
Dim B As Integer:   B = Int(9 * Rnd) + 1
Dim C As Integer:   C = Int(9 * Rnd) + 1
Dim d As Integer:   d = Int(9 * Rnd) + 1

Dim y As Integer:   y = Int(26 * Rnd) + 1

Select Case y

    Case 1: GenerateIDcode = "A" & a & B & C & d
    Case 2: GenerateIDcode = "B" & a & B & C & d
    Case 3: GenerateIDcode = "C" & a & B & C & d
    Case 4: GenerateIDcode = "D" & a & B & C & d
    Case 5: GenerateIDcode = "E" & a & B & C & d
    Case 6: GenerateIDcode = "F" & a & B & C & d
    Case 7: GenerateIDcode = "G" & a & B & C & d
    Case 8: GenerateIDcode = "H" & a & B & C & d
    Case 9: GenerateIDcode = "I" & a & B & C & d
    Case 10: GenerateIDcode = "J" & a & B & C & d
    Case 11: GenerateIDcode = "K" & a & B & C & d
    Case 12: GenerateIDcode = "L" & a & B & C & d
    Case 13: GenerateIDcode = "M" & a & B & C & d
    Case 14: GenerateIDcode = "N" & a & B & C & d
    Case 15: GenerateIDcode = "O" & a & B & C & d
    Case 16: GenerateIDcode = "P" & a & B & C & d
    Case 17: GenerateIDcode = "Q" & a & B & C & d
    Case 18: GenerateIDcode = "R" & a & B & C & d
    Case 19: GenerateIDcode = "S" & a & B & C & d
    Case 20: GenerateIDcode = "T" & a & B & C & d
    Case 21: GenerateIDcode = "U" & a & B & C & d
    Case 22: GenerateIDcode = "V" & a & B & C & d
    Case 23: GenerateIDcode = "W" & a & B & C & d
    Case 24: GenerateIDcode = "X" & a & B & C & d
    Case 25: GenerateIDcode = "Y" & a & B & C & d
    Case 26: GenerateIDcode = "Z" & a & B & C & d

End Select

End Function


'5/19/2015 - Version 70 - added date tracking
'8/14/2015 - Version 75 - added option to convert any date to an abbott date

Public Function abbottDate(Optional aDate As Date = 0) As String

If aDate = 0 Then aDate = Now()

Dim currentDay As Integer
Dim currentMonthNumber As Integer
Dim currentYear As Integer
Dim currentMonth As String


currentDay = Day(aDate)
currentMonthNumber = month(aDate)
currentYear = Year(aDate)
currentMonth = ""


Select Case currentMonthNumber
    Case 1
       currentMonth = "Jan"
    Case 2
       currentMonth = "Feb"
    Case 3
       currentMonth = "Mar"
    Case 4
       currentMonth = "Apr"
    Case 5
       currentMonth = "May"
    Case 6
       currentMonth = "Jun"
    Case 7
       currentMonth = "Jul"
    Case 8
       currentMonth = "Aug"
    Case 9
       currentMonth = "Sep"
    Case 10
       currentMonth = "Oct"
    Case 11
       currentMonth = "Nov"
    Case 12
       currentMonth = "Dec"

End Select

abbottDate = currentDay & currentMonth & currentYear

End Function











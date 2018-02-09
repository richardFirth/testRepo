Attribute VB_Name = "ZZZ_CSVAndTextInteraction_1"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.5
'$*DATE*29Jan18
'$*ID*CSVAndTextInteraction

Option Explicit

    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0


Function appendtoCSV(filePath As String, contents() As String) As Boolean
    
    Dim x As Integer
    On Error GoTo BADappendtoCSV
    Open filePath For Append As #1
        For x = LBound(contents) To UBound(contents)
            If x = UBound(contents) Then
                Write #1, contents(x)
            Else
                Write #1, contents(x),
            End If
        Next x
    Close #1
    appendtoCSV = True
    
Exit Function
BADappendtoCSV:
    appendtoCSV = False
End Function




Function getCSVFromFile(filePath As String) As String()

Dim csvRow() As String
Dim dataIN As String
Dim x As Integer
      
Open filePath For Input As #4
    
    For x = 1 To 4500
        On Error GoTo BADgetCSVFromFile
        'Input #4, dataIN
        Line Input #4, dataIN
        ReDim Preserve csvRow(1 To x) As String
        csvRow(x) = dataIN
    Next x
    
finishedA:
    
    Close #4
    
getCSVFromFile = csvRow
    
Exit Function
BADgetCSVFromFile:

    Resume finishedA

End Function



Function getTxTDocumentAsString(thePath As String) As String()
    Dim theFileContents As String
    
    On Error GoTo getTXTerr
    theFileContents = CreateObject("Scripting.FileSystemObject").GetFile(thePath).OpenAsTextStream(ForReading, TristateUseDefault).ReadAll
    
    Dim docFeed() As String
    docFeed = Split(theFileContents, Chr(10))
    getTxTDocumentAsString = TrimAndCleanArray(docFeed)
Exit Function
getTXTerr:

End Function



Sub createTextFromStringArr(theContents() As String, theFullPathName As String)

    Dim fs As Object, f As Object
    Set fs = CreateObject("Scripting.FileSystemObject") ' this creates the fileSystemObject object for all file operations
    
    Call createFile(f, fs, theFullPathName)

    Dim n As Integer
    If arrayHasStuff(theContents) Then
        For n = LBound(theContents) To UBound(theContents)
            f.WriteLine theContents(n)
        Next n
    End If

End Sub


Sub createTextFromString(theContents As String, theFullPathName As String)

    Dim fs As Object, f As Object
    Set fs = CreateObject("Scripting.FileSystemObject") ' this creates the fileSystemObject object for all file operations
    
    Call createFile(f, fs, theFullPathName)

    f.WriteLine theContents

End Sub



Function createFile(fileobject As Object, FSO As Object, fileName As String) As Boolean
On Error GoTo createFileError

    Set fileobject = FSO.CreateTextFile(fileName, True)
    
Exit Function
createFileError:
    MsgBox "Problem: " & fileName
End Function







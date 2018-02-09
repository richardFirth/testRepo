VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AA_FolderItemHistory_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'$VERSIONCONTROL
'$*MINOR_VERSION*1.5
'$*DATE*29Jan18
'$*ID*folderItemsCLS



Option Explicit


Private mainFolderPath As String
Private dataListFileName As String
Private setupHappened As Boolean

' SETUP Valid filetypes



Public Sub initializeData(tPath As String, tDatList As String)
    mainFolderPath = tPath
    dataListFileName = tDatList
    setupHappened = True
    
    If Not FileThere(mainFolderPath & "\" & dataListFileName) Then
        ' initial setup of data?
    End If
    
End Sub


Public Sub commitPresentData()
    Call createTextFromStringArr(ZgetPresentData, mainFolderPath & "\" & dataListFileName)
End Sub


Public Function ZgetPresentData() As String()
    Dim locSTR() As String: locSTR = getFilePathsInFolder2Array(mainFolderPath)
    Dim locSTRB() As String
    
    Dim x As Integer
    Dim n As Integer: n = 1
    For x = LBound(locSTR) To UBound(locSTR)
        If Right(locSTR(x), 4) = "xlsx" Then
            ReDim Preserve locSTRB(1 To n) As String
            locSTRB(n) = locSTR(x)
            n = n + 1
        End If
    
    Next x
    
    ZgetPresentData = locSTRB

End Function

Public Function ZgetDataList() As String()

If FileThere(mainFolderPath & "\" & dataListFileName) Then
    ZgetDataList = getTxTDocumentAsString(mainFolderPath & "\" & dataListFileName)
End If

End Function

Public Function getFilesAdded() As String()
   getFilesAdded = newStrings(ZgetDataList, ZgetPresentData)
End Function

Public Function getFilesRemoved() As String()
   getFilesRemoved = newStrings(ZgetPresentData, ZgetDataList)
End Function

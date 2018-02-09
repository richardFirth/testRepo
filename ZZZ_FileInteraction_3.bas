Attribute VB_Name = "ZZZ_FileInteraction_3"
'$VERSIONCONTROL
'$*MINOR_VERSION*1.1
'$*DATE*29Jan18
'$*ID*FileInteraction



Option Explicit

'/---ZZZ_FileInteraction_3---------------updated 8Jan18-------------------------------------\
'  Function Name       | Return          |   Description                                    |
'----------------------|-----------------|--------------------------------------------------|
' CopyFileRF           | Boolean         | copies a file                                    |
' MoveFileRF           | Boolean         | moves a file                                     |
' RenameFileRF         | Boolean         | renames a file                                   |
' createFolderOnDesktop| Boolean         | creates a directory on desktop                   |
' createDirectoryRF    | Boolean         | creates a directory                              |
' FolderThere          | Boolean         | checks if a folder is present                    |
' FileThere            | Boolean         | checks if a file is present                      |
'converts a variant array to a string array



'\------------------------------------------------------------------------------------------/


Public Enum getFileType
    A_CSV
    B_EXCEL
    C_EXCEL_OLD
    D_EXCEL_MACRO
    E_InoFile
    F_Proc3File
    G_VBAModule
    H_Text
    I_Lib
End Enum





 ' /------------------\
 ' |copies a file     |
 ' \------------------/
Public Function CopyFileRF(source As String, destination As String) As Boolean

On Error GoTo CopyFileRF_Fail
    FileCopy source, destination
    CopyFileRF = True
Exit Function

CopyFileRF_Fail:
    CopyFileRF = False
End Function


 ' /------------------\
 ' |moves a file      |
 ' \------------------/
Public Function MoveFileRF(source As String, destination As String) As Boolean

On Error GoTo MoveFileRF_Fail

If CopyFileRF(source, destination) Then
    Kill source
    MoveFileRF = True
End If

Exit Function
MoveFileRF_Fail:
    MoveFileRF = False
End Function

Function DeleteFileRF(thePath As String) As Boolean
'You can use this to delete all the files in the folder Test
    On Error GoTo DeleteFileRF_Fail
    Kill thePath
    DeleteFileRF = True
Exit Function
DeleteFileRF_Fail:
    DeleteFileRF = False
End Function

Function DeleteFolderRF(thePath As String) As Boolean
'You can use this to delete the whole folder
'Note: RmDir delete only a empty folder
    On Error GoTo DeleteFolderRF_Fail
    Kill thePath & "\*.*"    ' delete all files in the folder
    RmDir thePath & "\"  ' delete folder
    DeleteFolderRF = True
Exit Function
DeleteFolderRF_Fail:
    DeleteFolderRF = False
End Function





Sub Clear_All_Files_And_SubFolders_In_Folder()
'Delete all files and subfolders
'Be sure that no file is open in the folder
    Dim FSO As Object
    Dim MyPath As String

    Set FSO = CreateObject("scripting.filesystemobject")

    MyPath = "C:\Users\Ron\Test"  '<< Change
Exit Sub ' avoid bad problem

    If Right(MyPath, 1) = "\" Then
        MyPath = Left(MyPath, Len(MyPath) - 1)
    End If

    If FSO.FolderExists(MyPath) = False Then
        MsgBox MyPath & " doesn't exist"
        Exit Sub
    End If

    On Error Resume Next
    'Delete files
    FSO.deletefile MyPath & "\*.*", True
    'Delete subfolders
    FSO.deletefolder MyPath & "\*.*", True
    On Error GoTo 0

End Sub













 ' /-----------------------\
 ' |renames a file         |
 ' \-----------------------/
Public Function RenameFileRF(source As String, destination As String) As Boolean
    RenameFileRF = MoveFileRF(source, destination)
End Function



 ' /----------------------------------\
 ' |creates a directory on desktop    |
 ' \----------------------------------/
Public Function createFolderOnDesktop(ByVal dirName As String) As Boolean
    createFolderOnDesktop = createDirectoryRF(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & dirName)
End Function



 ' /----------------------\
 ' |creates a directory   |
 ' \----------------------/
Public Function createDirectoryRF(ByVal dirName As String) As Boolean

    On Error GoTo noFolder
        MkDir dirName
    On Error GoTo 0
    createDirectoryRF = True
    
    Exit Function
    
noFolder:
    createDirectoryRF = False
End Function






 ' /--------------------------------\
 ' |checks if a folder is present   |
 ' \--------------------------------/
Public Function FolderThere(folderPathToTest As String) As Boolean

    If folderPathToTest = "" Then FolderThere = False: Exit Function

    If Len(Dir(folderPathToTest, vbDirectory)) = 0 Then
        FolderThere = False
    Else
        FolderThere = True
    End If
    
    
End Function



' /-------------------------------\
' |checks if a file is present    |
' \-------------------------------/
Public Function FileThere(theFileNameToTest As String) As Boolean

    If theFileNameToTest = "" Then FileThere = False: Exit Function

    If Len(Dir(theFileNameToTest)) = 0 Then
       FileThere = False
    Else
       FileThere = True
    End If
     
End Function


Public Function BrowseToMacro() As Workbook
    Set BrowseToMacro = Workbooks.Open(BrowseFilePath(D_EXCEL_MACRO))
End Function

Public Function BrowseFilePath(theType As getFileType) As String
'Gets path of text file for importing data

    Dim sFullName As String
    Dim sFileName As String

    sFullName = Application.GetOpenFilename(browse4type(theType))

    If sFullName = "False" Then End
    'Debug.Print sFullName, sFileName
    
    
    BrowseFilePath = sFullName
    
End Function




Public Function BrowseFilePaths(theType As getFileType) As String()

On Error GoTo BrowseFilePathsError

    Dim sFullName() As Variant
    sFullName() = Application.GetOpenFilename(browse4type(theType), , , , True)
    
    BrowseFilePaths = ConvertVariantToSTRArr(sFullName)

Exit Function

BrowseFilePathsError:
    End
    Dim errorStr(1 To 1) As String
    'errorStr = -1
    errorStr(1) = "No Selection"
    BrowseFilePaths = errorStr

End Function


Private Function browse4type(theType As getFileType) As String

   If theType = A_CSV Then browse4type = "*.csv,*.csv"
   If theType = B_EXCEL Then browse4type = "*.xlsx,*.xlsx"
   If theType = C_EXCEL_OLD Then browse4type = "*.xls,*.xls"
   If theType = D_EXCEL_MACRO Then browse4type = "*.xlsm,*.xlsm"
   If theType = E_InoFile Then browse4type = "*.ino,*.ino"
   If theType = F_Proc3File Then browse4type = "*.pde,*.pde"
   If theType = G_VBAModule Then browse4type = "*.bas,*.bas"
   If theType = H_Text Then browse4type = "*.txt,*.txt"
   If theType = I_Lib Then browse4type = "*.lbr,*.lbr"
' "Visual Basic Files (.bas; *.txt),.bas;*.txt"
End Function



' /--------------------------------------------\
' |converts a variant array to a string array  |
' \--------------------------------------------/
Function ConvertVariantToSTRArr(theVariant() As Variant) As String()
    
    Dim strARR() As String
    
    Dim x As Integer

    For x = LBound(theVariant) To UBound(theVariant)
        ReDim Preserve strARR(1 To x) As String
        strARR(x) = theVariant(x)
    Next x
    ConvertVariantToSTRArr = strARR

End Function







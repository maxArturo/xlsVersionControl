Attribute VB_Name = "mod_VersionControl"
Public FSO As FileSystemObject
Sub export_modules_for_version_control()
    
    Set objMyProj = Application.VBE.ActiveVBProject
    Set FSO = CreateObject("Scripting.FileSystemObject")
     
    For Each objVBComp In objMyProj.VBComponents
        
        'Modules are type 1, class modules type 2, forms are type 3
        If objVBComp.Type = 1 Or objVBComp.Type = 2 Then
            objVBComp.Export ActiveWorkbook.Path & "\" & objVBComp.Name & ".bas"
        End If
    Next
    
    originalFileName = ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name
    gitSaveLocation = ActiveWorkbook.Path & Application.PathSeparator & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "_git.zip"

    On Error Resume Next
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs gitSaveLocation
        ActiveWorkbook.SaveAs originalFileName
        UnzipAndPretty (gitSaveLocation)
        Application.DisplayAlerts = True

        If FSO.FileExists(gitSaveLocation) Then
          FSO.DeleteFile (gitSaveLocation)
        End If
    On Error GoTo 0
     
End Sub

Function UnzipAndPretty(sPath)

    Dim oApp
    Dim gitXMLFolder As Folder
            
    Set oApp = CreateObject("Shell.Application")
    
    git_folder = ActiveWorkbook.Path & Application.PathSeparator & "git_xslm"
    
    If FSO.FolderExists(git_folder) Then
        FSO.DeleteFolder (git_folder)
    End If
    FSO.CreateFolder (git_folder)
    
    oApp.Namespace(git_folder).CopyHere oApp.Namespace(sPath).Items
    
    Application.Wait (6000)
    
    Set gitXMLFolder = FSO.GetFolder(git_folder)
    Call cleanXMLinFolder(gitXMLFolder)
    
    Set oApp = Nothing
    Set FSO = Nothing
    
End Function


Function cleanXMLinFolder(mainFolder As Folder)
    Dim subfile As File
    Dim readFile As TextStream
    Dim writeFile As TextStream
    Dim inputXML As String
    Dim outputXML As String
    Dim recursiveFolder As Folder
    
    On Error Resume Next
    For Each subfile In mainFolder.Files
        If subfile.Type = "XML Document" Then
            Set readFile = FSO.OpenTextFile(subfile.Path, ForReading, False)
            inputXML = readFile.ReadAll
            readFile.Close
            outputXML = PrettyPrintXML(inputXML)
            Set writeFile = FSO.OpenTextFile(subfile.Path, ForWriting)
            writeFile.Write (outputXML)
            writeFile.Close
        End If
    Next
    On Error GoTo 0
    
    Set childFolders = mainFolder.SubFolders
    For Each recursiveFolder In childFolders
        Call cleanXMLinFolder(recursiveFolder)
    Next
End Function

Public Function PrettyPrintXML(XML As String) As String

  Dim Reader As New SAXXMLReader60
  Dim Writer As New MXXMLWriter60

  Writer.Indent = True
  Writer.standalone = False
  Writer.omitXMLDeclaration = False
  Writer.Encoding = "utf-8"

  Set Reader.contentHandler = Writer
  Set Reader.dtdHandler = Writer
  Set Reader.errorHandler = Writer

  Call Reader.putProperty("http://xml.org/sax/properties/declaration-handler", _
          Writer)
  Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", _
          Writer)

  Call Reader.Parse(XML)

  PrettyPrintXML = Writer.output

End Function

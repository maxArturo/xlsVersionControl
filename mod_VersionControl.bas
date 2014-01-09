Attribute VB_Name = "mod_VersionControl"
Public FSO As FileSystemObject
Sub export_modules_for_version_control()
    
    Set objMyProj = Application.VBE.ActiveVBProject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    vba_bas_folder = ActiveWorkbook.Path & Application.PathSeparator & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "_vba"
    
    If FSO.FolderExists(vba_bas_folder) Then
        FSO.DeleteFolder (vba_bas_folder)
    End If
    FSO.CreateFolder (vba_bas_folder)
     
    For Each objVBComp In objMyProj.VBComponents
        
        'Modules are type 1, class modules type 2, forms are type 3
        If objVBComp.Type = 1 Or objVBComp.Type = 2 Then
            objVBComp.Export vba_bas_folder & "\" & objVBComp.Name & ".bas"
        End If
    Next
    
    originalFileName = ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name
    gitsavelocation = ActiveWorkbook.Path & Application.PathSeparator & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "_git.zip"

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs gitsavelocation
    ActiveWorkbook.SaveAs originalFileName
    UnzipAndPretty (gitsavelocation)
    Application.DisplayAlerts = True

    If FSO.FileExists(gitsavelocation) Then
      FSO.DeleteFile (gitsavelocation)
    End If
    
    Set FSO = Nothing
     
End Sub

Function UnzipAndPretty(sPath)

    Dim oApp
    Dim gitXMLFolder As Folder
            
    Set oApp = CreateObject("Shell.Application")
    
    git_folder = ActiveWorkbook.Path & Application.PathSeparator & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & "_git_xslm"
    
    If FSO.FolderExists(git_folder) Then
        FSO.DeleteFolder (git_folder)
    End If
    FSO.CreateFolder (git_folder)
    
    oApp.Namespace(git_folder).CopyHere oApp.Namespace(sPath).Items
    
    Application.Wait (6000)
    
    Set gitXMLFolder = FSO.GetFolder(git_folder)
    Call cleanXMLinFolder(gitXMLFolder)
    
    Set oApp = Nothing
    
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

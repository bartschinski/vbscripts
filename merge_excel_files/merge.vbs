''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Merge Excel files
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The script can merge all files with .xls, .xlsx and .xlsm to one .xlsx file
' filename convention is (*_)BPNumber_Year.xls(x/m)...merged to BPNumber.xlsx with tabs as Year
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

MsgBox("Click OK to start the excel file merger")

Dim iMergedFiles()
Redim iMergedFiles(0)
iMergeFileCount = 0
lcScriptPath = WScript.ScriptFullName
lcScriptName = WScript.ScriptName
SFolder = Mid(lcScriptPath, 1, Len(lcScriptPath) - Len(lcScriptName)-1)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set Folder = FSO.GetFolder(SFolder)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' check if is the same file if not count one more
'
' @param sName  string  file name
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function addToFileCount(sName)
  iCheck = 0
  For i = 0 to UBound(iMergedFiles) 
    If iMergedFiles(i) = sName Then
      iCheck = 1
    End If
  Next 

  If iCheck = 0 Then
    Redim iMergedFiles(UBound(iMergedFiles) + 1)
    iMergedFiles(UBound(iMergedFiles)) = sName
  End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' replace the ' in numbers
'
' @param sFileLink  string  file link
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function replace_number(sFileLink)
  If FSO.GetFile(sFileLink).Size > 0 Then
    Set objFileRead = FSO.OpenTextFile(sFileLink, 1, True)
    sFileContent = objFileRead.ReadAll
    objFileRead.Close
    
    Set regEx = New RegExp
    regEx.Pattern = "([0-9])[\']([0-9])"
    regEx.IgnoreCase = True
    regEx.Global = True

    sFileContent = regEx.Replace(sFileContent, "$1$2")
    
    Set objFileWrite = FSO.OpenTextFile(sFileLink, 2, True)
    objFileWrite.Write(sFileContent)
    objFileWrite.Close
  end if
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' merge excel
'
' @param NFile      string  file link for new file
' @param sFileLink  string  file link for existing file
' @param sTabYear   string  year
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function merge(NFile, sFileLink, sTabYear)
  Set PXLS = CreateObject("Excel.Application") 
  
  If FSO.FileExists(NFile) Then
    Set PWB = PXLS.Workbooks.Open (NFile) 
  Else
    Set PWB = PXLS.Workbooks.Add() 
  End If
  
  Set PWS = PWB.Sheets.add
  PWS.Name = sTabYear

  Set CXLS = CreateObject("Excel.Application") 
  CXLS.DisplayAlerts = False
  Set CWB = CXLS.Workbooks.Open (sFilelink) 
  Set CWS = CWB.Sheets(1)

  CWS.Cells.Copy
  PWS.Range("A1").PasteSpecial -4163
  
  If FSO.FileExists(NFile) Then
    PXLS.ActiveWorkbook.Save 
  Else
    PWB.SaveAs(NFile)
    PXLS.ActiveWorkbook.Save 
  End If
  
  PXLS.ActiveWindow.Close 
  PXLS.Application.Quit 
  CXLS.ActiveWindow.Close 
  CXLS.Application.Quit 
End Function


For Each File in Folder.Files
  If Left(File.Name, 1) <> "~" And (Right(File.Name, 5) = ".xlsx" Or Right(File.Name, 5) = ".xlsm" Or Right(File.Name, 4) = ".xls") Then 
    If Right(File.Name, 5) = ".xlsx" Or Right(File.Name, 5) = ".xlsm" Then 
      sFilename = Mid(File.Name, 1, Len(File.Name)-5)
    End If
    If Right(File.Name, 4) = ".xls" Then 
      sFilename = Mid(File.Name, 1, Len(File.Name)-4)
    End If
    
    arFilename = Split(sFilename, "_")
    iFileCount = UBound(arFilename)

    If iFileCount >= 1 Then
      sDestinationFilename = arFilename(iFileCount - 1)
      NFile = SFolder + "\" + sDestinationFilename + ".xlsx"
      sFilelink = SFolder + "\" + File.Name
      
      addToFileCount(sDestinationFilename)
      replace_number(sFilelink)
      Call merge(NFile, sFilelink, arFilename(iFileCount))

      iMergeFileCount = iMergeFileCount + 1
    End If
  End If
Next

MsgBox("Excel file merge is finished"  & chr(13) & iMergeFileCount & " Files are merged to " & UBound(iMergedFiles) & " Files")

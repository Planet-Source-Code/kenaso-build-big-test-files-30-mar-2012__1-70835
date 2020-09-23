Attribute VB_Name = "modBuildFiles"
' ***************************************************************************
' Routine:   modMain
'
' Purpose:   Create big files filled with binary zeroes.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Created module
' 01-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated CreateBigFile() routine with edit checks and
'                information messages
'              - Added IsSpaceAvailable() routine
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME  As String = "modBuildFiles"
  Private Const ONE_GB       As String = "One_GB."
  Private Const CMD_PREFIX   As String = "cmd /c copy /b /y "   ' NT based systems use "cmd"
                                                                ' else use "command"
  Private Const GB_1         As Currency = 1073741824@
  Private Const SYNCHRONIZE  As Long = &H100000
  Private Const INFINITE     As Long = &HFFFF
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' Waits until the specified object is in the signaled state or the time-out
  ' interval elapses.  If the function fails, the return value is -1.
  Private Declare Function WaitForSingleObject Lib "kernel32" _
          (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

  ' Opens an existing local process object.  If the function succeeds, the
  ' return value is an open handle to the specified process.  If function
  ' fails, a null value is returned.
  Private Declare Function OpenProcess Lib "kernel32" _
          (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
          ByVal dwProcessId As Long) As Long

  ' Closes an open object handle.
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long

                              
' ***************************************************************************
' Routine:       CreateBigFile
'
' Description:   Create one or more one gigabyte files filled with binary
'                zeroes.  If requested, then concatenate them into a single
'                file and delete the smaller files.  If concatenating, you
'                must have double the free space.
'
' Parameters:    strPath - Target path where files are to be created
'                lngFileCnt - Number of files to be created
'                blnOneBigFile - Flag designating if one big file should be
'                   created.  Flag is ignored if only one file is created.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Feb-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 01-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated with edit checks and messages
' ***************************************************************************
Public Sub CreateBigFile(ByVal strPath As String, _
                         ByVal lngFileCnt As Long, _
                         ByVal blnOneBigFile As Boolean)
    
    Dim hFile          As Long
    Dim lngIndex       As Long
    Dim lngProcessID   As Long
    Dim strCmdLine     As String
    Dim strPathFile    As String
    Dim strDestFile    As String
    Dim strAppendFiles As String
    
    Const ROUTINE_NAME As String = "CreateBigFile"

    On Error GoTo CreateBigFile_Error

    ' If no path is passed then
    ' use application folder
    If Len(Trim$(strPath)) = 0 Then
        strPath = App.Path
    End If
    
    strAppendFiles = vbNullString
    lngProcessID = 0                ' Set process ID to zero
    strPath = QualifyPath(strPath)  ' Verify a trailing backslash
    
    ' Is there enough space to create these files?
    If IsSpaceAvailable(strPath, lngFileCnt, False) Then
    
        ' Loop thru and create files
        For lngIndex = 1 To lngFileCnt
            
            DoEvents
            strPathFile = strPath & ONE_GB & Format$(lngIndex, "000")
                             
            If IsPathValid(strPathFile) Then
                ' Verify target file is empty
                hFile = FreeFile                        ' Get first free file handle
                DoEvents
                Open strPathFile For Output As #hFile   ' Create empty file
                DoEvents
                Close #hFile                            ' Close file
            End If
            
            ' The character being used (ex: 0) only designates the
            ' last ASCII value in the file.  All previous values
            ' will be null (ASCII 0) values.
            '
            ' File size should be 1gb or smaller for optimal use.
            hFile = FreeFile                                     ' Get first free file handle
            Open strPathFile For Binary Access Write As #hFile   ' Open for writing
            Put #hFile, GB_1, Chr$(0)                            ' Fill with binary 0's
            Close #hFile                                         ' Close file
            
        Next lngIndex
                                
        ' See if multiple files were created
        If lngFileCnt = 1 Then
                        
            InfoMsg "A one gigabyte file has been created." & _
                    vbNewLine & vbNewLine & strDestFile, "Big File Creation"
        
        Else  ' More than one file was created
             
            ' Are we creating one big file from smaller files?
            If blnOneBigFile Then
                 
                ' If not enough space is available
                ' then remove files that were created.
                If Not IsSpaceAvailable(strPath, lngFileCnt, True) Then
                
                    On Error Resume Next
                    
                    ' Loop thru list of 1gb
                    ' files and delete them
                    For lngIndex = 1 To lngFileCnt
                                
                        DoEvents
                        strAppendFiles = strPath & ONE_GB & Format$(lngIndex, "000")
                        Kill strAppendFiles
                        DoEvents
                        strAppendFiles = vbNullString
                        
                    Next lngIndex
                    
                    GoTo CreateBigFile_CleanUp
                    
                End If
                    
                ' Concatenate list of file names
                For lngIndex = 1 To lngFileCnt
                            
                    strAppendFiles = strAppendFiles & strPath & ONE_GB & Format$(lngIndex, "000")
                    
                    If lngIndex <> lngFileCnt Then
                        
                        ' Add a plus sign for appending
                        ' one file to another
                        strAppendFiles = strAppendFiles & "+"
                    
                    End If
                    
                Next lngIndex
                                
                ' Destination path and file name
                strDestFile = strPath & "GB_" & CStr(lngFileCnt) & ".dat"
                
                ' If file already exist then delete it
                If IsPathValid(strDestFile) Then
                    Kill strDestFile
                    DoEvents
                End If
                
                ' Append final destination file
                strCmdLine = CMD_PREFIX & strAppendFiles & " " & strDestFile
                
                frmMain.Hide                                     ' Hide window form because DOS window will open
                lngProcessID = Shell(strCmdLine, vbNormalFocus)  ' Open DOS window to create one big file
                WaitUntilFinished lngProcessID                   ' Wait until processing is finished
                CloseHandle lngProcessID                         ' Always close process ID handle
                lngProcessID = 0                                 ' Reset process ID value to zero
                strAppendFiles = vbNullString
                frmMain.Show                                     ' Display window form
                
                On Error Resume Next
                    
                ' Loop thru list of 1gb
                ' files and delete them
                For lngIndex = 1 To lngFileCnt
                            
                    DoEvents
                    strAppendFiles = strPath & ONE_GB & Format$(lngIndex, "000")
                    Kill strAppendFiles
                    DoEvents
                    strAppendFiles = vbNullString
                    
                Next lngIndex
                
                ' ex:  "One 3 gigabyte file created."
                InfoMsg "One " & CStr(lngFileCnt) & " gigabyte file created." & _
                        vbNewLine & vbNewLine & strDestFile, "Big File Creation"
                        
            Else
            
                ' ex:  "3 one gigabyte files have been created."
                InfoMsg CStr(lngFileCnt) & " one gigabyte files have been created." & _
                        vbNewLine & vbNewLine & strDestFile, "Big File Creation"
            End If
        End If
    End If

CreateBigFile_CleanUp:
    On Error GoTo 0
    Exit Sub

CreateBigFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume CreateBigFile_CleanUp
    
End Sub



' ***************************************************************************
' ****                Internal Procedures and Functions                  ****
' ***************************************************************************

' ***************************************************************************
' Routine:       WaitUntilFinished
'
' Description:   Wait until a process is finished.
'
' Parameters:    lngProcessID - Numeric value of the target process
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub WaitUntilFinished(ByVal lngProcessID As Long)
    
    Dim lngHwnd As Long
    
    ' Capture process handle
    lngHwnd = OpenProcess(SYNCHRONIZE, 0, lngProcessID)
    
    ' If process is still active then the
    ' handle will be greater than zero
    DoEvents
    If lngHwnd > 0 Then
        WaitForSingleObject lngHwnd, INFINITE  ' Wait until process is finished
        CloseHandle lngHwnd                    ' Always close process handle
        lngHwnd = 0                            ' Reset handle to zero
    End If
    
End Sub

Private Function IsSpaceAvailable(ByVal strPath As String, _
                                  ByVal lngFileCnt As Long, _
                                  ByVal blnAppend As Boolean) As Boolean

    Dim strDrive       As String
    Dim curSpaceAvail  As Currency
    Dim curSpaceNeeded As Currency
    Dim objFSO         As Scripting.FileSystemObject
    Dim objDrive       As Drive
    
    Const FMT_STR1 As String = "!@@@@@@@@@@"            ' 10 space holders (left justified)
    Const FMT_STR2 As String = "@@@@@@@@@@@@@@@@@@@@"   ' 20 space holders (right justified)
                               ' 999,999,999,999,999
    On Error Resume Next
    
    IsSpaceAvailable = False                   ' Preset flag to FALSE
    strDrive = QualifyPath(Left$(strPath, 2))  ' Format drive letter (ex: "C:\")
        
    If IsPathValid(strPath) Then
    
        Set objFSO = New Scripting.FileSystemObject  ' Instantiate objects
        Set objDrive = objFSO.GetDrive(strDrive)
        
        curSpaceAvail = CCur(objDrive.FreeSpace)     ' get free space available
        
        Set objDrive = Nothing  ' Free objects from memory
        Set objFSO = Nothing
        
        ' Calc amount of space needed
        If blnAppend Then
            ' Concatenating multiple files into one
            curSpaceNeeded = CCur(GB_1 * (lngFileCnt * 2))
        Else
            ' Creating single files
            curSpaceNeeded = CCur(GB_1 * lngFileCnt)
        End If
        
        ' Enough space available?
        If curSpaceNeeded < curSpaceAvail Then
            IsSpaceAvailable = True   ' Set flag to TRUE
        Else
            ' Not enough free space
            If blnAppend Then
                InfoMsg "Not enough free space available to " & vbNewLine & _
                        "concatenate " & CStr(lngFileCnt) & " one gigabyte files " & vbNewLine & _
                        "into a single big file." & vbNewLine & vbNewLine & _
                        Format$("Available:", FMT_STR1) & Format$(Format$(curSpaceAvail, "#,##0"), FMT_STR2) & vbNewLine & _
                        Format$("Needed:", FMT_STR1) & Format$(Format$(curSpaceNeeded, "#,##0"), FMT_STR2)
            Else
                InfoMsg "Not enough free space available to" & vbNewLine & _
                        "create " & CStr(lngFileCnt) & " one gigabyte file(s)." & _
                        vbNewLine & vbNewLine & _
                        Format$("Available:", FMT_STR1) & Format$(Format$(curSpaceAvail, "#,##0"), FMT_STR2) & vbNewLine & _
                        Format$("Needed:", FMT_STR1) & Format$(Format$(curSpaceNeeded, "#,##0"), FMT_STR2)
            End If
        End If
    
    Else
        ' Invalid drive designation
        InfoMsg "Invalid path." & vbNewLine & vbNewLine & strPath
    End If
    
    On Error GoTo 0   ' Nullify this error trap
    
End Function

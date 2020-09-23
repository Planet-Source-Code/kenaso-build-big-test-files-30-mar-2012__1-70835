Attribute VB_Name = "modMain"
' ***************************************************************************
' Module:        modMain
'
' Description:   This is a generic module I use to start and stop an
'                application
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced FileExists() and PathExists() routines with
'              IsPathValid() routine.
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Deleted RemoveTrailingNulls() routine from this module. 
'              - Changed call to RemoveTrailingNulls() to TrimStr module 
'                due to speed and accuracy.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Global constants
' ***************************************************************************
  Public Const AUTHOR_NAME           As String = "Kenneth Ives"
  Public Const AUTHOR_MSG            As String = "Click to send email to "
  Public Const AUTHOR_EMAIL          As String = "kenaso@tx.rr.com"
  Public Const PGM_NAME              As String = "Build_Big_Files"
  Public Const MAX_SIZE              As Long = 260

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const ERROR_ALREADY_EXISTS As Long = 183&
  Private Const MODULE_NAME          As String = "modMain"

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long
  
  ' The GetCurrentProcess function returns a pseudohandle for the current
  ' process. A pseudohandle is a special constant that is interpreted as
  ' the current process handle. The calling process can use this handle to
  ' specify its own process whenever a process handle is required. The
  ' pseudohandle need not be closed when it is no longer needed.
  Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
  
  ' The GetExitCodeProcess function retrieves the termination status of the
  ' specified process. If the function succeeds, the return value is nonzero.
  Private Declare Function GetExitCodeProcess Lib "kernel32" _
          (ByVal hProcess As Long, lpExitCode As Long) As Long
  
  ' ExitProcess function ends a process and all its threads
  ' ex:     ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
  Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
  
  ' The CreateMutex function creates a named or unnamed mutex object.  Used
  ' to determine if an application is active.
  Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" _
          (lpMutexAttributes As Any, ByVal bInitialOwner As Long, _
          ByVal lpName As String) As Long
  
  ' This function releases ownership of the specified mutex object.
  ' Finished with the search.
  Private Declare Function ReleaseMutex Lib "kernel32" _
          (ByVal hMutex As Long) As Long

  ' The ShellExecute function opens or prints a specified file.  The file
  ' can be an executable file or a document file.
  Private Declare Function ShellExecute Lib "shell32.dll" _
          Alias "ShellExecuteA" (ByVal hwnd As Long, _
          ByVal lpOperation As String, ByVal lpFile As String, _
          ByVal lpParameters As String, ByVal lpDirectory As String, _
          ByVal nShowCmd As Long) As Long

  ' Always close a handle if not being used
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long
  
  ' Truncates a path to fit within a certain number of characters by replacing
  ' path components with ellipses.  Called by ShrinkTofit().
  Private Declare Function PathCompactPathEx Lib "shlwapi.dll" _
          Alias "PathCompactPathExA" _
          (ByVal pszOut As String, ByVal pszSrc As String, _
          ByVal cchMax As Long, ByVal dwFlags As Long) As Long

' ***************************************************************************
' Global Variables
'
'                    +-------------- Global level designator
'                    |  +----------- Data type (String)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   g str Version
' Variable name:     gstrVersion
'
' ***************************************************************************
  Public gstrVersion         As String
  Public gblnIDE_Environment As Boolean  ' TRUE if in VB IDE



' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Main
'
' Description:   This is a generic routine to start an application
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Main()

    Const ROUTINE_NAME As String = "Main"

    On Error Resume Next
    ChDrive App.Path
    ChDir App.Path
    On Error GoTo 0
    
    On Error GoTo Main_Error
    
    ' See if there is another instance of this program
    ' running.  The parameter being passed is the name
    ' of this executable without the EXE extension.
    If Not AlreadyRunning(App.EXEName) Then
        gblnStopProcessing = False    ' preset global stop flag
        Load frmMain                  ' Load main form
    End If

Main_CleanUp:
    On Error GoTo 0
    Exit Sub

Main_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume Main_CleanUp
    
End Sub

' ***************************************************************************
' Routine:       TerminateProgram
'
' Description:   This routine will perform the shutdown process for this
'                application.  The proper sequence to follow is:
'
'                    1.  Deactivate and free from memory all global objects
'                        or classes
'                    2.  Verify there are no file handles left open
'                    3.  Deactivate and free from memory all form objects
'                    4.  Shut this application down
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub TerminateProgram()

    ' Free any global objects from memory.
    ' EXAMPLE:    Set gobjFSO = Nothing
    
    CloseAllFiles   ' close any open files accessed by this application
    UnloadAllForms  ' Unload any forms from memory

    ' While in the VB IDE (integrated developement environment),
    ' do not call ExitProcess API.  ExitProcess API will close all
    ' processes associated with this application.  This will close
    ' the VB IDE immediately and no changes will be saved that were
    ' not previously saved.
    If gblnIDE_Environment Then
        End    ' Terminate this application while in the VB IDE
    Else
        ExitProcess GetExitCodeProcess(GetCurrentProcess, 0)
    End If

End Sub
 
' ***************************************************************************
' Routine:       CloseAllFiles
'
' Description:   Closes any files that were opened within this application.
'                The FreeFile() function returns an integer representing the
'                next file handle opened by this application.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function CloseAllFiles() As Boolean

    While FreeFile > 1
        Close #FreeFile - 1
    Wend
    
End Function

' ***************************************************************************
' Routine:       UnloadAllForms
'
' Description:   Unload all active forms associated with this application.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-APR-2001  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub UnloadAllForms()

    Dim frm As Form
    Dim ctl As Control

    ' Loop thru all active forms
    ' associated with this application
    For Each frm In Forms
        
        frm.Hide            ' hide selected form
        
        ' free all controls from memory
        For Each ctl In frm.Controls
            Set ctl = Nothing
        Next
        
        Unload frm          ' deactivate form object
        Set frm = Nothing   ' free form object from memory
                            ' (prevents memory fragmenting)
    Next frm

End Sub

' ***************************************************************************
' Routine:       FindRequiredFile
'
' Description:   Test to see if a required file is in the application folder
'                or in any of the folders in the PATH environment variable.
'
' Parameters:    strFilename - name of the file without path information
'                strFullPath - Optional - If found then the fully qualified
'                     path and filename are returned
'
' Returns:       TRUE  - Found the required file
'                FALSE - File could not be found
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 04-Apr-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function FindRequiredFile(ByVal strFilename As String, _
                        Optional ByRef strFullPath As String = vbNullString) As Boolean

    Dim strPath     As String    ' Fully qualified search path
    Dim strMsgFmt   As String    ' Format each message line
    Dim strDosPath  As String    ' DOS environment variable
    Dim strSearched As String    ' List of searched folders (will be displayed if not found)
    Dim lngPointer  As Long      ' String pointer position
    Dim blnFoundIt  As Boolean   ' Flag (TRUE if found file else FALSE)

    
    On Error GoTo FindRequiredFile_Error

    strFullPath = vbNullString    ' Empty return variable
    strSearched = vbNullString
    strMsgFmt = "!" & String$(70, "@")
    blnFoundIt = False  ' Preset flag to FALSE
    lngPointer = 0
                  
    ' Prepare path for application folder
    strPath = QualifyPath(App.Path)
    
    ' Check application folder
    If IsPathValid(strPath & strFilename) Then
        
        blnFoundIt = True  ' Found in application folder
        
    Else
        ' Capture DOS environment variable
        ' so the PATH can be searched
        '
        ' Save application path to searched list
        strSearched = strPath & vbNewLine
    
        ' Capture environment variable PATH statement
        strDosPath = TrimStr(Environ$("PATH"))
        
        If Len(strDosPath) > 0 Then

            ' append semi-colon
            strDosPath = QualifyPath(strDosPath, ";")
            
            Do
                ' Find first semi-colon
                lngPointer = InStr(1, strDosPath, ";")
                
                ' Did we find a semi-colon?
                If lngPointer > 0 Then
                    
                    strPath = Mid$(strDosPath, 1, lngPointer - 1)  ' Capture path
                    strPath = GetLongName(strPath)                 ' Format path name
                    
                    If Len(strPath) > 0 Then
                    
                        strPath = QualifyPath(strPath)                 ' Append backslash
                        strDosPath = Mid$(strDosPath, lngPointer + 1)  ' Resize path string
                        
                        ' Add path to searched list
                        strSearched = strSearched & Format$(strPath, strMsgFmt) & vbNewLine
                        
                        ' See if the file is in this folder
                        If IsPathValid(strPath & strFilename) Then
                            blnFoundIt = True   ' Success
                            Exit Do             ' Exit this loop
                        End If
                        
                    End If
                End If
                
            Loop While lngPointer > 0
            
        Else
            strSearched = Format$(strSearched, strMsgFmt) & vbNewLine & _
                          Format$("PATH environment variable does not exists.", strMsgFmt) & vbNewLine
        End If
    End If
    
FindRequiredFile_CleanUp:
    If blnFoundIt Then
        strFullPath = strPath & strFilename   ' Return full path/filename
    Else
        InfoMsg Format$("A required file that supports this application cannot be found.", strMsgFmt) & _
                vbNewLine & vbNewLine & _
                Format$(Chr$(34) & UCase$(strFilename) & Chr$(34) & _
                " not in any of these folders:", strMsgFmt) & vbNewLine & vbNewLine & _
                strSearched, "File not found"
    End If
    
    FindRequiredFile = blnFoundIt   ' Set status flag
    On Error GoTo 0                 ' Nullify this error trap
    Exit Function

FindRequiredFile_Error:
    If Err.Number <> 0 Then
        Err.Clear
    End If

    Resume FindRequiredFile_CleanUp
  
End Function

' ***************************************************************************
' Procedure:     GetLongName
'
' Description:   The Dir() function can be used to return a long filename
'                but it does not include path information. By parsing a
'                given short path/filename into its constituent directories,
'                you can use the Dir() function to build a long path/filename.
'
' Example:       Syntax:
'                   GetLongName C:\DOCUME~1\KENASO\LOCALS~1\Temp\~ki6A.tmp
'
'                Returns:
'                   "C:\Documents and Settings\Kenaso\Local Settings\Temp\~ki6A.tmp"
'
' Parameters:    strShortName - Path or file name to be converted.
'
' Returns:       A readable path or file name.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jul-2004  http://support.microsoft.com/kb/154822
'              "How To Get a Long Filename from a Short Filename"
' 09-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 09-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added removal of all double quotes prior to formatting
' ***************************************************************************
Public Function GetLongName(ByVal strShortName As String) As String

    Dim strTemp     As String
    Dim strLongName As String
    Dim intPosition As Integer
    
    On Error Resume Next
    
    GetLongName = vbNullString
    strLongName = vbNullString
    
    ' Remove all double quotes
    strShortName = Replace(strShortName, Chr$(34), "")
    
    ' Add a backslash to short name, if needed,
    ' to prevent Instr() function from failing.
    strShortName = QualifyPath(strShortName)
    
    ' Start at position 4 so as to ignore
    ' "[Drive Letter]:\" characters.
    intPosition = InStr(4, strShortName, "\")
    
    ' Pull out each string between
    ' backslash character for conversion.
    Do While intPosition > 0
        
        strTemp = vbNullString   ' Init variable
        
        ' Progressively parse path to verify
        ' each portion does exist and
        ' capture its expanded version.
        strTemp = Dir$(Left$(strShortName, intPosition - 1), _
                       vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)
        
        ' If no data then exit this loop
        If Len(Trim$(strTemp)) = 0 Then
            strShortName = vbNullString
            strLongName = vbNullString
            Exit Do   ' exit DO..LOOP
        End If
        
        ' Append new elongated portion to output string
        ' after converting it to propercase format.
        strLongName = strLongName & "\" & StrConv(strTemp, vbProperCase)
        
        ' Find next backslash
        intPosition = InStr(intPosition + 1, strShortName, "\")
    
    Loop
    
GetLongName_CleanUp:
    If Len(strShortName & strLongName) > 0 Then
        GetLongName = UCase$(Left$(strShortName, 2)) & strLongName
    Else
        GetLongName = "[Unknown]"
    End If
    
    On Error GoTo 0   ' Nullify this error trap
    
End Function

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Determines whether a path to a file system object such as
'                a file or directory is valid. This function tests the
'                validity of the path. A path specified by Universal Naming
'                Convention (UNC) is limited to a file only; that is,
'                \\server\share\file is permitted. A UNC path to a server
'                or server share is not permitted; that is, \\server or
'                \\server\share. This function returns FALSE if a mounted
'                remote drive is out of service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'
' Reference:     http://msdn.microsoft.com/en-us/library/bb773584(v=vs.85).aspx
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))
   
End Function
 
' ***************************************************************************
' Routine:       AlreadyRunning
'
' Description:   This routine will determine if an application is already
'                active, whether it be hidden, minimized, or displayed.
'
' Parameters:    strTitle - partial/full name of application
'
' Returns:       TRUE  - Currently active
'                FALSE - Inactive
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-DEC-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function AlreadyRunning(ByVal strAppTitle As String) As Boolean
        
    Dim hMutex As Long
    
    Const ROUTINE_NAME As String = "AlreadyRunning"

    On Error GoTo AlreadyRunning_Error

    gblnIDE_Environment = False  ' preset flags to FALSE
    AlreadyRunning = False

    ' Are we in VB development environment?
    gblnIDE_Environment = IsVB_IDE
    
    ' Multiple instances can be run while
    ' in the VB IDE but not as an EXE
    If Not gblnIDE_Environment Then

        ' Try to create a new Mutex handle
        hMutex = CreateMutex(ByVal 0&, 1, strAppTitle)
        
        ' Did mutex handle already exist?
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
             
            ReleaseMutex hMutex     ' Release Mutex handle from memory
            CloseHandle hMutex      ' Close the Mutex handle
            Err.Clear               ' Clear any errors
            AlreadyRunning = True   ' prior version already active
        End If
    End If

AlreadyRunning_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

AlreadyRunning_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume AlreadyRunning_CleanUp

End Function

Private Function IsVB_IDE() As Boolean
    
    ' 09-16-2000  Michael Culley  m_culley@one.net.au
    '             http://forums.devx.com/showthread.php?t=37676
    '
    ' Set DebugMode flag.  Call can only be
    ' successful if in VB development environment.
    Debug.Assert SetTrue(IsVB_IDE) Or True

End Function

Private Function SetTrue(ByRef blnValue As Boolean) As Boolean
    
    ' Can only be set to TRUE if Debug.Assert call
    ' call is successful.  Call can only be
    ' successful if in VB development environment.
    blnValue = True

End Function

' ***************************************************************************
' Routine:       QualifyPath
'
' Description:   Adds a trailing character to the path, if missing.
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to append.
'                          Default = "\"
'
' Returns:       Fully qualified path with a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Function QualifyPath(ByVal strPath As String, _
                   Optional ByVal strChar As String = "\") As String

    strPath = Trim$(strPath)
    
    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        QualifyPath = strPath
    Else
        QualifyPath = strPath & strChar
    End If
    
End Function

' ***************************************************************************
' Routine:       UnQualifyPath
'
' Description:   Removes a trailing character from the path
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to remove
'                          Default = "\"
'
' Returns:       Fully qualified path without a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Function UnQualifyPath(ByVal strPath As String, _
                     Optional ByVal strChar As String = "\") As String

    strPath = Trim$(strPath)
    
    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        UnQualifyPath = Left$(strPath, Len(strPath) - 1)
    Else
        UnQualifyPath = strPath
    End If
    
End Function

' ***************************************************************************
' Routine:       SendEmail
'
' Description:   When the email hyperlink is clicked, this routine will fire.
'                It will create a new email message with the author's name in
'                the "To:" box and the name and version of the application
'                on the "Subject:" line.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-FEB-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Sub SendEmail()

    Dim strMail As String

    Const ROUTINE_NAME As String = "SendEmail"

    On Error GoTo SendEmail_Error

    ' Create email heading for user
    strMail = "mailto:" & AUTHOR_EMAIL & "?subject=" & gstrVersion

    ' Call ShellExecute() API to create an email to the author
    ShellExecute 0&, vbNullString, strMail, _
                 vbNullString, vbNullString, vbNormalFocus

SendEmail_CleanUp:
    On Error GoTo 0
    Exit Sub

SendEmail_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume SendEmail_CleanUp

End Sub

' ***************************************************************************
' Routine:       ShrinkToFit
'
' Description:   This routine creates the ellipsed string by specifying
'                the size of the desired string in characters.  Adds
'                ellipses to a file path whose maximum length is specified
'                in characters.
'
' Parameters:    strPath - Path to be resized for display
'                intMaxLength - Maximum length of the return string
'
' Returns:       Resized path
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 20-May-2004  Randy Birch
'              http://vbnet.mvps.org/code/fileapi/pathcompactpathex.htm
' 22-Jun-2004  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Function ShrinkToFit(ByVal strPath As String, _
                            ByVal intMaxLength As Integer) As String

    Dim strBuffer As String
    
    strPath = TrimStr(strPath)
    
    ' See if ellipses need to be inserted into the path
    If Len(strPath) <= intMaxLength Then
        ShrinkToFit = strPath
        Exit Function
    End If
    
    ' intMaxLength is the maximum number of characters to be contained in the
    ' new string, **including the terminating NULL character**. For example,
    ' if intMaxLength = 8, the resulting string would contain a maximum of
    ' seven characters plus the termnating null.
    '
    ' Because of this, add 1 to the value passed as intMaxLength to ensure
    ' the resulting string is the size requested.
    intMaxLength = intMaxLength + 1
    strBuffer = Space$(MAX_SIZE)
    PathCompactPathEx strBuffer, strPath, intMaxLength, 0&
    
    ' Return the readjusted data string
    ShrinkToFit = TrimStr(strBuffer)
    
End Function

Public Sub Wait(ByVal sngSecond As Single)

    Dim sngTarget As Single
        
    sngTarget = Timer  ' Capture current time
        
    Do While (Timer - sngTarget) < sngSecond
            
        ' If we cross midnight, back up one day
        DoEvents
        If Timer < sngTarget Then
            sngTarget = sngTarget - 86400  ' Number of seconds in a day
        End If
        
    Loop
    
End Sub


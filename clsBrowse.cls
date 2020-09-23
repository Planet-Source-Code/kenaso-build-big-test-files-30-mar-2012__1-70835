VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsBrowse
'
' Description:   This class is used to browse for a folder and other generic
'                routines
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote class
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Created flag BIF_FOLDERSONLY to display folder selection
'              window regardless of Windows version.
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Deleted RemoveTrailingNulls() routine from this module. 
'              - Changed call to RemoveTrailingNulls() to TrimStr module 
'                due to speed and accuracy.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants - Miscellaneous
' ***************************************************************************
  Private Const MODULE_NAME           As String = "clsBrowse"
  Private Const DEFAULT_TITLE         As String = "Browse for a folder"
  Private Const MAX_SIZE              As Long = 260

' ***************************************************************************
' Constants used for Coloring progress bar
' ***************************************************************************
  Private Const WM_USER               As Long = &H400
  Private Const CCM_FIRST             As Long = &H2000&
  Private Const CCM_SETBKCOLOR        As Long = (CCM_FIRST + 1)
  Private Const PBM_SETBKCOLOR        As Long = CCM_SETBKCOLOR
  Private Const PBM_SETBARCOLOR       As Long = (WM_USER + 9)

' ***************************************************************************
' Constants used for browsing for a folder
' ***************************************************************************
  Private Const BIF_RETURNONLYFSDIRS  As Long = &H1&      ' only file system directories
  Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2&      ' no network folders below domain level
  Private Const BIF_STATUSTEXT        As Long = &H4&      ' include status area for callback
  Private Const BIF_RETURNFSANCESTORS As Long = &H8&      ' only return file system ancestors
  Private Const BIF_NEWDIALOGSTYLE    As Long = &H40&     ' use the new dialog layout
  Private Const BIF_NONEWFOLDERBUTTON As Long = &H200&
  Private Const BIF_FOLDERSONLY       As Long = BIF_RETURNONLYFSDIRS Or _
                                                BIF_DONTGOBELOWDOMAIN Or _
                                                BIF_STATUSTEXT Or _
                                                BIF_RETURNFSANCESTORS Or _
                                                BIF_NEWDIALOGSTYLE Or _
                                                BIF_NONEWFOLDERBUTTON

' ***************************************************************************
' Type structures used for browsing for a folder
' ***************************************************************************
  ' Contains parameters for the SHBrowseForFolder function and receives
  ' information about the folder selected by the user.
  Private Type BROWSEINFO
      hOwner         As Long
      pidlRoot       As Long
      pszDisplayName As String
      lpszTitle      As String
      ulFlags        As Long
      lpfn           As Long
      lParam         As Long
      iImage         As Long
  End Type

' ***************************************************************************
' API Declares used for changing color of a progress bar
' ***************************************************************************
  ' The SendMessage function sends the specified message to a window or
  ' windows. The function calls the window procedure for the specified
  ' window and does not return until the window procedure has processed
  ' the message.
  Private Declare Function SendMessage Lib "user32" _
          Alias "SendMessageA" _
          (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, lParam As Any) As Long

' ***************************************************************************
' API Declares used for browsing for a folder
' ***************************************************************************
  ' Converts an item identifier list to a file system path.
  Private Declare Function SHGetPathFromIDList Lib "shell32" _
          Alias "SHGetPathFromIDListA" _
          (ByVal pidl As Long, ByVal pszPath As String) As Long

  ' Displays a dialog box that enables the user to select a shell folder.
  Private Declare Function SHBrowseForFolder Lib "shell32" _
          Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

  ' Frees a block of task memory previously allocated through a call
  ' to the CoTaskMemAlloc or CoTaskMemRealloc function.
  Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

  ' Truncates a path to fit within a certain number of characters by replacing
  ' path components with ellipses.
  Private Declare Function PathCompactPathEx Lib "shlwapi.dll" _
          Alias "PathCompactPathExA" _
          (ByVal pszOut As String, ByVal pszSrc As String, _
          ByVal cchMax As Long, ByVal dwFlags As Long) As Long


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       BrowseForFolder
'
' Description:   This function will open the folder browse dialog box.
'
' Parameters:    frm - Form that is calling this routine
'                strTitle - [Optional] Title to be displayed on the dialog
'                      box.  Uses default title if none is provided.
'
' Returns:       Name of folder selected.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch http://www.mvps.org/vbnet/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added new flag BIF_FOLDERSONLY to display folder selection
'              window regardless of Windows version.
' ***************************************************************************
Public Function BrowseForFolder(ByRef frm As Form, _
                       Optional ByVal strTitle As String = DEFAULT_TITLE) As String
Attribute BrowseForFolder.VB_Description = "Shutdown the operating system."

    Dim typBI         As BROWSEINFO
    Dim strPath       As String
    Dim lngPathHandle As Long

    On Error GoTo BrowseForFolder_Error
    
    With typBI
        ' Hwnd of the window that receives messages from the call. Can be your
        ' application or the handle from GetDesktopWindow().
        .hOwner = frm.hwnd
    
        ' Pointer to the item identifier list specifying the location of the "root"
        ' folder to browse from.  If NULL, the desktop folder is used.
        .pidlRoot = 0&
    
        .lpszTitle = strTitle       ' message to be displayed in the Browse dialog
        .ulFlags = BIF_FOLDERSONLY  ' the type of folder to return
    End With
    
    lngPathHandle = SHBrowseForFolder(typBI) ' show the browse for folders dialog

    ' the dialog has closed, so parse & display the user's returned folder
    ' selection contained in lngPathHandle
    strPath = Space$(MAX_SIZE)

    ' Remove all trailing nulls from the folder selected
    If SHGetPathFromIDList(ByVal lngPathHandle, ByVal strPath) Then
        strPath = TrimStr(strPath)
    Else
        strPath = vbNullString
    End If

BrowseForFolder_CleanUp:
    ' Always close any open handles when not in use
    CoTaskMemFree lngPathHandle
    BrowseForFolder = strPath
    On Error GoTo 0
    Exit Function

BrowseForFolder_Error:
    ErrorMsg MODULE_NAME, "BrowseForFolder", Err.Description
    strPath = vbNullString
    Resume BrowseForFolder_CleanUp

End Function

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
    ' Because of this, one has been added to the value passed as intMaxLength
    ' to ensure the resulting string is the size requested.
    intMaxLength = intMaxLength + 1
    strBuffer = Space$(MAX_SIZE)
    PathCompactPathEx strBuffer, strPath, intMaxLength, 0&
    
    ' Return the readjusted data string
    ShrinkToFit = TrimStr(strBuffer)
    
End Function

' ***************************************************************************
' Routine:       SetPBarForegroundColor
'
' Description:   Set the Microsoft progress bar progression color
'
' Parameters:    lngPBarHwnd - Handle designating the progress bar
'                lngColor - long integer representing the color desired
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch http://www.mvps.org/vbnet/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Sub SetPBarForegroundColor(ByRef lngPBarHwnd As Long, _
                                  ByVal lngColor As Long)

    On Error GoTo SetPBarForegroundColor_Error

    ' Change progress color
    '
    ' Syntax:
    '   SetPBarForegroundColor ProgressBar.hwnd, RGB(205, 0, 0)  ' red
    SendMessage lngPBarHwnd, PBM_SETBARCOLOR, 0&, ByVal lngColor

SetPBarForegroundColor_CleanUp:
    On Error GoTo 0
    Exit Sub

SetPBarForegroundColor_Error:
    ErrorMsg MODULE_NAME, "SetPBarForegroundColor", Err.Description
    Resume SetPBarForegroundColor_CleanUp

End Sub

' ***************************************************************************
' Routine:       SetPBarBackgroundColor
'
' Description:   Set the Microsoft progress bar background color
'
' Parameters:    lngPBarHwnd - Handle designating the progress bar
'                lngColor - long integer representing the color desired
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch http://www.mvps.org/vbnet/index.html
'              Original routine
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Public Sub SetPBarBackgroundColor(ByRef lngPBarHwnd As Long, _
                                  ByVal lngColor As Long)

    On Error GoTo SetPBarBackgroundColor_Error

    ' Change background color
    '
    ' With CommonDialog1
    '      .CancelError = True
    '      .ShowColor
    '      SetPBarBackgroundColor ProgressBar.hwnd, .Color
    ' End With
    SendMessage lngPBarHwnd, PBM_SETBKCOLOR, 0&, ByVal lngColor

SetPBarBackgroundColor_CleanUp:
    On Error GoTo 0
    Exit Sub

SetPBarBackgroundColor_Error:
    ErrorMsg MODULE_NAME, "SetPBarBackgroundColor", Err.Description
    Resume SetPBarBackgroundColor_CleanUp

End Sub

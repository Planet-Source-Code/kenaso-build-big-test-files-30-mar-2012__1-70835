VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Build Big Files"
   ClientHeight    =   3105
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   5880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdChoice 
      Caption         =   "&Stop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   4545
      TabIndex        =   4
      Top             =   2385
      Width           =   555
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   90
      ScaleHeight     =   1470
      ScaleWidth      =   5655
      TabIndex        =   8
      Top             =   720
      Width           =   5685
      Begin VB.CommandButton cmdPath 
         Height          =   375
         Left            =   5175
         Picture         =   "frmMain.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   855
         Width           =   420
      End
      Begin VB.OptionButton optOneBigFile 
         Caption         =   "Yes"
         Height          =   240
         Index           =   1
         Left            =   4905
         TabIndex        =   2
         Top             =   270
         Width           =   600
      End
      Begin VB.OptionButton optOneBigFile 
         Caption         =   "No"
         Height          =   240
         Index           =   0
         Left            =   4275
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.TextBox txtTargetPath 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "txtTargetPath"
         Top             =   855
         Width           =   5010
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "txtQty"
         Top             =   225
         Width           =   510
      End
      Begin VB.Label lblFiles 
         BackStyle       =   0  'Transparent
         Caption         =   "Create one big file"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2880
         TabIndex        =   14
         Top             =   270
         Width           =   1320
      End
      Begin VB.Label lblFiles 
         BackStyle       =   0  'Transparent
         Caption         =   "Target path"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label lblFiles 
         Caption         =   "Number of files to create"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Width           =   1950
      End
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   5175
      TabIndex        =   6
      Top             =   2385
      Width           =   555
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   4545
      TabIndex        =   5
      Top             =   2385
      Width           =   555
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2393
      TabIndex        =   12
      Top             =   405
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Build Big Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      TabIndex        =   11
      Top             =   90
      Width           =   2310
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisclaimer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      TabIndex        =   10
      Top             =   2475
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Routine:       frmMain
'
' Description:   Build GOST S-Box table sets and Skipjack F-Table sets.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Sep-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MIN_QTY As Long = 1
  Private Const MAX_QTY As Long = 10

' ***************************************************************************
' Module Variables
'
' Variable name:     mblnOneBigFile
' Naming standard:   m bln OneBigFile
'                    - --- ----------
'                    |  |   |_______ Variable subname
'                    |  |___________ Data type (Boolean)
'                    |______________ Module level designator
'
' ***************************************************************************
  Private mblnOneBigFile As Boolean
  Private mstrTargetPath As String
  Private mlngFileCnt    As Long
  
Private Sub SetControls()

    DoEvents
    With frmMain
        .picFrame.Enabled = False      ' main frame
        .cmdChoice(0).Visible = False  ' Go button
        .cmdChoice(0).Enabled = False
        .cmdChoice(1).Enabled = True   ' Stop button
        .cmdChoice(1).Visible = True
        .cmdChoice(2).Enabled = False  ' Exit button
    End With
                
End Sub

Private Sub ResetControls()

    DoEvents
    With frmMain
        .picFrame.Enabled = True       ' main frame
        .cmdChoice(0).Enabled = True   ' Go button
        .cmdChoice(0).Visible = True
        .cmdChoice(1).Visible = False  ' Stop button
        .cmdChoice(1).Enabled = False
        .cmdChoice(2).Enabled = True   ' Exit button
    End With
    
End Sub

Private Function IsGoodData() As Boolean

    ' Test here instead of Lost_Focus event
    ' so user is not forced to enter data
    ' if attempting to exit application
    
    IsGoodData = False   ' Preset to FALSE
    
    If Not IsNumeric(txtQty.Text) Then
        
        InfoMsg "Quantity must be numeric." & vbNewLine & _
                "Min=" & CStr(MIN_QTY) & "    Max=" & CStr(MAX_QTY)
        txtQty.SetFocus   ' Highlight quantity textbox
    
    ElseIf Val(txtQty.Text) < MIN_QTY Then
        
        InfoMsg "Quantity must be greater than zero." & vbNewLine & _
                "Min=" & CStr(MIN_QTY) & "    Max=" & CStr(MAX_QTY)
        txtQty.SetFocus   ' Highlight quantity textbox
    
    ElseIf Val(txtQty.Text) > MAX_QTY Then
        
        InfoMsg "Quantity exceeds maximum range." & vbNewLine & _
                "Min=" & CStr(MIN_QTY) & "    Max=" & CStr(MAX_QTY)
        txtQty.SetFocus   ' Highlight quantity textbox
    
    ElseIf Len(txtTargetPath.Text) = 0 Then
        
        InfoMsg "Destination path is empty."
        txtTargetPath.SetFocus   ' Highlight path textbox
    
    Else
        
        IsGoodData = True                ' Set flag for good data
        mlngFileCnt = CLng(txtQty.Text)  ' Save number of files to be created
        
    End If
    
End Function

Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index
            
           Case 0  ' GO button
           
                gblnStopProcessing = False
                DoEvents
                
                ' Evaluate textbox data
                If IsGoodData Then
                    SetControls   ' Reset buttons before starting
                    CreateBigFile mstrTargetPath, mlngFileCnt, mblnOneBigFile
                End If
                
                ResetControls  ' Reset buttons when finished
           
           Case 1  ' Stop button
                gblnStopProcessing = True  ' Reset boolean flag
                ResetControls              ' Reset buttons when finished
                DoEvents                   ' Allow time for notification process
                
           Case Else
                DoEvents
                gblnStopProcessing = True  ' Reset boolean flag
                TerminateProgram           ' End this application
    End Select
    
End Sub

Private Sub cmdPath_Click()

    Dim objBrowse As cBrowse
    
    Set objBrowse = New cBrowse   ' Instantiate class module
    mstrTargetPath = objBrowse.BrowseForFolder(frmMain, "Select destination folder")
    Set objBrowse = Nothing       ' Free class object from memory
    
    ' see if a folder was selected
    If Len(Trim$(mstrTargetPath)) > 0 Then
        txtTargetPath.Text = ShrinkToFit(mstrTargetPath, 50)
    Else
        txtTargetPath.Text = vbNullString
    End If
    
End Sub

Private Sub Form_Load()

    mblnOneBigFile = True
    mstrTargetPath = vbNullString
    ResetControls
    
    With frmMain
        .Caption = PGM_NAME & gstrVersion
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                 "No warranties or guarantees implied or intended."
                                 
        .txtQty.Text = 1
        .txtTargetPath.Text = vbNullString
        
        .optOneBigFile(0).Value = True  ' Select No
        optOneBigFile_Click 0           ' Toggle boolean flag to No
        
        ' Center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        TerminateProgram   ' "X" selected in upper right corner
    End If
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail   ' Send email to author of this application
End Sub

Private Sub optOneBigFile_Click(Index As Integer)
    
    mblnOneBigFile = Not mblnOneBigFile  ' Toggle value
    
End Sub

Private Sub txtQty_GotFocus()

    ' Highlight all data in textbox
    With txtQty
         .SelStart = 0             ' start with first char
         .SelLength = Len(.Text)   ' to end of data string
    End With
  
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    ' Evaluate data as it is entered into textbox
    Select Case KeyAscii
           Case 9             ' Tab key
                KeyAscii = 0
                SendKeys "{TAB}"
           Case 13            ' Enter key (no bell sound)
                KeyAscii = 0
           Case 8, 48 To 57   ' backspace & numeric keys only
                ' good data
           Case Else          ' everything else
                KeyAscii = 0
    End Select
                              
End Sub


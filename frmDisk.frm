VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Path Full/Incremental Backup"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmDisk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDiff 
      Caption         =   "Differential"
      Height          =   825
      Left            =   390
      TabIndex        =   15
      Top             =   5040
      Width           =   4155
      Begin VB.CheckBox chkByDateTime 
         Caption         =   "by Date && Time"
         Height          =   285
         Left            =   2220
         TabIndex        =   17
         Top             =   330
         Width           =   1545
      End
      Begin VB.CheckBox chkByDateOnly 
         Caption         =   "by Date only"
         Height          =   285
         Left            =   390
         TabIndex        =   16
         Top             =   330
         Width           =   1425
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Hel&p"
      Height          =   375
      Left            =   5370
      TabIndex        =   14
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtHLQ 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3840
      TabIndex        =   8
      Text            =   "Backup"
      Top             =   6030
      Width           =   2775
   End
   Begin MSComctlLib.ProgressBar pbarDone 
      Height          =   375
      Left            =   345
      TabIndex        =   12
      Top             =   6630
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CheckBox chkSimulate 
      Caption         =   "&Simulate backup and log only.  No file copying."
      Height          =   255
      Left            =   420
      TabIndex        =   6
      Top             =   4620
      Width           =   4215
   End
   Begin VB.CheckBox chkForceFull 
      Caption         =   "&Force full backup of all ""From"" files if checked, else incremental backup will be performed"
      Height          =   405
      Left            =   420
      TabIndex        =   5
      Top             =   4140
      Width           =   4215
   End
   Begin VB.Frame fraTo 
      Caption         =   """To"" Drive"
      Height          =   795
      Left            =   150
      TabIndex        =   3
      Top             =   3270
      Width           =   6615
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   6225
      End
   End
   Begin MSComctlLib.StatusBar sbr 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   8085
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7091
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "5/1/2007"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:49 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5370
      TabIndex        =   10
      Top             =   5300
      Width           =   855
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Backup"
      Height          =   375
      Left            =   5370
      TabIndex        =   9
      Top             =   4300
      Width           =   855
   End
   Begin VB.Frame fraFrom 
      Caption         =   """From"" Path"
      Height          =   2895
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   6615
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   240
         TabIndex        =   2
         Top             =   712
         Width           =   6225
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   6225
      End
   End
   Begin VB.Label lblBytes 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   330
      TabIndex        =   19
      Top             =   7080
      Width           =   6255
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   330
      TabIndex        =   18
      Top             =   7560
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "&High level qualifier (Blank to have a same name, mirror copy of source directory tree.)"
      Height          =   390
      Left            =   420
      TabIndex        =   7
      Top             =   6030
      Width           =   3270
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCurrDir 
      Height          =   390
      Left            =   345
      TabIndex        =   13
      Top             =   6570
      Width           =   6255
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  'This code was based on a source submission on
  'Planet-Source-Code in 2003 named "diskscan"
  'written by Manoz Shrivastava
  'My thanks to the author for posting a very good
  'base for my use in creating this backup program.

  Public NotUsed As String  ' Just to keep the comments here and out of the first code module. (error in VB)
  
  Public gbChkRunning As Boolean
Function DoDateFmt(sInDate As String) As String

' Dim sInput As String
' Dim sPiece As String
 Dim sBuild As String
' Dim i As Long
 
 'Date coming in in the format 6/18/2003 4:57:04 PM
 'Reformat to YYYYMMDDHHMMSS and the HH will be 24 hour
 
 sBuild = Year(sInDate) & Right("0" & Month(sInDate), 2) & Right("0" & Day(sInDate), 2)
 sBuild = sBuild & Right("0" & Hour(sInDate), 2) & Right("0" & Minute(sInDate), 2) & Right("0" & Second(sInDate), 2)
 
 DoDateFmt = sBuild
 
End Function

Sub MakeNewPath(sNewPath As String)

 'Well, it seems that the directory path I need is not there and FSO, stupidly, won't make it for me.
 'Why?  Will never know!  It really, really should.  MkDir won't do it either.  Has to be done one level at a time.
 'OK, Here we go...  slowly...
 
  Dim iCurrBS As Long  ' Where the current backslash I am working with is.
  Dim sCurrPath As String  ' Built up from sNewPath in the error catcher routine.
  
  iCurrBS = 3
  Do While iCurrBS > 0
    iCurrBS = InStr(iCurrBS + 1, sNewPath, "\")
    sCurrPath = Mid$(sNewPath, 1, iCurrBS)
    On Error Resume Next
    If iCurrBS > 0 Then MkDir sCurrPath
    On Error GoTo 0
  Loop
  
End Sub

Function BuildTargetPath(sSourcePath As String) As String

  Dim i As Long
  Dim sBuild As String
  Dim sRebuild As String
  
  i = InStrRev(sSourcePath, "\")
  sBuild = Mid$(sSourcePath, 2, i - 1)
  If gsBackupHiLvlName = "" Then
    sBuild = sTargetDrive & Left(sBuild, 2) & Mid$(sBuild, 3, i - 3)
    i = 4
  Else
    sBuild = sTargetDrive & Left(sBuild, 2) & gsBackupHiLvlName & "\" & Mid$(sBuild, 3, i - 3)
    i = 5 + Len(gsBackupHiLvlName)
  End If
  
  i = InStr(i, sBuild, "\")
  sRebuild = Left(sBuild, i - 1)
  
  If bParmDateDifferential Then
    sRebuild = sRebuild & "_" & Left(DoDateFmt(gsSaveStartTime), 8)
    ElseIf bParmDateTimeDifferential Then sRebuild = sRebuild & "_" & Left(DoDateFmt(gsSaveStartTime), 12)
  End If
  
  BuildTargetPath = sRebuild & Mid$(sBuild, i)
  
End Function
Sub DoBackup()

  Dim i As Long
  Dim sNewPath As String
  Dim sTemp As String
  Dim cStartSeconds As Currency
  Dim cBytesPerSecond As Currency
  Dim cSecsNow As Currency
  Dim cRunTimeSeconds As Currency
  
  pbarDone.Max = CLng(dKiloBytesTotal)
  pbarDone.Visible = True
  lblCurrDir.Visible = False
  lblBytes.Visible = True
  lblSpeed.Visible = True
  lFilesCopied = 0
  cStartSeconds = Timer()
  cBytesPerSecond = 0
  
  sbr.Panels(1).Text = "Copying up to " & glFiles & " files."
 'If bParmSimulateExist Then Me.Caption = "Backup: Copying up to " & glFiles & " files."
  If Me.WindowState = vbMinimized Then
    Me.Caption = "Backup: Copying up to " & glFiles & " files."
  Else
    Me.Caption = gsFullSizeCaption
  End If
  
  For i = 1 To glFiles
    If bExitReq Then
      Close
      Unload Me
      End
    End If
    'Debug.Print sFileModDateTime(i) & " " & sLastRunDateTime
    If sFileModDateTime(i) > sLastRunDateTime Then
      If gbDetailed_Logging Then DebugLog "Candidate " & Right("     " & i, 5) & " " & sFilesList(i)
     'If the force checkbox is not checked and the input parm did not come in copy the file.
      If Not (bParmSimulate) Then
        
        sNewPath = BuildTargetPath(sFilesList(i))  ' Add in all the optional stuff, if any, to get the output path.
        
        sTemp = Dir$(sNewPath, vbDirectory)  ' Is the output path there?
        If sTemp = "" Then MakeNewPath sNewPath  ' If not, go make it up NOW!
        
        If Not bParmForcedFull Then DebugLog "Copying " & sFilesList(i) & " to " & sNewPath
        
        On Error GoTo CantDoCopyfile
        FSO.CopyFile sFilesList(i), sNewPath, True
        On Error GoTo 0
        dKiloBytesDone = dKiloBytesDone + dFileSizes(i)
        lFilesCopied = lFilesCopied + 1
        If Timer() < cStartSeconds Then  ' Have we just gone past midnight?
          cSecsNow = 86400 - cStartSeconds + Timer()  ' Yes, take yesterday's seconds plus todays.
        Else
          cSecsNow = Timer() - cStartSeconds  ' No, just get elapsed seconds today so far.
        End If
        lblSpeed = "Copying " & Format(dKiloBytesDone / cSecsNow, "standard") & " Kilobytes/second."
      End If
    Else
      Debug.Print "Simulated Copy " & sFilesList(i)
    End If
    pbarDone.Value = CLng(dKiloBytesDone)
    sbr.Panels(1).Text = "Of " & glFiles & " source files, " & lFilesCopied & " copied."
    lblBytes = Int(dKiloBytesDone) & " KB of " & Int(dKiloBytesTotal) & " KB copied."
    If i Mod 10 = 0 Then
      If Me.WindowState = vbMinimized Then
        Me.Caption = "Backup: Of " & glFiles & " candidates, " & lFilesCopied & " copied."
      Else
        Me.Caption = gsFullSizeCaption
      End If
      DoEvents
    End If
  Next
  
  DebugLog "Backup complete."
  
Exit Sub

CantDoCopyfile:
  
  DebugLog "Error '" & Err.Description & "' copying " & sFilesList(i), True
  Close
  Unload Me
  End
  
End Sub

Sub DriveProcess()

  Dim i As Integer
  Dim sDateSettingParm As String  ' Concat of input path, output drive and HLQ for hanging backup date on.  All 3 are needed for uniqueness.
  
  DebugLog "Starting."
  bRunning = True
  bExitReq = False  ' Not exiting (Exit button has not been pressed)
  pbarDone.Visible = False

  If chkByDateOnly.Value = vbChecked Then bParmDateDifferential = True
  If chkByDateTime.Value = vbChecked Then bParmDateTimeDifferential = True
  
  If bParmInputPath Then  ' If this is an automatic run
    sSourceFolder = gsStartingSourceDir  ' Use directory name from command$
  Else
    sSourceFolder = Dir1.List(Dir1.ListIndex)  ' Use directory name from the directory listbox
  End If
 'Accept 2, 3, 4, 5 and 6 only
  i = GetDriveType(Left(sSourceFolder, 3))
  Select Case i
    Case 2, 3, 4, 5, 6
     'Nop
    Case Else
      DebugLog "Invalid drive selection for input drive."
      Close
      Unload Me
      End
  End Select
  
' Select Case GetDriveType(drive)
'   Case 2
'     getType = "Removable"
'   Case 3
'     getType = "Drive Fixed"
'   Case 4
'     getType = "Remote"
'   Case 5
'     getType = "Cd-Rom"
'   Case 6
'     getType = "Ram disk"
'   Case Else
'     getType = "Unrecognized"
'  'End Case Else
' End Select
  
  If bParmOutputPath Then  ' If this is an automatic run
    sTargetDrive = gsOutputDriveLetter
  Else
    sTargetDrive = Left(Drive2.Drive, 1)
  End If
 'Accept 2, 3, 4 and 6 only
  i = GetDriveType(Left(sTargetDrive & ":\", 3))
  Select Case i
    Case 2, 3, 4, 6
     'Nop
    Case Else
      DebugLog "Invalid drive selection for output drive."
      Close
      Unload Me
      End
  End Select
  
  gsBackupHiLvlName = Trim$(txtHLQ)  ' Either user typed or input parms might have put something here.
  If sTargetDrive = Left(sSourceFolder, 1) And Trim$(gsBackupHiLvlName) = "" Then
    DebugLog "Cannot backup to same drive without entering a high level qualifier.  Select another drive or enter a high level qualifier and try again.", True
    Exit Sub
  End If
  
  sDateSettingParm = UCase(sSourceFolder & " " & sTargetDrive & " " & gsBackupHiLvlName)
  If chkForceFull.Value = vbChecked Or bParmForcedFull Then  ' Either the box is checked or
    sLastRunDateTime = "01/01/1900 00:00:01 AM"
    DebugLog "Forced full backup copy will be made."
  Else
    sLastRunDateTime = GetSetting(App.EXEName, "LastRunDate", sDateSettingParm, "01/01/1900 00:00:01 AM")
    DebugLog "Updates only run.  Copying all files after " & sLastRunDateTime
  End If
  gsSaveStartTime = Now()
  
  sLastRunDateTime = DoDateFmt(sLastRunDateTime)
    
  fraFrom.Enabled = False
  fraTo.Enabled = False
  fraDiff.Enabled = False
  cmdScan.Enabled = False
  chkForceFull.Enabled = False
  chkSimulate.Enabled = False
  txtHLQ.Enabled = False
  
  glFiles = 0  ' Reset to use first array entry next time.
  DebugLog "Copying from " & sSourceFolder
  DebugLog "Copying to   " & sTargetDrive & ":\" & gsBackupHiLvlName & "\" & Mid$(sSourceFolder, 4)
  
  Set strFolder = FSO.GetFolder(sSourceFolder)
  
  MapDirs (strFolder)
  
  DebugLog "Found " & glFiles & " files in selected path."
  DoEvents
    
  DoBackup
  
 'fraFrom.Enabled = True
 'fraTo.Enabled = True
 'cmdScan.Enabled = True
 'chkForceFull.Enabled = True
 'chkSimulate.Enabled = True
 'txtHLQ.Enabled = True
  
  DebugLog lFilesCopied & " of " & glFiles & " files copied to target."
  sbr.Panels(1).Text = "Copied " & lFilesCopied & " of " & glFiles & " files."
 'If bParmSimulateExist Then Me.Caption = "Backup: Copied " & lFilesCopied & " of " & glFiles & " files."
  If Me.WindowState = vbMinimized Then
    Me.Caption = "Backup: Copied " & lFilesCopied & " of " & glFiles & " files."
  Else
    Me.Caption = gsFullSizeCaption
  End If
  If chkSimulate.Value = vbChecked Then bParmSimulate = True
  
  bRunning = False
  
 'Not that the run is finished, update the cutoff time for the next run.
  If chkSimulate.Value = vbUnchecked And Not (bParmDateTimeDifferential) And Not (bParmDateDifferential) Then SaveSetting App.EXEName, "LastRunDate", sDateSettingParm, gsSaveStartTime
  
End Sub

Private Sub chkByDateOnly_Click()

    
  If gbChkRunning Then Exit Sub
  gbChkRunning = True
  
 'chkByDateOnly.Value = vbChecked
  chkByDateTime.Value = vbUnchecked
  
  gbChkRunning = False
  
End Sub

Private Sub chkByDateTime_Click()

  If gbChkRunning Then Exit Sub
  gbChkRunning = True
  
 'chkByDateTime.Value = vbChecked
  chkByDateOnly.Value = vbUnchecked
  
  gbChkRunning = False

End Sub


Private Sub chkForceFull_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  sbr.Panels(1).Text = "Force complete backup"
End Sub


Private Sub cmdExit_Click()
  
  If Not bRunning Then
    Close
    Unload Me
    End
    Exit Sub
  End If
  
  bExitReq = True

End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  sbr.Panels(1).Text = "Exit"
End Sub

Private Sub cmdHelp_Click()
  frmAbout.Show
End Sub

Private Sub cmdScan_Click()
  DriveProcess
End Sub

Private Sub cmdScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  sbr.Panels(1).Text = "Backup files from 'From' path to 'To' drive."
End Sub

Private Sub Dir1_Change()
 'ChDir Dir1.Path
End Sub

Private Sub Drive1_Change()
  
  On Error Resume Next
  ChDrive Drive1.Drive
  Dir1.Path = Drive1.Drive
  On Error GoTo 0
  
End Sub

Private Sub Form_Load()
      
  Dim sMyParms() As String
  Dim sMyCommand As String
  Dim sTemp As String
  Dim i As Long
  
  gsFullSizeCaption = "Path Full/Incremental Backup vers:" & App.Major & "." & App.Minor & "." & App.Revision
  Me.Caption = gsFullSizeCaption
  
 'Test parms to copy to D disk with HLQ, forced DateTimeDiff backup:    -i C:\My.VB\File Backup -q Backups -o d -t
 'Test parms to copy to D disk with HLQ, forced DateDiff backup:        -i C:\My.VB\File Backup -q Backups -o d -d
 'Test parms to copy to D disk with HLQ, forced full backup, simulated: -i C:\My.VB\File Backup -q Backups -o d -f -s
 'Test parms to copy to C disk with HLQ, forced full backup:            -i C:\My.VB\File Backup -q Backups -o c -f
 'Test parms to copy to D disk NO   HLQ, incremental (error):           -i C:\My.VB\File Backup -o c
 
 'Set up defaults to start with.
  strPath = App.Path
  Set strFolder = FSO.GetFolder(strPath)  ' Just a default will almost certainly be overridden later.

  bRunning = False  ' Not running yet.  Exit works differently depending on running or not.
  gbChkRunning = False
  
  pbarDone.Value = pbarDone.Max
  dKiloBytesTotal = 0
  dKiloBytesDone = 0
  
 'Input parms description
 '  -i is input path for top level to copy.  All lower levels will be copied.
 '  -q is the high level qualifier to use to put the files under.  If missing, "Backup" is used.
 '  -o is the drive to place the backup tree structure on.
 '  -f (switch -- no parms required) sets a forced full backup of the entire tree.
 '  -s (switch -- no parms required) only simulate the copy and log it, do not copy any files.
 '  -d (switch -- no parms required) Do date differential (-f will cause a dated full)
 '  -t (switch -- no parms required) Do Date & Time differential (-f will cause a dated full)
 
 'Assume no parms are there.
  bParmInputPath = False
  bParmHighLvlQual = False
  bParmDateDifferential = False
  bParmForcedFull = False
  bParmSimulate = False
  bParmDateDifferential = False
  bParmDateTimeDifferential = False
  
  sMyCommand = Command$
  sMyCommand = Replace(sMyCommand, sQuote, "")
  If Trim$(sMyCommand) <> "" Then DebugLog "Input parms to program: " & sMyCommand
  sMyParms = Split(sMyCommand, "-")
  bErrorInParms = False  ' assume no errors

  If UBound(sMyParms) > 0 Then
    Me.Show
    Me.WindowState = vbMinimized
    
    bParmSimulateExist = True
    
    For i = 0 To UBound(sMyParms)
      DebugLog i & " " & sMyParms(i)
      Select Case UCase(Left(sMyParms(i), 1))
  '    -i is input path for top level to copy.  All lower levels will be copied.
        Case "I"
          bParmInputPath = True
          gsStartingSourceDir = Trim$(Mid$(sMyParms(i), 3))
          Drive1.Drive = gsStartingSourceDir
          Dir1.Path = gsStartingSourceDir
          
  '    -q is the high level qualifier to use to put the files under.  If missing, "Backup" is used.
        Case "Q"
          bParmHighLvlQual = True
          txtHLQ = Trim$(Mid$(sMyParms(i), 3))  ' This is where it is picked up later and put into the global
       
  '    -o is the output drive to place the backup tree structure on. Can be the same drive as I if Q is not blank.
        Case "O"
          bParmOutputPath = True
          gsOutputDriveLetter = UCase(Mid$(sMyParms(i), 3, 1))
          Drive2.Drive = gsOutputDriveLetter
          
  '    -f (switch -- no parms required) sets a forced full backup of the entire tree.
        Case "F"
          bParmForcedFull = True
          chkForceFull.Value = vbChecked
          
  '    -s (switch -- no parms required) only simulate the copy and log it, do not copy any files.
        Case "S"
          bParmSimulate = True
          chkSimulate.Value = vbChecked
          
  '    -d (switch -- no parms required) force a separate 'date differential' path on each run based on date.
        Case "D"
          bParmDateDifferential = True
          chkByDateOnly.Value = vbChecked
          
  '    -t (switch -- no parms required) force a separate 'date_time differential' path on each run based on date.
        Case "T"
          bParmDateTimeDifferential = True
          chkByDateTime.Value = vbChecked
          
      End Select
    Next
    DoEvents
    DebugLog "Input parms found.  Run mode: Automatic in the background."
    DebugLog "Input Parm -- Source Directory: " & gsStartingSourceDir
    DebugLog "Input Parm -- Output drive letter: " & gsOutputDriveLetter
    DebugLog "Input Parm -- High Level Output Qualifier: " & txtHLQ
    If bParmForcedFull Then
      DebugLog "Input Parm -- This will be a forced full backup."
    Else
      DebugLog "Input Parm -- This will be a partial back of changed files since last run."
    End If
    If bParmSimulate Then
      DebugLog "Input Parm -- This will be a simulation run, no data copying.  Run date not udpated."
    Else
      DebugLog "Input Parm -- This will be a live run with data copying and run date will be updated."
    End If
    sTemp = Dir$(gsStartingSourceDir, vbDirectory)
    If sTemp = "" Then
      DebugLog "ERROR - The input source directory '" & gsStartingSourceDir & "' does not exist.  Cannot continue in automatic mode."
      Close
      Unload Me
      End
    End If
    If bParmSimulateExist And Not bParmInputPath Then bErrorInParms = True
    If bParmSimulateExist And Not bParmDateDifferential Then bErrorInParms = True
   
   'Be sure the two drive letters are not the same unless HLQ is in use.
    If Left(gsStartingSourceDir, 1) = gsOutputDriveLetter And Trim$(gsBackupHiLvlName) = "" Then
      DebugLog "Cannot use same drive for source  and destination."
      Close
      Unload Me
      End
    End If
    
   'Now, the parms are set into the proper variables, run the two routines and exit
    DriveProcess
    Close
    Unload Me
    End
  Else
    bParmSimulateExist = False
    DebugLog "No input parms found.  Run mode: Manual."
  End If
  
End Sub
Private Function MapDirs(sFolder As String)

  Dim pFolder As Folder
  
  If bExitReq Then
    Close
    Unload Me
    End
  End If
  
  On Error Resume Next
  
  Set pFolder = FSO.GetFolder(sFolder)
  
  For Each strSubFolder In pFolder.SubFolders
    'txtRpt = txtRpt & vbNewLine & UCase$(strSubFolder.Path) & vbNewLine & vbNewLine
    'Debug.Print "2-" & UCase$(strSubFolder.Path)
    sCurDir = strSubFolder.Path
    If gbDetailed_Logging Then DebugLog sCurDir
    lblCurrDir = sCurDir
    DoEvents
    MapDirs (strSubFolder.Path)
  Next
  
  For Each strFile In pFolder.Files
    If Left(strFile.Name, 1) <> "~" Then  ' Skil temp files left hanging about by Word and, possibly, others.
      glFiles = glFiles + 1
      If glFiles > MAX_FILES Then
        DebugLog "More than " & MAX_FILES & " files in the path selected.  Program cannot continue.", True
        Close
        Unload Me
        End
      End If
      sFilesList(glFiles) = strFile
      sFileModDateTime(glFiles) = DoDateFmt(strFile.DateLastModified)
      dFileSizes(glFiles) = (strFile.Size / 1000)
      dKiloBytesTotal = dKiloBytesTotal + dFileSizes(glFiles)
      If glFiles Mod 20 = 0 Then
        If bExitReq Then
          Close
          Unload Me
          End
        End If
        sbr.Panels(1).Text = "Found " & glFiles & " files."
        If bParmSimulateExist Or Me.WindowState = vbMinimized Then
          Me.Caption = "Backup: Found " & glFiles & " files."
        Else
          Me.Caption = gsFullSizeCaption
        End If
        DoEvents
      End If
    End If
  Next

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  DebugLog "Exiting"
  DebugLog "-------"
  
End Sub

Private Sub fraFrom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  sbr.Panels(1).Text = "Select backup source directory"
End Sub


Private Sub fraTo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  sbr.Panels(1).Text = "Select backup target directory"
End Sub


Private Sub Label3_Click()

End Sub

Private Sub txtHLQ_GotFocus()

  txtHLQ.SelStart = 0
  txtHLQ.SelLength = Len(txtHLQ.Text)
  
End Sub



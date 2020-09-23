VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   7410
   ClientLeft      =   2250
   ClientTop       =   1935
   ClientWidth     =   13215
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5114.514
   ScaleMode       =   0  'User
   ScaleWidth      =   12409.57
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrev 
      Cancel          =   -1  'True
      Caption         =   "< &Back"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   3049
      TabIndex        =   7
      Top             =   6737
      Width           =   1260
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   345
      Left            =   4646
      TabIndex        =   6
      Top             =   6720
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   8906
      TabIndex        =   0
      Top             =   6737
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   7308
      TabIndex        =   2
      Top             =   6737
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   12171.99
      Y1              =   4451.905
      Y2              =   4451.905
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5220
      Left            =   150
      TabIndex        =   3
      Top             =   1125
      Width           =   12945
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1044
      TabIndex        =   4
      Top             =   240
      Width           =   3546
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   84.515
      X2              =   12171.99
      Y1              =   4472.611
      Y2              =   4472.611
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   195
      Left            =   1044
      TabIndex        =   5
      Top             =   600
      Width           =   3546
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private iWhichShowing As Integer
Private Const MAX_SHOWING = 3
Sub ShowLabel()
  
  Select Case iWhichShowing
  
  Case 1
    lblDescription = "Operational overview" & vbCrLf & _
                      vbCrLf & _
                      "The default operation of this backup is to create a single output path and update that single" & vbCrLf & _
                      "copywith each run.  You can force a full backup of the source path on the first run and then" & vbCrLf & _
                      "update theoutput path with only the changed files on subsequent runs.  This is the simpliest" & vbCrLf & _
                      "execution and is a way to keep an up-to-date copy of a path on a separate hard drive.  This" & vbCrLf & _
                      "hard drive can be local, USB or NAS." & vbCrLf & _
                      vbCrLf & _
                      "The other operational mode is differential copy.  In this mode, the date or date and time is" & vbCrLf & _
                      "insered into the output path.  This mode allows the user to maintain multiple versions of" & vbCrLf & _
                      "frequently changed files. Each backuped up version will be in a separate output path by date or" & vbCrLf & _
                      "date and time.  Frequency of runswill determine how many versions of changed files will be" & vbCrLf & _
                      "saved." & vbCrLf & _
                      vbCrLf & _
                      "Press 'Next' and 'Back' for additional information."

  Case 2
    lblDescription = "There are seven input parms to this program.  If 'i', 'q' or 'o' are input, then all three of them must be input.  " & _
                     "'f', 'd', 't' and 's' are idependent and can be used as desired with 'i', 'q' and 'o'." & vbCrLf & _
                     vbCrLf & _
                     "Here are the parms:" & vbCrLf & _
                     "-i   is full path for top level to copy.  All lower levels will be copied." & vbCrLf & _
                     "-q   is the high level qualifier to use to put the files under.  If missing, 'Backup' is used." & vbCrLf & _
                     "-o   is the letter of the output drive on which to place the backup tree structure." & vbCrLf & _
                     "      Do not include a direcory name, just a single letter." & vbCrLf & _
                     "-f   (switch -- no parms required) sets a forced full backup of the entire tree." & vbCrLf & _
                     "-s   (switch -- no parms required) only simulate the copy and log it, do not copy any files." & vbCrLf & _
                     "-d   (switch -- no parms required) sets the mode to date differential." & vbCrLf & _
                     "      A new output path will be created for runs each day with only the files changed since last" & vbCrLf & _
                     "      run." & vbCrLf & _
                     "-d   (switch -- no parms required) sets the mode to date differential." & vbCrLf & _
                     "      A new folder will be created daily with only the files changed since last day/run." & vbCrLf & _
                     "-t   (switch -- no parms required) sets the mode to date/time differential." & vbCrLf & _
                     "      A new folder will be created for each run with only the files changed since last run."
  Case 3
    lblDescription = "The parameters IDs (-i, -q, -o, -f, -d && -s) can be either upper or lower case." & vbCrLf & _
                     "The parameters are case sensitive.  If you want really 'frEMbish' for high qualifier level, type it exactly that way." & vbCrLf & _
                     vbCrLf & _
                     "Note: Path && directory names MAY have spaces and all other valid Windows path/file name characters."
    
    End Select
  
End Sub

Private Sub cmdNext_Click()

  cmdPrev.Enabled = True
  
  If iWhichShowing < MAX_SHOWING Then iWhichShowing = iWhichShowing + 1
  If iWhichShowing = MAX_SHOWING Then cmdNext.Enabled = False
  
  ShowLabel

End Sub


Private Sub cmdPrev_Click()

  cmdNext.Enabled = True
  
  If iWhichShowing > 1 Then iWhichShowing = iWhichShowing - 1
  If iWhichShowing = 1 Then cmdPrev.Enabled = False
  
  ShowLabel
  
  
End Sub


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = "About " & App.Title
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  lblTitle.Caption = App.Title
  
  iWhichShowing = 1
  ShowLabel
                     
End Sub
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


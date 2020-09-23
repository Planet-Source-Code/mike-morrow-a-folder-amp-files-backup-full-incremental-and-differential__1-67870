Attribute VB_Name = "modGlobal"
Option Explicit

  Public Const sQuote = """"
  Public Const gbDetailed_Logging = False
  Public gsFullSizeCaption As String
  
  Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
  
  Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
  
 'This code was based on a source submission on
 'Planet-Source-Code in 2003 named "diskscan"
 'written by Manoz Shrivastava
 'My thanks to the author for posting a very good
 'base for my use in creating this backup program.

  Public strReportFile As String, intSlashPos As Integer
  Public intCnt As Integer
  Public FSO As New FileSystemObject
  Public strFolder As Folder, strSubFolder As Folder, strFile As File
  Public sSourceFolder As String
  Public strPath As String, strDsc As Boolean
  Public intTotFiles As Integer, intTotDirs As Integer
  Public intCurDepth As Integer, strInitDir As String
  
  Public sCurDir As String
  Public Const MAX_FILES = 25000
  Public glFiles As Long  ' Count of files found
  Public sFilesList(1 To MAX_FILES) As String  ' Make a list of all file paths here
  Public dFileSizes(1 To MAX_FILES) As Double  ' Filesize of each file / 1000
  Public dKiloBytesTotal As Double  ' Total of all files / 1000
  Public dKiloBytesDone As Double  ' Total of file bytes done / 1000
  Public sFileModDateTime(1 To MAX_FILES) As String  ' Date last modified here
  Public sLastRunDateTime As String  ' Last time this process has been run.  First run it is "19000101000000"
    
  Public sTargetDrive As String

  Public lFilesCopied As Long

  Public bExitReq As Boolean

  Public bRunning As Boolean

  Public gsBackupHiLvlName As String
  Public gsStartingSourceDir As String
  Public gsOutputDriveLetter As String
  Public gsSaveStartTime As String
  
  Public bParmSimulateExist As Boolean  ' True if any parms found and are valid -- Means automatic running.
  Public bParmInputPath As Boolean  ' True if the "i" parm exists and is valid
  Public bParmHighLvlQual As Boolean  ' True if HLQ entered in parms
  Public bParmForcedFull As Boolean
  Public bParmSimulate As Boolean
  Public bParmOutputPath As Boolean
  Public bParmDateDifferential As Boolean  ' Do Date Differential
  Public bParmDateTimeDifferential As Boolean  ' Do Date & Time Differential
  Public bErrorInParms As Boolean
  
Sub DebugLog(sIn As String, Optional DoMB As Boolean = False)

  Dim iFNo As Integer
  
  iFNo = FreeFile()
  
  Open App.Path & "\RunLog.txt" For Append Access Write As iFNo
  Print #iFNo, Now() & " " & sIn
  Close iFNo
  
  Debug.Print Now() & " " & sIn
  
  If DoMB And Not bParmSimulateExist Then MsgBox (sIn)  ' If parms exist, probably a batch (console) initiation.
  
End Sub
Public Function Enumerate_Drives() As String
  
  Dim result As String
  
  result = String(255, Chr$(0))
  GetLogicalDriveStrings 255, result
  Enumerate_Drives = result

End Function


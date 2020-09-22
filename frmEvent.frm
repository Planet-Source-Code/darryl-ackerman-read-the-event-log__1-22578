VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEvent 
   Caption         =   "frmEvent"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6855
      TabIndex        =   1
      Top             =   3255
      Width           =   6915
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open Log"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton optLog 
         Caption         =   "Application"
         Height          =   465
         Index           =   0
         Left            =   540
         TabIndex        =   4
         Top             =   90
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optLog 
         Caption         =   "System"
         Height          =   465
         Index           =   1
         Left            =   1890
         TabIndex        =   3
         Top             =   90
         Width           =   1185
      End
      Begin VB.OptionButton optLog 
         Caption         =   "Security"
         Height          =   465
         Index           =   2
         Left            =   3150
         TabIndex        =   2
         Top             =   90
         Width           =   1185
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2610
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvent.frx":0000
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvent.frx":0452
            Key             =   "Information"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvent.frx":08A4
            Key             =   "Error"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvent.frx":0CF6
            Key             =   "Success Audit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvent.frx":0E50
            Key             =   "Failed Audit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstLog 
      Height          =   3165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   5583
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Event"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Computer"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I've seen many futile attempts to read the Windows NT event log.
'Although this code is not complete it does however show the
'correct way to get the total number of event log records for
'the givin log file.
'I'm currently working on completeing the code, however I thought
'since it was so hard for me to find help that someone out there
'may be struggling with the same issues I was. So here's the code
'to open the event log and read all the records without missing records
'or erroring out due to incorrect constants.
'Do what you wish with the code just remember to tell anyone who asks you for the code
'who the real author is...
'keep looking for updates as i'll be posting them as i continue to develop this program
'**************************
'Copyright 2001 Darryl Ackerman darryl@optonline.net
'**************************

Option Explicit

Public intOptionButton As Integer
Public strComputer As String

Private Const EVENTLOG_SUCCESS = 0
Private Const EVENTLOG_ERROR_TYPE = 1
Private Const EVENTLOG_WARNING_TYPE = 2
Private Const EVENTLOG_INFORMATION_TYPE = 4
Private Const EVENTLOG_AUDIT_SUCCESS = 8
Private Const EVENTLOG_AUDIT_FAILURE = 10
Private Const EVENTLOG_SEQUENTIAL_READ = &H1
Private Const EVENTLOG_SEEK_READ = &H2
Private Const EVENTLOG_FORWARDS_READ = &H4
Private Const EVENTLOG_BACKWARDS_READ = &H8
Private Const ERROR_HANDLE_EOF = 38
Private Const BUFFER_SIZE As Long = 256
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122

'
Private Type EVENTLOGRECORD
     Length As Long     '  Length of full record
     Reserved As Long     '  Used by the service
     RecordNumber As Long     '  Absolute record number
     TimeGenerated As Long     '  Seconds since 1-1-1970
     TimeWritten As Long     'Seconds since 1-1-1970
     EventID As Long
     EventType As Integer
     NumStrings As Integer
     EventCategory As Integer
     ReservedFlags As Integer     '  For use with paired events (auditing)
     ClosingRecordNumber As Long     'For use with paired events (auditing)
     StringOffset As Long     '  Offset from beginning of record
     UserSidLength As Long
     UserSidOffset As Long
     DataLength As Long
     DataOffset As Long     '  Offset from beginning of record
End Type

Private Declare Function OpenEventLog Lib "advapi32.dll" Alias "OpenEventLogA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function CloseEventLog Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Private Declare Function ReadEventLog Lib "advapi32.dll" Alias "ReadEventLogA" (ByVal hEventLog As Long, ByVal dwReadFlags As Long, ByVal dwRecordOffset As Long, lpBuffer As EVENTLOGRECORD, ByVal nNumberOfBytesToRead As Long, pnBytesRead As Long, pnMinNumberOfBytesNeeded As Long) As Long
Private Declare Function GetNumberOfEventLogRecords Lib "advapi32.dll" (ByVal hEventLog As Long, NumberOfRecords As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long



Private Function OpenLog(ByVal EventLog As String, Optional ByVal ComputerName As String)
Dim lngHandle As Long, lngRet As Long, i As Long, lngTotRec As Long
Dim strError As String, lngReadFlags As Long, lngBytesToRead As Long
Dim lngBytesRead As Long, lngMinBytesNeeded As Long
Dim itmX() As ListItem, strType As String
Dim lngBuffer As Long, EVT_REC() As EVENTLOGRECORD

On Error GoTo GetEvtErr

If ComputerName = "" Then ComputerName = vbNullString

lngReadFlags = EVENTLOG_FORWARDS_READ Or EVENTLOG_SEQUENTIAL_READ

lngTotRec = GetNumRecords(EventLog, ComputerName)

If lngTotRec = -1 Then GoTo GetEvtErr

ReDim itmX(1 To lngTotRec)

lngHandle = OpenEventLog(ComputerName, EventLog)
If lngHandle = 0 Then GoTo GetEvtErr

ReDim EVT_REC(0)
i = 0
lngBuffer = 0

Do
    'pass a zero-length buffer to get the lngMinBytesNeeded variable to return the minimum bytes needed for the buffer...
    lngRet = ReadEventLog(lngHandle, lngReadFlags, 0, EVT_REC(0), lngBuffer, lngBytesRead, lngMinBytesNeeded)
    
    If lngRet = 0 Then
        Select Case GetLastError()
            
           Case ERROR_INSUFFICIENT_BUFFER
                'Since the function returns 0 (failed) we'll re-initialize all the variables now that
                'we know how large to make the lngBuffer variable and the EVT_REC() EVENTLOGRECORD structure
                lngBuffer = lngMinBytesNeeded
                ReDim EVT_REC(0 To lngMinBytesNeeded)
                
                'now call the function again to get the record...
                lngRet = ReadEventLog(lngHandle, lngReadFlags, 0, EVT_REC(0), lngBuffer, lngBytesRead, lngMinBytesNeeded)
            
            Case Else
                GoTo GetEvtErr
        End Select
    End If

    'figure out what type of record it is...
    Select Case EVT_REC(0).EventType
        Case EVENTLOG_SUCCESS
            strType = "Success"
        Case EVENTLOG_ERROR_TYPE
           strType = "Error"
        Case EVENTLOG_WARNING_TYPE
           strType = "Warning"
        Case EVENTLOG_INFORMATION_TYPE
            strType = "Information"
        Case EVENTLOG_AUDIT_SUCCESS
            strType = "Success Audit"
        Case EVENTLOG_AUDIT_FAILURE
            strType = "Failed Audit"
    End Select
    
    'add the record to the list box...
    i = i + 1
    
    Set itmX(i) = lstLog.ListItems.Add()

    itmX(i).Text = strType
    itmX(i).SmallIcon = strType
    itmX(i).SubItems(1) = EVT_REC(0).TimeGenerated     'date
    itmX(i).SubItems(2) = EVT_REC(0).TimeWritten           'time
    itmX(i).SubItems(3) = EVT_REC(0).StringOffset           'source
    'category
    If EVT_REC(0).EventCategory = 0 Then itmX(i).SubItems(4) = "None" Else itmX(i).SubItems(4) = EVT_REC(0).EventCategory
    itmX(i).SubItems(5) = EVT_REC(0).EventID                  'event
    itmX(i).SubItems(6) = "User" 'EVT_REC(0)                                'user
    itmX(i).SubItems(7) = "computer" 'EVT_REC(0)                                'computer
    
'reset the buffer back to zero so we can figure out the size of the next record...
lngBuffer = 0
Loop While lngBytesRead > 0 And i < lngTotRec

Done:
lngRet = CloseEventLog(lngHandle)
OpenLog = 1

Exit Function
GetEvtErr:
lngRet = CloseEventLog(lngHandle)
MsgBox "An error has occured... GetLastError():" & GetLastError()
OpenLog = 0
End Function

Public Function GetNumRecords(ByVal EvtLog As String, Optional ByVal PCName As String) As Long
Dim lngHandle As Long
Dim lngCount As Long
Dim lngRet As Long
Dim strError As String

If PCName = "" Then PCName = vbNullString

lngHandle = OpenEventLog(PCName, EvtLog)
If lngHandle = 0 Then GoTo GetErr

lngRet = GetNumberOfEventLogRecords(lngHandle, lngCount)
If lngRet = 0 Then GoTo GetErr

lngRet = CloseEventLog(lngHandle)
If lngRet = 0 Then GoTo GetErr

GetNumRecords = lngCount

Exit Function
GetErr:
MsgBox "An error was returned byt the GetNumRecords() function." & vbCrLf & "Error: " & GetLastError()
GetNumRecords = -1
End Function

Private Function GetLocalPCName() As String
Dim lngRet As Long, strBuff As String, lngSize As Long

lngSize = 255
strBuff = String(lngSize, Chr(0))

lngRet = GetComputerName(strBuff, lngSize)

GetLocalPCName = Left(strBuff, InStr(1, strBuff, Chr(0)) - 1)


End Function

Private Sub cmdOpen_Click()
Dim strLog As String, lngRet As Long

lstLog.ListItems.Clear

strLog = optLog(intOptionButton).Caption

lngRet = OpenLog(strLog, strComputer)

If lngRet = 0 Then MsgBox "There was an error reading the event log." & vbCrLf & "Error: " & GetLastError(): Exit Sub

Caption = "Computer: [" & strComputer & "] Log: [" & strLog & "]"

End Sub

Private Sub Form_Load()
strComputer = GetLocalPCName()
Caption = "Computer: [" & strComputer & "]"
End Sub

Private Sub Form_Resize()
lstLog.Move 0, 0, ScaleWidth, ScaleHeight - Picture1.Height
End Sub

Private Sub optLog_Click(Index As Integer)
intOptionButton = Index
End Sub

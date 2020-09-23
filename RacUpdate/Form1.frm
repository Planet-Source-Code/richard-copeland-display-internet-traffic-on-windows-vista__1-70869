VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Connections"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ListView lsvListView2 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lsvListView1 
      Height          =   5415
      Left            =   50
      TabIndex        =   1
      Top             =   50
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9551
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type MIB_TCPROW_OWNER_PID
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
    dwOwningPid As Long
End Type



Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
                                                        ByVal bInheritHandle As Long, _
                                                        ByVal dwProcId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByVal hModule As Long, _
                                                        ByVal ModuleName As String, _
                                                        ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, _
                                                        ByRef lphModule As Long, _
                                                        ByVal cb As Long, _
                                                        ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const MAX_PATH = 256
Private Const AF_INET6 = 23
Private Const AF_INET = 2

Public Enum TCP_TABLE_CLASS
  TCP_TABLE_BASIC_LISTENER
  TCP_TABLE_BASIC_CONNECTIONS
  TCP_TABLE_BASIC_ALL
  TCP_TABLE_OWNER_PID_LISTENER
  TCP_TABLE_OWNER_PID_CONNECTIONS
  TCP_TABLE_OWNER_PID_ALL
  TCP_TABLE_OWNER_MODULE_LISTENER
  TCP_TABLE_OWNER_MODULE_CONNECTIONS
  TCP_TABLE_OWNER_MODULE_ALL
End Enum

Private Declare Function htons Lib "ws2_32.dll" (ByVal dwLong As Long) As Long

Private Declare Function GetExtendedTcpTable Lib "iphlpapi.dll" (pTcpTableEx As Any, lsize As Long, ByVal bOrder As Long, ByVal flags As Long, ByVal TableClass As TCP_TABLE_CLASS, ByVal bReserved As Long) As Long



Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private pTablePtr() As Byte
Public nRows As Long
Private pDataRef As Long

Public Function GetRefresh() As Boolean '??TcpTable,??RefreshStack
Dim lngSize As Long, nRet As Long
    lngSize = 4
    nRet = GetExtendedTcpTable(0&, lngSize, 1, AF_INET, TCP_TABLE_OWNER_PID_ALL, 0)  'Requires Windows Vista or Windows XP SP2.
    ReDim pTablePtr(lngSize - 1)

    nRet = GetExtendedTcpTable(pTablePtr(0), lngSize, 1, AF_INET, TCP_TABLE_OWNER_PID_ALL, 0)

'    nRet = GetTcpTable(0&, lngSize, 0)
'    ReDim pTablePtr(lngSize - 1)
'    If nRet <> 0 Then nRet = GetTcpTable(pTablePtr(0), lngSize, 0)

    If nRet = 0 Then
        CopyMemory nRows, pTablePtr(0), 4
    Else
        GetRefresh = False
        Exit Function
    End If
    
    If nRows = 0 Or pTablePtr(0) Then
    GetRefresh = False
    Exit Function
    End If



End Function

Public Sub DoProcess()
GetRefresh
RefreshStack
RefreshAndCompareLists
End Sub
Public Sub RefreshStack()

Dim pPath As String
Dim pName As String
Dim str As String
Dim i As Long
Dim tcpTable As MIB_TCPROW_OWNER_PID
    pDataRef = 0
Dim lstLine As ListItem
On Error Resume Next

lsvListView2.ListItems.Clear 'clears listbox so only current/new items are present

For i = 0 To nRows ' read 24 bytes at a time

    CopyMemory tcpTable, pTablePtr(0 + pDataRef + 4), LenB(tcpTable)

        If tcpTable.dwRemoteAddr <> 0 Or GetPort(tcpTable.dwRemotePort) <> 0 Or GetPort(tcpTable.dwLocalPort) <> 0 Then

Set lstLine = lsvListView2.ListItems.Add(1, vbNullString, c_state(tcpTable.dwState)) ' Blank
'lstLine.SubItems(0) = c_state(tcpTable.dwState) 'not needed
str = GetIPAddress(tcpTable.dwLocalAddr)
lstLine.SubItems(1) = str
lstLine.SubItems(2) = GetPort(tcpTable.dwLocalPort)
str = GetIPAddress(tcpTable.dwRemoteAddr)
lstLine.SubItems(3) = str
lstLine.SubItems(4) = GetPort(tcpTable.dwRemotePort)
lstLine.SubItems(5) = tcpTable.dwOwningPid
pPath = GetPath(getPidPathName(tcpTable.dwOwningPid)) ' Get just directory
If pPath = "[" Then pPath = ""
If pPath = "" Then pPath = " "
lstLine.SubItems(6) = pPath
pName = GetPathFileN(getPidPathName(tcpTable.dwOwningPid)) ' Get just the filename
lstLine.SubItems(7) = pName
        End If
        pDataRef = pDataRef + LenB(tcpTable)
        DoEvents
Next i

'RefreshAndCompareLists 'see sub for details
End Sub
Public Function GetPort(ByVal dwPort As Long) As Long
On Error Resume Next
'???????,?htons?????long?
    GetPort = htons(dwPort)
End Function
Public Function GetIPAddress(dwAddr As Long) As String
    Dim arrIpParts(3) As Byte
    On Error Resume Next
    CopyMemory arrIpParts(0), dwAddr, 4
    GetIPAddress = CStr(arrIpParts(0)) & "." & _
    CStr(arrIpParts(1)) & "." & _
    CStr(arrIpParts(2)) & "." & _
    CStr(arrIpParts(3))
End Function
Function c_state(s) As String
  On Error Resume Next
  Select Case s
  Case "0": c_state = "UNKNOWN"
  Case "1": c_state = "CLOSED"
  Case "2": c_state = "LISTENING"
  Case "3": c_state = "SYN_SENT"
  Case "4": c_state = "SYN_RCVD"
  Case "5": c_state = "ESTABLISHED"
  Case "6": c_state = "FIN_WAIT1"
  Case "7": c_state = "FIN_WAIT2"
  Case "8": c_state = "CLOSE_WAIT"
  Case "9": c_state = "CLOSING"
  Case "10": c_state = "LAST_ACK"
  Case "11": c_state = "TIME_WAIT"
  Case "12": c_state = "DELETE_TCB"
  End Select
End Function

Public Function getPidPathName(pid As Long) As String
On Error Resume Next
Dim cbNeeded As Long
Dim Modules(1 To 2000) As Long
Dim nSize As Long
Dim lRet As Long
Dim ModuleName As String
Dim hProcess As Long
If pid = 0 Then getPidPathName = "[System Idle Process]": Exit Function
If pid = 4 Then getPidPathName = "[System]": Exit Function

 hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pid)
 If hProcess <> 0 Then
     lRet = EnumProcessModules(hProcess, Modules(1), 2000, cbNeeded)
    If lRet <> 0 Then
        ModuleName = Space(MAX_PATH)
        nSize = MAX_PATH
        lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
        If CBool(InStr(1, (Left(ModuleName, lRet)), "", vbTextCompare)) Then '''ModuleName??????,????????“”,????????????
            getPidPathName = Left(ModuleName, lRet)
        End If
    End If
End If
lRet = CloseHandle(hProcess)
End Function
''?form?





Public Function ListItemWholeString(xInt As Integer, xLst As ListView)
If xLst.ListItems.Count >= xInt Then ListItemWholeString = xLst.ListItems(xInt).Text & "-" & xLst.ListItems(xInt).SubItems(1) & "-" & xLst.ListItems(xInt).SubItems(2) & "-" & xLst.ListItems(xInt).SubItems(3) & "-" & xLst.ListItems(xInt).SubItems(4) & "-" & xLst.ListItems(xInt).SubItems(5) & "-" & xLst.ListItems(xInt).SubItems(6) & "-" & xLst.ListItems(xInt).SubItems(7)
End Function
Public Sub RefreshAndCompareLists()
On Error Resume Next
Dim xItem1 As Integer, xItem2 As Integer
Dim CurItem1 As String, CurItem2 As String
Dim IsThere1 As Boolean, IsThere2 As Boolean
Dim xNewItem As ListItem
'lsvListView2 - has the real/current connections (just refreshed)
'lsvListView1 - is the display for the connections (update each item only if needed)

For xItem2 = 1 To lsvListView2.ListItems.Count
CurItem2 = ListItemWholeString(xItem2, lsvListView2)
 IsThere1 = False
 For xItem1 = 1 To lsvListView1.ListItems.Count
  CurItem1 = ListItemWholeString(xItem1, lsvListView1)
  If CurItem2 = CurItem1 Then IsThere1 = True
 DoEvents
 Next xItem1

If IsThere1 = False Then
Set xNewItem = lsvListView1.ListItems.Add(1, vbNullString, lsvListView2.ListItems(xItem2).Text)
xNewItem.SubItems(1) = lsvListView2.ListItems(xItem2).SubItems(1)
xNewItem.SubItems(2) = lsvListView2.ListItems(xItem2).SubItems(2)
xNewItem.SubItems(3) = lsvListView2.ListItems(xItem2).SubItems(3)
xNewItem.SubItems(4) = lsvListView2.ListItems(xItem2).SubItems(4)
xNewItem.SubItems(5) = lsvListView2.ListItems(xItem2).SubItems(5)
xNewItem.SubItems(6) = lsvListView2.ListItems(xItem2).SubItems(6)
xNewItem.SubItems(7) = lsvListView2.ListItems(xItem2).SubItems(7)
End If

DoEvents
Next xItem2



For xItem1 = 1 To lsvListView1.ListItems.Count + 1
CurItem1 = ListItemWholeString(xItem1, lsvListView1)
IsThere2 = False
 For xItem2 = 1 To lsvListView2.ListItems.Count + 1
  CurItem2 = ListItemWholeString(xItem2, lsvListView2)
  If CurItem2 = CurItem1 Then IsThere2 = True
 DoEvents
 Next xItem2

If IsThere2 = False Then
lsvListView1.ListItems.Remove xItem1
End If

DoEvents
Next xItem1

End Sub

Private Sub Form_Load()
Form2.Visible = True
 With lsvListView2
        .View = lvwReport
        .ColumnHeaders.Add 1, , "State"
        .ColumnHeaders.Add 2, , "Local IP Address"
        .ColumnHeaders.Add 3, , "Local Port"
        .ColumnHeaders.Add 4, , "Remote IP address"
        .ColumnHeaders.Add 5, , "Remote Port "
        .ColumnHeaders.Add 6, , "Process Id "
        .ColumnHeaders.Add 7, , "Process Name with Path"
        .ColumnHeaders.Add 8, , "Process Name"
End With
 With lsvListView1
        .View = lvwReport
        .ColumnHeaders.Add 1, , "State"
        .ColumnHeaders.Add 2, , "Local IP Address"
        .ColumnHeaders.Add 3, , "Local Port"
        .ColumnHeaders.Add 4, , "Remote IP address"
        .ColumnHeaders.Add 5, , "Remote Port "
        .ColumnHeaders.Add 6, , "Process Id "
        .ColumnHeaders.Add 7, , "Process Name with Path"
        .ColumnHeaders.Add 8, , "Process Name"
End With
End Sub





Private Sub lsvListView2_Click()
Label2 = lsvListView2.SelectedItem.Index
'MsgBox ListItemWholeString(lsvListView2.SelectedItem.Index, lsvListView2)
'MsgBox lsvListView2.SelectedItem.Text & "-" & lsvListView2.SelectedItem.SubItems(1) & "-" & lsvListView2.SelectedItem.SubItems(2) & "-" & lsvListView2.SelectedItem.SubItems(3) & "-" & lsvListView2.SelectedItem.SubItems(4) & "-" & lsvListView2.SelectedItem.SubItems(5) & "-" & lsvListView2.SelectedItem.SubItems(6) & "-" & lsvListView2.SelectedItem.SubItems(7)

End Sub

Private Sub Timer1_Timer()
DoProcess
End Sub

VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   1845
   ClientTop       =   1905
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   6270
   Begin VB.CommandButton cmdCheckBalance 
      Caption         =   "Check Balance"
      Height          =   480
      Left            =   5070
      TabIndex        =   12
      Top             =   1455
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5010
      Top             =   2865
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "localhost"
      RemotePort      =   8888
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   5775
      Top             =   2145
   End
   Begin VB.CommandButton cmdHangup 
      Caption         =   "Hangup"
      Height          =   480
      Left            =   5385
      TabIndex        =   11
      Top             =   780
      Width           =   780
   End
   Begin VB.CommandButton cmdDial 
      Caption         =   "Dial"
      Height          =   480
      Left            =   5385
      TabIndex        =   10
      Top             =   285
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   420
      Left            =   4440
      TabIndex        =   9
      Top             =   285
      Width           =   855
   End
   Begin VB.ListBox Log 
      Height          =   1620
      ItemData        =   "Form1.frx":0000
      Left            =   180
      List            =   "Form1.frx":0002
      TabIndex        =   8
      Top             =   1485
      Width           =   4680
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "AT"
      Height          =   390
      Left            =   3810
      TabIndex        =   7
      Top             =   285
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5085
      Top             =   2190
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5520
      Top             =   2670
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Height          =   390
      Left            =   4110
      TabIndex        =   6
      Top             =   810
      Width           =   735
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   390
      Left            =   2820
      TabIndex        =   5
      Top             =   270
      Width           =   930
   End
   Begin VB.TextBox AT 
      Height          =   285
      Left            =   1245
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   2790
   End
   Begin VB.ComboBox Ports 
      Height          =   315
      ItemData        =   "Form1.frx":0004
      Left            =   960
      List            =   "Form1.frx":0006
      TabIndex        =   1
      Text            =   "COM1"
      Top             =   300
      Width           =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "Log:"
      Height          =   435
      Left            =   180
      TabIndex        =   4
      Top             =   1215
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "AT Command:"
      Height          =   300
      Left            =   195
      TabIndex        =   3
      Top             =   885
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "ComPort"
      Height          =   255
      Left            =   195
      TabIndex        =   0
      Top             =   345
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim gRespond As String
Const timeout As Long = 10
Dim maxreach As Boolean
Dim gComport As String
Dim gHanguped As Boolean


Private Sub LogMessage(msg As String, idx As Integer)
    Dim lmsg As String
    lmsg = "comport=" + Ports.Text + "&" + msg + "&time=" + Format(Now(), "dd/MM/yyy hh:mm:ss")
    Winsock1.SendData lmsg
    Log.AddItem lmsg, idx
End Sub


Public Sub cmdDial_Click()
    Dim command As String
    Dim try As Integer
    
    try = 3
    
dialagain:
    Timer2.Enabled = False
    try = try - 1
    
    Timer2.Enabled = True
    gHanguped = False
    command = "ATD" + gDialStr + Chr$(13)
    LogMessage "status=" + command, 0
    MSComm1.Output = command
           
    Counter = timeout * 3
    Do While (Counter)
       If InStr(1, gRespond, "NO CARRIER", vbTextCompare) > 0 Then
          LogMessage "status=no_carrier", 0
          If (try > 0) Then
            GoTo dialagain
          End If
          Exit Sub
       ElseIf gHanguped = True Then
          Exit Sub
       End If
       
       Sleep (1000)
       DoEvents
       Counter = Counter - 1
    Loop
    LogMessage "status=timeout", 0

End Sub

Public Sub cmdHangup_Click()
    Dim command As String
    command = "ATH" + Chr$(13)
    MSComm1.Output = command
    gHanguped = True
           
    Counter = timeout
    Do While (Counter)
       If InStr(1, gRespond, "ok", vbTextCompare) > 0 Then
          LogMessage "status=hangup", 0
          Exit Sub
       End If
       Sleep (1000)
       DoEvents
       Counter = Counter - 1
    Loop
    LogMessage "status=timeout", 0

End Sub


Public Sub cmdOpen_Click()
    MSComm1.Settings = "115200,n,8,1"
    MSComm1.CommPort = CLng(Mid(Ports.Text, 4))
    MSComm1.PortOpen = True
    Timer1.Enabled = True
    Form1.Caption = Ports.Text
    
End Sub

Public Sub cmdCheckBalance_Click()
''''OK
''''
''''+CUSD: 2,"Balance 0193903984:RM 15.00,valid until: 21/02/2013,tariff plan: Celco
''''m First 1+5.",15
    Dim sbalance As String
    Dim sphone As String
    Dim pos As Integer
    Dim lpos As Integer
    ''MSComm1.Output = "AT+CUSD=1,*124#" + vbCrLf
    MSComm1.Output = gCheckBalanceDialStr + vbCrLf
    
    Counter = timeout
    Do While (Counter)
       pos = InStr(1, gRespond, "Balance", vbTextCompare)
       If pos > 0 Then
          sbalance = Mid(gRespond, pos, 27)
          lpos = InStr(sbalance, ":")
          sphone = Trim(Mid(gRespond, pos + 8, lpos - 9))
          LogMessage "status=balance&other=" + sbalance + "&phone=" + sphone, 0
          Exit Sub
       End If
       Sleep (1000)
       DoEvents
       Counter = Counter - 1
    Loop
    LogMessage Ports.Text + " " + command + " timeout", 0
End Sub


Public Sub cmdTest_Click()
    Dim command As String
    
    command = "AT" + Chr$(13)
    MSComm1.Output = command
           
    stepthru = False
    Counter = timeout
    Do While (Counter)
       If InStr(1, gRespond, "ok", vbTextCompare) > 0 Then
          stepthru = True
          LogMessage "status=opened", 0
          Exit Sub
       End If
       Sleep (1000)
       DoEvents
       Counter = Counter - 1
    Loop
    LogMessage "status=timeout", 0
    
End Sub

Public Sub cmdClose_Click()
    Timer1.Enabled = False
    MSComm1.PortOpen = False
    LogMessage "status=closed", 0
    
End Sub


Public Sub Form_Load()
    Dim i As Integer
    Timer1.Enabled = False
    For i = 0 To 100
      Ports.AddItem "COM" + CStr(i)
    Next i
    

End Sub


Public Sub Timer1_Timer()
      gRespond = MSComm1.Input
      If Len(gRespond) > 0 Then
        ''LogMessage gRespond, 0
      End If
End Sub

Private Sub Timer2_Timer()
     Call cmdHangup_Click
     Timer2.Enabled = False
End Sub

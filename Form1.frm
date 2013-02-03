VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "SimMonitor"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Save 2 File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   45
      TabIndex        =   12
      Top             =   5235
      Width           =   840
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Check Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   0
      TabIndex        =   11
      Top             =   2955
      Width           =   930
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Unselect All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   30
      TabIndex        =   10
      Top             =   1380
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   30
      TabIndex        =   9
      Top             =   945
      Width           =   900
   End
   Begin VB.ComboBox cboCount 
      Height          =   315
      Left            =   2145
      TabIndex        =   6
      Text            =   "1"
      Top             =   420
      Width           =   810
   End
   Begin VB.ComboBox cboStart 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2145
      List            =   "Form1.frx":0031
      TabIndex        =   4
      Text            =   "1"
      Top             =   45
      Width           =   810
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Activate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   15
      TabIndex        =   3
      Top             =   2235
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5085
      Top             =   360
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   5625
      TabIndex        =   2
      Top             =   135
      Width           =   5280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "List ComPort"
      Height          =   705
      Left            =   3000
      TabIndex        =   1
      Top             =   45
      Width           =   990
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4560
      Top             =   315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   8888
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   8535
      Left            =   960
      TabIndex        =   0
      Top             =   930
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   15055
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ComPort"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Other"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Phone"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Message Log:"
      Height          =   300
      Left            =   4365
      TabIndex        =   8
      Top             =   105
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "Number of Ports:"
      Height          =   300
      Left            =   930
      TabIndex        =   7
      Top             =   465
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "First ComPort:"
      Height          =   270
      Left            =   960
      TabIndex        =   5
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim gColls As New Collection
Dim gItem As ListItem
Dim gMessages As New Collection

Private Sub Command1_Click()
    Dim item As ListItem
    Dim comport As String
    Dim start As Integer
    Dim number As Integer
    
    Set gColls = Nothing
    Set gColls = New Collection
    ListView1.ListItems.Clear
    
    start = CInt(cboStart.Text)
    number = CInt(cboCount.Text)
    
    For idx = start To start + number - 1
        comport = "COM" + CStr(idx)
        Set item = ListView1.ListItems.Add(, , UCase(comport))
        item.SubItems(1) = "idle"
        gColls.Add item, comport
    Next idx

End Sub

Private Function GetItem(comport As String) As ListItem
    On Error GoTo Handler
    Set GetItem = gColls.item(comport)
    Exit Function
Handler:
    Set GetItem = Nothing
End Function



Private Sub Command2_Click()
    For idx = 1 To ListView1.ListItems.Count
       ListView1.ListItems(idx).Checked = True
    Next idx
End Sub

Private Sub Command3_Click()
    Dim comport As String
    Dim filenum As Integer
    Dim extraparam As String
    extraparam = "-d 0123601348 -h 6000"
    
    filenum = FreeFile
    Open "config.txt" For Input As #filenum
    While EOF(filenum) = 0
      Line Input #filenum, tmp
      If Not Mid(tmp, 1, 2) = "##" Then
        If InStr(tmp, "activation=") Then
          extraparam = Mid(tmp, 12)
        End If
      End If
    Wend
    Close #filenum
    
    
    For idx = 1 To ListView1.ListItems.Count
       If ListView1.ListItems(idx).Checked = True Then
           comport = ListView1.ListItems(idx).Text
           ''cmd = "..\simactivate\simactivate -p " + comport + " -d 0123601349"
           params = "-p " + comport + " " + extraparam
           cmd = "simactivate.exe " + params
           List1.AddItem cmd, 0
           ShellEx "simactivate.exe", , params, Owner:=Me.hWnd
           DoEvents
       End If
     Next idx

End Sub

Private Sub Command4_Click()
    For idx = 1 To ListView1.ListItems.Count
       ListView1.ListItems(idx).Checked = False
    Next idx

End Sub


Private Sub Command5_Click()
    Dim comport As String
    Dim extraparam As String
    extraparam = "-d 0123601348 -c AT+CUSD=1,*124#"
    
    filenum = FreeFile
    Open "config.txt" For Input As #filenum
    While EOF(filenum) = 0
      Line Input #filenum, tmp
      If Not Mid(tmp, 1, 2) = "##" Then
        If InStr(tmp, "checkbalance=") Then
          extraparam = Mid(tmp, 14)
        End If
      End If
    Wend
    Close #filenum
    
    
    For idx = 1 To ListView1.ListItems.Count
       If ListView1.ListItems(idx).Checked = True Then
           comport = ListView1.ListItems(idx).Text
           params = "-p " + comport + " " + extraparam
           cmd = "simactivate.exe " + params
           List1.AddItem cmd, 0
           ShellEx "simactivate.exe", , params, Owner:=Me.hWnd
            
           ''retVal = Shell(cmd, vbNormalFocus)
           DoEvents
       End If
     Next idx
     ''retVal = Shell("run2.bat", vbNormalFocus)

    
End Sub

Private Sub Command6_Click()
    Dim item As ListItem
    Dim filename As String
    Dim filenum, Mode, Handle
    Dim comport As String
    Dim status As String
    Dim stime As String
    Dim phone As String
    Dim other As String
    
    filename = Format(Now(), "yyyy-MM-dd_hhmmss") + ".txt"
    
    filenum = FreeFile
    Open filename For Append As #filenum

    For idx = 1 To ListView1.ListItems.Count
       If ListView1.ListItems(idx).Checked = True Then
           Set item = ListView1.ListItems(idx)
           comport = item.Text
           status = item.SubItems(1)
           stime = item.SubItems(2)
           other = item.SubItems(3)
           phone = item.SubItems(4)
           Print #filenum, comport + "," + status + "," + stime + "," + other + "," + phone
           
       End If
     Next idx
     ''retVal = Shell("run2.bat", vbNormalFocus)
    Close filenum   ' Close file.

End Sub

Private Sub Form_Load()
    Winsock1.Bind 8888
    
    For idx = 1 To 100
      cboStart.AddItem CStr(idx)
      cboCount.AddItem CStr(idx)
    Next idx
    
End Sub

Private Sub ProcessMsg(msg As String)
    Dim status()  As String
    Dim comport As String
    Dim sstatus As String
    Dim item As ListItem
    Dim stime As String
    Dim other As String
    Dim phone As String
    
    
    status = Split(msg, "&")
        
    For idx = 0 To UBound(status)
        If (InStr(status(idx), "comport=") > 0) Then
            comport = Mid(status(idx), 9)
        ElseIf (InStr(status(idx), "status=") > 0) Then
            sstatus = Mid(status(idx), 8)
        ElseIf (InStr(status(idx), "time=") > 0) Then
            stime = Mid(status(idx), 6)
        ElseIf (InStr(status(idx), "other=") > 0) Then
            other = Mid(status(idx), 7)
        ElseIf (InStr(status(idx), "phone=") > 0) Then
            phone = Mid(status(idx), 7)
        End If
    Next idx
    
    Set item = GetItem(comport)
    If Not item Is Nothing Then
     item.SubItems(1) = sstatus
     item.SubItems(2) = stime
     If (Len(other) > 0) Then
        item.SubItems(3) = other
     End If
     If (Len(phone) > 0) Then
        item.SubItems(4) = phone
     End If
    End If
    
    
End Sub


Private Sub Timer1_Timer()
    Dim msg As String
    If gMessages.Count > 0 Then
        msg = gMessages.item(1)
        gMessages.Remove (1)
        Call ProcessMsg(msg)
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim msg  As String
    Winsock1.GetData msg
    List1.AddItem msg, 0
    
    gMessages.Add msg
    ''Call ProcessMsg(msg)

End Sub


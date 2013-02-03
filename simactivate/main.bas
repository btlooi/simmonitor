Attribute VB_Name = "Module1"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public gDialStr As String
Public gCheckBalanceDialStr As String


Sub Main()

    Dim strCmd As String
    Dim options() As String
    Dim idx As Integer
    Dim LastNonEmpty As Integer
    Dim checkbalance As Boolean
    
    checkbalance = False
    gDialStr = "0123601348"
    gCheckBalanceDialStr = "AT+CUSD=1,*124#"
    
    strCmd = command
    options = Split(strCmd, " ")
    
    Load Form1
    
    If (UBound(options) <= 0) Then
        Form1.Show
        Exit Sub
    End If
    
    For idx = 0 To UBound(options) Step 2
        If options(idx) = "-p" Then
            Form1.Ports.Text = options(idx + 1)
        ElseIf options(idx) = "-d" Then
            gDialStr = options(idx + 1)
        ElseIf options(idx) = "-c" Then
            checkbalance = True
            gCheckBalanceDialStr = options(idx + 1)
        End If
    Next idx
    
    Call Form1.cmdOpen_Click
    Sleep (1000)
    Call Form1.cmdTest_Click
    Sleep (1000)
    
    If checkbalance = False Then
        Call Form1.cmdDial_Click
        Sleep (1000)
    Else
        Call Form1.cmdCheckBalance_Click
        Sleep (1000)
    End If
    
    Call Form1.cmdClose_Click
    Sleep (1000)
    Unload Form1

    
End Sub


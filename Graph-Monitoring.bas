Sub start()
    Worksheets("Sheet1").Range("F1").Value = "RUNNING"
    Worksheets("Sheet1").Range("A2:A40").ClearContents
    t = Timer + 10
    pinging (t)
    
End Sub

Sub stop_ping()
    Worksheets("Sheet1").Range("F1").Value = "STOP"
End Sub


Function pinging(s)
    Dim target As String
    
    t = Timer
    
    'If s = t Then
        'Exit Function
    'End If
    
    Do Until Timer = t + 1
        DoEvents
    
    Loop
    
    target = Worksheets("Sheet1").Range("H1").Value
    
    If Worksheets("Sheet1").Range("F1").Value = "STOP" Then
        Exit Function
    End If
    
    Set rng1 = Worksheets("Sheet1").Range("A2:A40")
    lr = rng1.Count
    
    Worksheets("Sheet1").Cells(2, 1).Value = rsp_time(target)
    
    For x = lr To 2 Step -1
        Worksheets("Sheet1").Cells(1 + x, 1).Value = Worksheets("Sheet1").Cells(x, 1).Value
        If Worksheets("Sheet1").Range("F1").Value = "STOP" Then
            Exit For
        End If
    Next
        
    Call pinging(s)

End Function



Function rsp_time(target) As String
    Dim oPing As Object, oRetStatus As Object
    
    Set oPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
      ("select * from Win32_PingStatus where address = '" & target & "'")

    For Each oRetStatus In oPing
        If IsNull(oRetStatus.StatusCode) Or oRetStatus.StatusCode <> 0 Then
            rsp_time = -5
            Worksheets("Sheet1").Range("H2").Value = "Offline"
            Worksheets("Sheet1").Range("H2").Interior.ColorIndex = 0
            Worksheets("Sheet1").Range("H2").Font.Color = RGB(200, 0, 0)
            Application.Wait (Now + TimeValue("0:00:01"))
            Worksheets("Sheet1").Range("H2").Interior.ColorIndex = 6
        Else

            rsp_time = oRetStatus.ResponseTime
            Worksheets("Sheet1").Range("H2").Value = "Online"
            Worksheets("Sheet1").Range("H2").Interior.ColorIndex = 2
            Worksheets("Sheet1").Range("H2").Font.Color = RGB(0, 0, 0)
            Application.Wait (Now + TimeValue("0:00:01"))
            Worksheets("Sheet1").Range("H2").Interior.ColorIndex = 4

        End If
    Next

End Function



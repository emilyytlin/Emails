Sub Send()
    Dim n As Integer
    Dim email As String
    For n = 0 To 37
        email = Sheets("1").Range("A2").Offset(n, 0).Value
        'MsgBox email & " " & n + 2
        Call Mail(email, n + 2)
    Next n
End Sub

Sub Mail(ByVal toAddress As String, ByVal i As Integer)
    Dim iMsg As Object
    Dim iConf As Object
    Dim strbody As String

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    
    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "user@gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "pass"
        .Update
    End With
    
    strbody = "Dear " & Cells(i, 4) & " " & Cells(i, 3) & "," & vbNewLine & vbNewLine & _
        "You receive a " & Cells(i, 7) & "/100 total." & vbNewLine & _
        "Part I: " & Cells(i, 5) & " Part II: " & Cells(i, 6) & vbNewLine & vbNewLine & _
        Cells(i, 8) & vbNewLine & vbNewLine & _
        "Regards," & vbNewLine & _
        "Emily Lin, Grader" & vbNewLine & _
        "Introduction to Computer Programming"

    With iMsg
        Set .Configuration = iConf
        .To = toAddress
        .CC = ""
        .BCC = ""
        .From = """Emily Lin"" <user@gmail.com>"
        .Subject = "CS 002: Assignment 1 grade"
        .TextBody = strbody
        .Send
    End With
End Sub

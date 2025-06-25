Attribute VB_Name = "Subprocess"
Option Explicit


Function RunAndCapture(ByVal Command As String) As Dictionary
    ' Run a process and capture its stdout, stderr, and exit code.
    '
    ' FIXME: This flashes a window. I could not find any other way to do this in VBA.
    ' Don't use this in projects that have Python, since it's much better to do this in Python.
    '
    ' Args:
    '   command: Command to execute.
    '
    ' Returns:
    '   A dictionary containing:
    '   - stdout: A list of lines in the standard output stream (collection of strings).
    '   - stderr: A list of lines in the standard error stream (collection of string).
    '   - code: The exit code (integer).
    
    Dim Shell As Object
    Set Shell = CreateObject("WScript.Shell")
    
    Dim Exec As Object
    Set Exec = Shell.Exec(Command)
    
    Dim Stdout As Collection
    Set Stdout = Col()
    Do While Not Exec.Stdout.AtEndOfStream
        Stdout.Add Exec.Stdout.ReadLine()
    Loop
    
    Dim Stderr As Collection
    Set Stderr = Col()
    Do While Not Exec.Stderr.AtEndOfStream
        Stderr.Add Exec.Stderr.ReadLine()
    Loop
    
    Set RunAndCapture = Dict( _
        "stdout", Stdout, _
        "stderr", Stderr, _
        "code", Exec.ExitCode _
    )
End Function


Function WhereIsExe(ByVal Name As String) As String
    ' Use the `where.exe` utility in Windows to find the full path to an executable that is in the PATH.
    '
    ' Don't use this in projects that have Python, since it's much better to do this in Python.
    '
    ' Args:
    '   name: The filename of the executable, e.g. "python.exe"
    '
    ' Returns:
    '   The full path of the executable, if found.
    '   Otherwise, an empty string.
    
    Dim Result As Dictionary
    Set Result = RunAndCapture("where.exe " & Name)
    
    Dim Stdout As Collection
    Set Stdout = Result("stdout")
    
    Dim FirstLine As String
    If Stdout.Count Then
        WhereIsExe = Trim(Stdout(1))
    Else
        WhereIsExe = ""
    End If
End Function


Function EscapeAndWrapCmdArg(InputStr As String) As String
    ' Escape a command line argument and wrap it in double quotes, so that the windows shell interprets it as we intend.
    '
    ' Args:
    '   inputStr: The argument to escape and wrap.
    '
    ' Returns:
    '   The escaped and wrapped argument.
    
    Dim I As Integer
    Dim ResultStr As String
    Dim BackslashCount As Integer

    ' Rule 1: Backslash behaviour is different when preceding a quote than when found in isolation.
    ' Escape them by doubling them up before a quote
    I = 1
    While I <= Len(InputStr)
        If Mid(InputStr, I, 1) = "\" Then
            BackslashCount = 1
            While I + BackslashCount <= Len(InputStr) And Mid(InputStr, I + BackslashCount, 1) = "\"
                BackslashCount = BackslashCount + 1
            Wend

            If I + BackslashCount <= Len(InputStr) And Mid(InputStr, I + BackslashCount, 1) = """" Then
                ResultStr = ResultStr & String(BackslashCount * 2, "\")
                I = I + BackslashCount - 1
            Else
                ResultStr = ResultStr & Mid(InputStr, I, 1)
            End If
        Else
            ResultStr = ResultStr & Mid(InputStr, I, 1)
        End If
        I = I + 1
    Wend

    ' Rule 2: Escape all quotes with a backslash
    InputStr = ResultStr
    ResultStr = ""
    For I = 1 To Len(InputStr)
        If Mid(InputStr, I, 1) = """" Then
            ResultStr = ResultStr & "\"
        End If
        ResultStr = ResultStr & Mid(InputStr, I, 1)
    Next I

    ' Rule 3: Wrap result in non-escaped quotes
    EscapeAndWrapCmdArg = """" & ResultStr & """"
End Function


Function FindBestTerminal() As String
    ' Find the full path to the best available terminal app.
    ' This can be prepended to almost any command to run that command in that terminal.
    '
    ' Don't use this in projects that have Python, since it's much better to do this in Python.
    '
    ' Returns:
    '   The full path to the best available terminal app.
    '   If none are found, an empty string is returned.
    '   It will never return the path to `cmd.exe`,
    '   since we don't consider that to be a good terminal,
    '   and it's usually the default anyway.
    
    ' Check if the Windows Terminal app is available.
    ' This is a modern terminal with good features.
    ' It is available on many Windows 10 PCs and all Windows 11 PCs.
    FindBestTerminal = WhereIsExe("wt.exe")
    If Len(FindBestTerminal) Then
        Exit Function
    End If
    
    ' The Windows Terminal app is not available.
    ' Try powershell.exe next, because it is a better terminal than cmd.exe
    FindBestTerminal = WhereIsExe("powershell.exe")
    If Len(FindBestTerminal) Then
        Exit Function
    End If

    ' Powershell is also not available.
    ' We should just call the command directly. It will probably open in a cmd.exe window, which is not ideal, but works.
    FindBestTerminal = ""
End Function


'Allows you run an external cmd.exe command by setting the cmdArgs variable as desired
const module_name	= "runCMD"
const module_ver	= "1.0"
const module_title	= "runCMD"

sub Init
  addMenuItem "&RunCMD", "", "externalCommand", "Shift+Ctrl+Alt+C"
end sub

'Gets ParentFolder?
Function ExtractFilePath( strPath )
    If Len(strPath) = 0 Then
      Exit Function                                    ' input string is empty
    Else
      strPath = Replace(strPath, Chr(47), Chr(92))     ' convert backslashes to forward slashes
      If InStr(1, strPath, Chr(92)) = 0 Then
        Exit Function                                  ' string contains no forward slashes
      End If
    End If
    ExtractFilePath = Left(strPath, InStrRev(strPath, Chr(92)))
End Function


sub externalCommand
  Set activeEditor = newEditor()
  activeEditor.assignActiveEditor()

    cmdArgs = Chr(34) & ExtractFilePath(activeEditor.fileName()) & Chr(34)
     'NOTE: cmdArgs can contain multiple commands by separating them with && like this: cmdArgs = "cd\php && php.exe"

    Set wshShell = CreateObject( "WScript.Shell" )
    wshShell.Run "cmd.exe /K cd /d " & cmdArgs , 1, False
     'NOTE:                               1 = show dos window
     '                                    0 = hide dos window
     '                    /K = keep dos window open when application terminates
     '                    /C = close dos window when application terminates

    Set wshShell = Nothing
    Set editor = Nothing
end sub
Attribute VB_Name = "Logging"
Private Sub LogFileMacro(caller As String, message As String, level As String)
Dim workDir As String
Dim macroVersion As String
Dim strFile_Path As String, strFilePathShared As String, dir_Path As String, logMessage As String

macroVersion = Config.[B19].Value

strFile_Path = Environ(“LOCALAPPDATA”) & "\" & "FORMEXTRACT.log"

timeStamp = Format(Now(), "M/d/yyyy hh:mm:ss AM/PM")

logMessage = "[" & level & "] " & timeStamp & " - FORMEXTRACT V_" & macroVersion & " - " & caller & " - " & Replace(Replace(message, Chr(13), ""), Chr(10), "")

Open strFile_Path For Append As #1
Print #1, logMessage
Close #1

Debug.Print logMessage

End Sub

Public Sub LogError(ByVal caller As String, ByVal message As String)

LogFileMacro caller, message, "ERR"

End Sub

Public Sub LogInfo(ByVal caller As String, ByVal message As String)
On error resume next
LogFileMacro caller, message, "INF"

End Sub

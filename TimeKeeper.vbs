Option Explicit

' Tắt hiển thị các lỗi
On Error Resume Next

' Khởi tạo đối tượng WScript.Shell và biến cần thiết
Dim WshShell: Set WshShell = CreateObject("WScript.Shell")
Dim serviceQuery, oExec

' Kiểm tra trạng thái dịch vụ Windows Time
Set oExec = WshShell.Exec("sc query w32time")
serviceQuery = oExec.StdOut.ReadAll()

' Kiểm tra và xử lý trạng thái dịch vụ Windows Time
If InStr(1, serviceQuery, "DOES_NOT_EXIST", vbTextCompare) > 0 Then
    ' Dịch vụ không tồn tại, tạo và đăng ký
    WshShell.Run "w32tm /register", 0, False
    WshShell.Run "sc config w32time start= auto", 0, False
ElseIf InStr(1, serviceQuery, "DISABLED", vbTextCompare) > 0 Then
    ' Dịch vụ bị vô hiệu hóa, kích hoạt lại
    WshShell.Run "sc config w32time start= auto", 0, False
End If

' Kiểm tra lại xem dịch vụ có bị dừng không
If InStr(1, serviceQuery, "STOPPED", vbTextCompare) > 0 OR InStr(1, serviceQuery, "DISABLED", vbTextCompare) > 0 Then
    ' Khởi động dịch vụ nếu đang dừng
    WshShell.Run "net start w32time", 0, False
    WScript.Sleep 1000
Else
    ' Dừng dịch vụ để chuẩn bị cấu hình lại
    WshShell.Run "net stop w32time", 0, False
    WScript.Sleep 1000
End If

' Thiết lập lại hoàn toàn dịch vụ Windows Time
WshShell.Run "w32tm /unregister", 0, False
WshShell.Run "w32tm /register", 0, False
WshShell.Run "sc config w32time start= auto", 0, False

' Thiết lập các tham số registry quan trọng
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v MaxPosPhaseCorrection /t REG_DWORD /d 0xFFFFFFFF /f", 0, False
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v MaxNegPhaseCorrection /t REG_DWORD /d 0xFFFFFFFF /f", 0, False
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"" /v NtpServer /t REG_SZ /d ""time.windows.com,0x1 time.google.com,0x1 pool.ntp.org,0x1 vn.pool.ntp.org,0x1"" /f", 0, False
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"" /v Type /t REG_SZ /d ""NTP"" /f", 0, False
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"" /v SpecialPollInterval /t REG_DWORD /d 300 /f", 0, False
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"" /v Enabled /t REG_DWORD /d 1 /f", 0, False
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpServer"" /v Enabled /t REG_DWORD /d 0 /f", 0, False

' Đợi một chút để các lệnh registry có thời gian áp dụng
WScript.Sleep 1000

' Khởi động dịch vụ Windows Time
WshShell.Run "net start w32time", 0, False

' Đợi dịch vụ khởi động
WScript.Sleep 2000

' Cập nhật cấu hình và buộc đồng bộ thời gian
WshShell.Run "w32tm /config /update", 0, False
WshShell.Run "w32tm /resync /force", 0, False

' Đợi và đồng bộ lại lần nữa
WScript.Sleep 3000
WshShell.Run "w32tm /resync /rediscover /force", 0, False

' Đợi một chút và đồng bộ lần cuối
WScript.Sleep 2000
WshShell.Run "w32tm /resync /force", 0, False

' Giải phóng tài nguyên
Set WshShell = Nothing
Set oExec = Nothing

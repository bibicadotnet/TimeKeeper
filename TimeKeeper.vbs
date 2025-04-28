Option Explicit

' Tắt hiển thị các lỗi
On Error Resume Next

' Khởi tạo đối tượng WScript.Shell và biến cần thiết
Dim WshShell: Set WshShell = CreateObject("WScript.Shell")
Dim serviceQuery, oExec, retVal

' ===== KIỂM TRA TRẠNG THÁI DỊCH VỤ WINDOWS TIME =====
Set oExec = WshShell.Exec("sc query w32time")
serviceQuery = oExec.StdOut.ReadAll()

' Đảm bảo dịch vụ Windows Time tồn tại và được kích hoạt
If InStr(1, serviceQuery, "DOES_NOT_EXIST", vbTextCompare) > 0 Then
    ' Dịch vụ không tồn tại, đăng ký mới
    WshShell.Run "w32tm /register", 0, True
    WshShell.Run "sc config w32time start= auto", 0, True
ElseIf InStr(1, serviceQuery, "DISABLED", vbTextCompare) > 0 Then
    ' Dịch vụ bị vô hiệu hóa, kích hoạt lại
    WshShell.Run "sc config w32time start= auto", 0, True
End If

' ===== DỪNG DỊCH VỤ WINDOWS TIME TRƯỚC KHI CẤU HÌNH =====
WshShell.Run "net stop w32time", 0, True
WScript.Sleep 1000

' ===== ĐẶT LẠI CẤU HÌNH DỊCH VỤ WINDOWS TIME =====
WshShell.Run "w32tm /unregister", 0, True
WshShell.Run "w32tm /register", 0, True

' ===== THIẾT LẬP CÁC THAM SỐ REGISTRY CHO WINDOWS TIME =====
' Cấu hình tổng quan
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v UpdateInterval /t REG_DWORD /d 100 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v MaxPosPhaseCorrection /t REG_DWORD /d 0xFFFFFFFF /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v MaxNegPhaseCorrection /t REG_DWORD /d 0xFFFFFFFF /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v FrequencyCorrectRate /t REG_DWORD /d 2 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v PollAdjustFactor /t REG_DWORD /d 1 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v PhaseCorrectRate /t REG_DWORD /d 1 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v MinPollInterval /t REG_DWORD /d 6 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config"" /v MaxPollInterval /t REG_DWORD /d 10 /f", 0, True

' Cấu hình máy chủ NTP
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"" /v NtpServer /t REG_SZ /d ""time.windows.com,0x9 time.google.com,0x9 pool.ntp.org,0x9 vn.pool.ntp.org,0x9"" /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"" /v Type /t REG_SZ /d ""NTP"" /f", 0, True

' Cấu hình NTP Client
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"" /v SpecialPollInterval /t REG_DWORD /d 900 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"" /v Enabled /t REG_DWORD /d 1 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"" /v CrossSiteSyncFlags /t REG_DWORD /d 2 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"" /v ResolvePeerBackoffMinutes /t REG_DWORD /d 1 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"" /v ResolvePeerBackoffMaxTimes /t REG_DWORD /d 5 /f", 0, True
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"" /v EventLogFlags /t REG_DWORD /d 0 /f", 0, True

' Tắt NTP Server
WshShell.Run "reg add ""HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpServer"" /v Enabled /t REG_DWORD /d 0 /f", 0, True

' ===== KHỞI ĐỘNG LẠI DỊCH VỤ WINDOWS TIME =====
WshShell.Run "sc config w32time start= auto", 0, True
WshShell.Run "net start w32time", 0, True
WScript.Sleep 2000

' ===== CẬP NHẬT CẤU HÌNH VÀ BUỘC ĐỒNG BỘ THỜI GIAN =====
' Cập nhật cấu hình
WshShell.Run "w32tm /config /update", 0, True
WScript.Sleep 1000

' Đồng bộ thời gian với các máy chủ NTP
WshShell.Run "w32tm /resync /rediscover", 0, True
WScript.Sleep 5000

' Đồng bộ lại lần nữa để đảm bảo
WshShell.Run "w32tm /resync /force", 0, True

' ===== TỐI ƯU CPU ĐỂ TRÁNH TRÔI THỜI GIAN =====
WshShell.Run "powercfg -setacvalueindex scheme_current sub_processor PERFINCPOL 2", 0, True
WshShell.Run "powercfg -setacvalueindex scheme_current sub_processor PERFDECPOL 1", 0, True
WshShell.Run "powercfg -setactive scheme_current", 0, True

' Giải phóng tài nguyên
Set WshShell = Nothing
Set oExec = Nothing

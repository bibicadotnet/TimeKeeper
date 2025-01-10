' Khai báo biến
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

' Thiết lập cấu hình Windows Time với độ chính xác cao
Sub ConfigureWindowsTime()
    ' Dừng service trước khi cấu hình
    WshShell.Run "net stop w32time", 0, True
    WScript.Sleep 1000

    ' Cấu hình độ chính xác cao
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v MaxPollInterval /t REG_DWORD /d 6 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v MinPollInterval /t REG_DWORD /d 6 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v UpdateInterval /t REG_DWORD /d 100 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v FrequencyCorrectRate /t REG_DWORD /d 2 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v MaxNegPhaseCorrection /t REG_DWORD /d 1 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v MaxPosPhaseCorrection /t REG_DWORD /d 1 /f", 0, True
    
    ' Cấu hình NTP client để ưu tiên độ chính xác
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Parameters /v Type /t REG_SZ /d NTP /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Parameters /v ServiceDll /t REG_EXPAND_SZ /d C:\Windows\system32\w32time.dll /f", 0, True
    
    ' Cấu hình time servers với nhiều server dự phòng
    WshShell.Run "w32tm /config /manualpeerlist:""time.windows.com,0x1 time.nist.gov,0x1 pool.ntp.org,0x1"" /syncfromflags:manual /reliable:yes /update", 0, True
    
    ' Khởi động lại service
    WshShell.Run "net start w32time", 0, True
    WScript.Sleep 2000
End Sub

' Kiểm tra và đồng bộ thời gian
Sub SynchronizeTime()
    ' Lấy thông tin về độ lệch thời gian
    Dim exec, output
    Set exec = WshShell.Exec("w32tm /stripchart /computer:time.windows.com /dataonly /samples:1")
    output = exec.StdOut.ReadAll
    
    ' Phân tích độ lệch
    If InStr(output, ",") > 0 Then
        Dim parts, offset, offsetNum
        parts = Split(output, ",")
        If UBound(parts) >= 1 Then
            offset = Trim(parts(1))
            offset = Replace(Replace(offset, vbCrLf, ""), "seconds", "")
            offset = Replace(offset, vbLf, "")
            offset = Replace(offset, "s", "")
            offset = Trim(offset)
            
            On Error Resume Next
            offsetNum = CDbl(offset)
            If Err.Number = 0 Then
                ' Nếu độ lệch lớn hơn 0.1s, thực hiện đồng bộ
                If Abs(offsetNum) > 0.1 Then
                    WshShell.Run "w32tm /resync /force /nowait", 0, True
                End If
            End If
            On Error GoTo 0
        End If
    End If
End Sub

' Hàm chính
On Error Resume Next

' Cấu hình Windows Time Service
ConfigureWindowsTime

' Thực hiện đồng bộ
SynchronizeTime

Set WshShell = Nothing
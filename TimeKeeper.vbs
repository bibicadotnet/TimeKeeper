' Khai báo biến
Option Explicit
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

' Kiểm tra và kích hoạt service Windows Time
Sub EnableWindowsTime()
    On Error Resume Next
    Dim exec, output
    Set exec = WshShell.Exec("sc query w32time")
    output = exec.StdOut.ReadAll
    
    If InStr(output, "STOPPED") > 0 Or InStr(output, "DISABLED") > 0 Then
        WshShell.Run "sc config w32time start= auto", 0, True
        WScript.Sleep 2000
        WshShell.Run "net start w32time", 0, True
        WScript.Sleep 2000
    End If
    On Error Goto 0
End Sub

' Thiết lập cấu hình Windows Time với độ chính xác cao
Sub ConfigureWindowsTime()
    On Error Resume Next
    WshShell.Run "net stop w32time", 0, True
    WScript.Sleep 2000
    
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v MaxPollInterval /t REG_DWORD /d 6 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v MinPollInterval /t REG_DWORD /d 6 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v UpdateInterval /t REG_DWORD /d 100 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v FrequencyCorrectRate /t REG_DWORD /d 2 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v MaxNegPhaseCorrection /t REG_DWORD /d 1 /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config /v MaxPosPhaseCorrection /t REG_DWORD /d 1 /f", 0, True
    
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Parameters /v Type /t REG_SZ /d NTP /f", 0, True
    WshShell.Run "reg add HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Parameters /v ServiceDll /t REG_EXPAND_SZ /d C:\Windows\system32\w32time.dll /f", 0, True
    
    WshShell.Run "w32tm /config /manualpeerlist:""time.windows.com,0x1 time.nist.gov,0x1 pool.ntp.org,0x1"" /syncfromflags:manual /reliable:yes /update", 0, True
    
    WshShell.Run "net start w32time", 0, True
    WScript.Sleep 2000
    WshShell.Run "w32tm /config /update", 0, True
    On Error Goto 0
End Sub

' Kiểm tra và đồng bộ thời gian
Sub SynchronizeTime()
    On Error Resume Next
    Dim exec, output
    Set exec = WshShell.Exec("sc query w32time")
    output = exec.StdOut.ReadAll
    
    If InStr(output, "RUNNING") > 0 Then
        WshShell.Run "w32tm /resync /force /nowait", 0, True
        WScript.Sleep 5000
        
        Set exec = WshShell.Exec("w32tm /stripchart /computer:time.windows.com /dataonly /samples:1")
        output = exec.StdOut.ReadAll
        
        If InStr(output, ",") > 0 Then
            Dim parts, offset, offsetNum
            parts = Split(output, ",")
            If UBound(parts) >= 1 Then
                offset = Trim(parts(1))
                offset = Replace(Replace(offset, vbCrLf, ""), "seconds", "")
                offset = Replace(offset, vbLf, "")
                offset = Replace(offset, "s", "")
                offset = Trim(offset)
                
                offsetNum = CDbl(offset)
                If Abs(offsetNum) > 0.1 Then
                    WScript.Sleep 2000
                    WshShell.Run "w32tm /resync /force /nowait", 0, True
                End If
            End If
        End If
    End If
    On Error Goto 0
End Sub

' Thực thi chính
On Error Resume Next
EnableWindowsTime
ConfigureWindowsTime
SynchronizeTime
Set WshShell = Nothing

@echo off
setlocal
echo ============================================
echo  Lab 3 Packet Capture Script (Fixed)
echo ============================================

set TSHARK="C:\Program Files\Wireshark\tshark.exe"
set OUTDIR=C:\Users\User\Downloads\lab3

echo.
echo [Step 1] Killing Surfshark and disabling WireGuard adapter...
taskkill /IM "Surfshark.WireguardServic" /F >nul 2>&1
taskkill /IM "Surfshark.exe" /F >nul 2>&1
taskkill /IM "Surfshark.Service.exe" /F >nul 2>&1
timeout /t 2 /nobreak >nul

REM Disable WireGuard tunnel adapter to stop VPN
powershell -Command "Start-Process powershell -ArgumentList '-Command','Disable-NetAdapter -Name \"Local Area Connection* 9\" -Confirm:$false; Disable-NetAdapter -Name \"SurfsharkWireGuard\" -Confirm:$false' -Verb RunAs -Wait" 2>nul
timeout /t 3 /nobreak >nul
echo    VPN fully disconnected.

echo.
echo [Step 2] Verifying direct internet...
ping -n 1 www.google.com
echo.

echo ============================================
echo  CAPTURE 1: ICMP Ping (10 pings to www.ust.hk)
echo ============================================
start /B "" %TSHARK% -i 6 -a duration:60 -w "%OUTDIR%\lab3_icmp_ping.pcapng"
timeout /t 3 /nobreak >nul
ping -n 10 www.ust.hk
timeout /t 3 /nobreak >nul
taskkill /IM tshark.exe /F >nul 2>&1
timeout /t 2 /nobreak >nul
echo    Saved: lab3_icmp_ping.pcapng

echo ============================================
echo  CAPTURE 2: ICMP Tracert (to www.inria.fr)
echo ============================================
start /B "" %TSHARK% -i 6 -a duration:120 -w "%OUTDIR%\lab3_icmp_tracert.pcapng"
timeout /t 3 /nobreak >nul
tracert www.inria.fr
timeout /t 3 /nobreak >nul
taskkill /IM tshark.exe /F >nul 2>&1
timeout /t 2 /nobreak >nul
echo    Saved: lab3_icmp_tracert.pcapng

echo ============================================
echo  CAPTURE 3: IP combined (56, 2000, 3500 bytes)
echo ============================================
start /B "" %TSHARK% -i 6 -a duration:180 -w "%OUTDIR%\lab3_ip_combined.pcapng"
timeout /t 3 /nobreak >nul
echo    --- 56 byte ping ---
ping -n 5 -l 56 www.inria.fr
echo    --- 2000 byte ping ---
ping -n 5 -l 2000 www.inria.fr
echo    --- 3500 byte ping ---
ping -n 5 -l 3500 www.inria.fr
timeout /t 3 /nobreak >nul
taskkill /IM tshark.exe /F >nul 2>&1
timeout /t 2 /nobreak >nul
echo    Saved: lab3_ip_combined.pcapng

echo.
echo ============================================
echo  Re-enabling adapter and restarting Surfshark...
echo ============================================
powershell -Command "Start-Process powershell -ArgumentList '-Command','Enable-NetAdapter -Name \"Local Area Connection* 9\" -Confirm:$false; Enable-NetAdapter -Name \"SurfsharkWireGuard\" -Confirm:$false' -Verb RunAs -Wait" 2>nul
start "" "C:\Program Files\Surfshark\Surfshark.exe"
echo    Done. Reconnect VPN in Surfshark.
echo.
echo  ALL CAPTURES COMPLETE!
echo ============================================
pause

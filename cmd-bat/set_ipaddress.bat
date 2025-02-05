@echo off
netsh interface ip set address name="local" static 172.16.0.2 255.255.255.0 172.16.0.1
pause
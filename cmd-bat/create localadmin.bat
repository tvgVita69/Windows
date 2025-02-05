@echo off
chcp 1251
net user admin 123 /add
net localgroup "Администраторы" admin /add
pause
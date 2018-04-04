@echo off
md \\agw-fs\access\%username%
copy "\\agw-fs\Be2\agworks8.accdb" \\agw-fs\Access\%username%\  /Y
copy "\\agw-fs\Be2\ImportWorkers.xlsx" \\agw-fs\Access\%username%\  /Y
copy "\\agw-fs\Be2\favicon.ico" \\agw-fs\Access\%username%\  /Y
copy "\\Agw-fs\Be2\agworks8_processes_redirect.txt" "\\agw-fs\Access\%username%\" /y


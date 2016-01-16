taskkill /im explorer.exe /f
timeout 1
start explorer
timeout 1
start explorer %~dp0

@echo off
chcp 65001 > nul

set "章回目录=d:\AI\books\白眉大侠\章回"
set "分段程序=d:\AI\books\自动分段程序.py"

for /l %%i in (1,1,145) do (
    setlocal enabledelayedexpansion
    set "序号=000%%i"
    set "序号=!序号:~-3!"
    set "源文件=!章回目录!\!序号!.txt"
    
    echo 正在处理: !源文件!
    python "!分段程序!" "!源文件!"
    
    endlocal
)

echo 所有章回分段完成！
pause
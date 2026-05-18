@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
echo 正在合并 page_005.txt 到 page_521.txt ...

:: 进入 txt 子目录（加引号，解决中文路径问题）
cd "txt"

:: 清空结果文件（中文路径必须引号）
type nul > "合并结果.txt"

:: 循环 5-521，自动补 3 位
for /l %%i in (5,1,521) do (
    set "num=00%%i"
    set "num=!num:~-3!"
    :: 合并（全部加引号，解决中文路径报错）
    copy /b "合并结果.txt" + "page_!num!.txt" "合并结果.txt" >nul
    echo 已合并：page_!num!.txt
)

cd..
move "txt\合并结果.txt" "合并结果.txt"

echo.
echo ==========================
echo 合并完成！文件：合并结果.txt
echo ==========================
pause
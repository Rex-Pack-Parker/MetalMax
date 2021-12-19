@Echo OFF
Title Metal Max 系列整合  初始化工具
Color 70
setlocal EnableDelayedExpansion
cd /d %~dp0
mode con: cols=120 lines=40

echo --描述--
echo  Metal Max 系列整合  初始化工具
echo  by:雷克斯.派
echo  --重要的事情索3遍！！
echo   对照下面的 当前路径  如果文件夹名称包含中文和符号 请修改文件夹名称再执行此程序.
echo   对照下面的 当前路径  如果文件夹名称包含中文和符号 请修改文件夹名称再执行此程序.
echo   对照下面的 当前路径  如果文件夹名称包含中文和符号 请修改文件夹名称再执行此程序.
echo …
echo   [ROM 游戏存档 即时存档 金手指] 的目录全指向Data文件夹,以便备份和通用性.
echo …

echo --初始--
echo  当前路径:%~dp0
echo …

echo --执行--
for /d %%i in (*) do (
	if %%i neq Data (
		Echo 处理:%%i
		rd "%~dp0%%i\ROM" 2>nul
		rd "%~dp0%%i\Save" 2>nul
		rd "%~dp0%%i\State" 2>nul
		rd "%~dp0%%i\Cheat" 2>nul
		mklink /j "%~dp0%%i\ROM" "%~dp0Data\_ROM" >nul
		mklink /j "%~dp0%%i\Save" "%~dp0Data\_Saves" >nul
		mklink /j "%~dp0%%i\State" "%~dp0Data\_States" >nul
		mklink /j "%~dp0%%i\Cheat" "%~dp0Data\_Cheats" >nul
	)
)

echo …
echo --完成--
echo  如遇运行方面的问题,请再次运行此程序,截图来或者按Alt＋Print键
echo  不然草泥马~~
echo …

pause
@Echo OFF
Title Metal Max ϵ������  ����ROM�������
Color A0
setlocal EnableDelayedExpansion
cd /d %~dp0
mode con: cols=120 lines=20

for /d %%i in (*) do (
	if %%i neq Data (
	echo ɾ�� %~dp0%%i\ROM
		rd "%~dp0%%i\ROM"
		rd "%~dp0%%i\Save"
		rd "%~dp0%%i\State"
		rd "%~dp0%%i\Cheat"
	)
)

echo --���--
pause
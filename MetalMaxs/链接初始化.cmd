@Echo OFF
Title Metal Max ϵ������  ��ʼ������
Color 70
setlocal EnableDelayedExpansion
cd /d %~dp0
mode con: cols=120 lines=40

echo --����--
echo  Metal Max ϵ������  ��ʼ������
echo  by:�׿�˹.��
echo  --��Ҫ��������3�飡��
echo   ��������� ��ǰ·��  ����ļ������ư������ĺͷ��� ���޸��ļ���������ִ�д˳���.
echo   ��������� ��ǰ·��  ����ļ������ư������ĺͷ��� ���޸��ļ���������ִ�д˳���.
echo   ��������� ��ǰ·��  ����ļ������ư������ĺͷ��� ���޸��ļ���������ִ�д˳���.
echo ��
echo   [ROM ��Ϸ�浵 ��ʱ�浵 ����ָ] ��Ŀ¼ȫָ��Data�ļ���,�Ա㱸�ݺ�ͨ����.
echo ��

echo --��ʼ--
echo  ��ǰ·��:%~dp0
echo ��

echo --ִ��--
for /d %%i in (*) do (
	if %%i neq Data (
		Echo ����:%%i
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

echo ��
echo --���--
echo  �������з��������,���ٴ����д˳���,��ͼ�����߰�Alt��Print��
echo  ��Ȼ������~~
echo ��

pause
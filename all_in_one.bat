@echo off
color 0A
cls
echo ============================================
echo         �Αӕ\���������c�[�� ���j���[
echo ============================================
echo.

:MENU
echo �����I�����Ă�������:
echo.
echo 1. �Α�CSV�����i�ŐV�t�@�C���j
echo 2. �Α�CSV�����i�t�@�C���I���j
echo 3. �t�H���_�Ď��J�n�i�Ɩ����ԁj
echo 4. kintone����f�[�^�擾
echo 5. kintone�Ƀf�[�^���M
echo 6. ��o�󋵊m�F�ƃ��}�C���h���M
echo 7. �ݒ���̕\��
echo 8. �I��
echo.

set /p choice="�ԍ�����͂��Ă������� (1-8): "

if "%choice%"=="1" goto NORMAL
if "%choice%"=="2" goto SELECT_FILE
if "%choice%"=="3" goto WATCH
if "%choice%"=="4" goto KINTONE_GET
if "%choice%"=="5" goto KINTONE_PUSH
if "%choice%"=="6" goto NOTIFY
if "%choice%"=="7" goto INFO
if "%choice%"=="8" goto END

echo �����ȑI���ł��B������x���������������B
goto MENU

:NORMAL
cls
echo �ŐV��CSV�t�@�C�����������܂�...
call run.bat
pause
cls
goto MENU

:SELECT_FILE
cls
echo CSV�t�@�C����I�����Ă��������B
set "psCommand="(new-object -COM 'Shell.Application').BrowseForFolder(0,'CSV�t�@�C����I�����Ă�������',0,0).self.path""
for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "folder=%%I"
if "%folder%"=="" goto MENU

set "psCommand="Get-ChildItem -Path '%folder%' -Filter *.csv | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName""
for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "csv_file=%%I"

if "%csv_file%"=="" (
  echo CSV�t�@�C����������܂���ł����B
  pause
  cls
  goto MENU
)

echo �I�����ꂽ�t�@�C��: %csv_file%
echo �������J�n���܂�...

python main.py run --file "%csv_file%" --template "templates\�Αӕ\���`_2025�N��.xlsx"

if %errorlevel% neq 0 (
  echo �G���[���������܂����B
) else (
  echo �������������܂����B�o�̓t�H���_���J���܂��B
  explorer output
)

pause
cls
goto MENU

:WATCH
cls
echo �t�H���_�Ď����J�n���܂��i8���ԁj...
start "�t�H���_�Ď�" run_watcher.bat
echo �ʃE�B���h�E�ŊĎ����J�n���܂����B
pause
cls
goto MENU

:KINTONE_GET
cls
echo kintone����f�[�^���擾���܂��B
set /p app_name="kintone�A�v��������͂��Ă�������: "
call run_kintone.bat "%app_name%"
pause
cls
goto MENU

:KINTONE_PUSH
cls
echo kintone�Ƀf�[�^�𑗐M���܂��B
set /p app_name="kintone�A�v��������͂��Ă�������: "

set "psCommand="(new-object -COM 'Shell.Application').BrowseForFolder(0,'���M����CSV�t�@�C����I�����Ă�������',0,0).self.path""
for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "folder=%%I"
if "%folder%"=="" goto MENU

set "psCommand="Get-ChildItem -Path '%folder%' -Filter *.csv | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName""
for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "csv_file=%%I"

if "%csv_file%"=="" (
  echo CSV�t�@�C����������܂���ł����B
  pause
  cls
  goto MENU
)

echo �I�����ꂽ�t�@�C��: %csv_file%
echo kintone�A�v��: %app_name%
echo ���M���J�n���܂�...

python main.py run --mode kintone_push --file "%csv_file%" --app_name "%app_name%"

if %errorlevel% neq 0 (
  echo �G���[���������܂����B
) else (
  echo ���M���������܂����B
)

pause
cls
goto MENU

:NOTIFY
cls
echo ��o�󋵊m�F�ƃ��}�C���h���M�����s���܂�...
python notifier.py

if %errorlevel% neq 0 (
  echo �G���[���������܂����B
) else (
  echo �������������܂����B
)

pause
cls
goto MENU

:INFO
cls
echo �ݒ����\�����܂�...
python main.py info
pause
cls
goto MENU

:END
cls
echo �Αӕ\���������c�[�����I�����܂��B
echo �����p���肪�Ƃ��������܂����B
exit /b 0
@echo off
echo === �ΑӊǗ�����^�X�N���s ===
echo ���s����: %date% %time%

REM ���O�t�@�C���ݒ�
set log_dir=logs
if not exist %log_dir% mkdir %log_dir%

set log_file=%log_dir%\scheduled_task_%date:~0,4%%date:~5,2%%date:~8,2%.log
echo ���s�J�n: %date% %time% > %log_file%

REM 1. ��o�󋵂̊m�F�ƃ��}�C���h���M
echo ��o�󋵊m�F�ƃ��}�C���h���M�����s���܂�... >> %log_file%
python notifier.py >> %log_file% 2>&1
if %errorlevel% neq 0 (
  echo [�G���[] ��o�󋵊m�F�Ɏ��s���܂���: %errorlevel% >> %log_file%
) else (
  echo ��o�󋵊m�F���������܂��� >> %log_file%
)

REM 2. ���������ؓ��̏ꍇ�͏��������s
python -c "from datetime import datetime; from config import init_config, get_deadline_date; config=init_config(); deadline=get_deadline_date(); today=datetime.now().strftime('%%Y-%%m-%%d'); print('YES' if deadline == today else 'NO')" > tmp.txt
set /p is_deadline=<tmp.txt
del tmp.txt

if "%is_deadline%"=="YES" (
  echo ���ؓ��̂��ߍŏI���������s���܂�... >> %log_file%
  
  REM �������t�@�C�����ꊇ����
  for /r input %%f in (*.csv) do (
    echo �t�@�C������: %%f >> %log_file%
    python main.py run --file "%%f" >> %log_file% 2>&1
  )
  
  REM kintone�Ƀf�[�^���A�b�v���[�h�i�K�v�ȏꍇ�j
  if exist "config\.upload_to_kintone" (
    echo kintone�Ƀf�[�^���A�b�v���[�h���܂�... >> %log_file%
    python main.py run --mode kintone_push --file "output\�W�v����.csv" --app_name "�ΑӏW�v" >> %log_file% 2>&1
  )
)

echo ���s�I��: %date% %time% >> %log_file%

REM �G���[������Ε\��
find /c "[�G���[]" %log_file% > nul
if not errorlevel 1 (
  echo �G���[���������܂����B�ڍׂ̓��O�t�@�C�����m�F���Ă�������: %log_file%
  pause
) else (
  echo ����������Ɋ������܂����B
)
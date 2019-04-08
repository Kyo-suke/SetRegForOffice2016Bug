@echo off
rem �萔��`
rem �Ώۃ��W�X�g���l�̏ꏊ��`
set REG_KEY_TARGET=HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters
set REG_VAL_FILEINFOCACHELIFETIME=FileInfoCacheLifetime
set REG_VAL_FILENOTFOUNDCACHELIFETIME=FileNotFoundCacheLifetime
set REG_VAL_DIRECTORYCACHELIFETIME=DirectoryCacheLifetime

rem �Ǘ��Ҍ����Ŏ��s���Ă��Ȃ��ꍇ�͎��s�s��
openfiles > nul
if %ERRORLEVEL% == 1 (
    echo ���̃o�b�`�t�@�C���𗘗p����ɂ͊Ǘ��Ҍ����Ŏ��s����K�v������܂��B
    exit /b 0
)

rem �R�}���h�I�v�V���������̉��
:check_option
if "%1" == "" (
    goto main
)

if "%1" == "/?" (
    goto usage
)
if "%1" == "/restore" (
    goto restore
)
shift
goto check_option

rem ----------------------------------------------------------
rem �g�����̕\��
rem ----------------------------------------------------------
:usage
@echo.
@echo %0 [/?] [/restore �t�@�C���p�X]
@echo.
@echo   Office2016 Access�ɂ����ăl�b�g���[�N��̃t�@�C�������s�����ꍇ��
@echo   �f�[�^�x�[�X���j�����錻�ۂ̈ꎞ�Ώ��Ƃ��āA���W�X�g���ύX���s���܂��B
@echo   �ڂ����͈ȉ��̃X���b�h�̏����m�F���Ă��������B
@echo   https://answers.microsoft.com/en-us/msoffice/forum/all/access-database-is-getting-corrupt-again-and-again/d3fcc0a2-7d35-4a09-9269-c5d93ad0031d?page=15
@echo.
@echo   /restore �t�@�C���p�X
@echo     �o�b�N�A�b�v�t�@�C�����烌�W�X�g���𕜌����܂��B
@echo     �t�@�C���p�X�ɂ͕����ɗ��p����t�@�C�����w�肵�܂��B 
@echo.
exit /b 0

rem ----------------------------------------------------------
rem ���C������
rem ----------------------------------------------------------
:main
setlocal enabledelayedexpansion

rem REG_DWORD
set /a FILEINFOCACHELIFETIME=nul
set /a FILENOTFOUNDCACHELIFETIME=nul
set /a DIRECTORYCACHELIFETIME=nul

set /a SET_VAL_FILEINFOCACHELIFETIME=0x0
set /a SET_VAL_FILENOTFOUNDCACHELIFETIME=0x0
set /a SET_VAL_DIRECTORYCACHELIFETIME=0x0

rem ���W�X�g���l�������ꍇ��reg query���G���[��f���ׁA�\�ߊm�F
rem ���W�X�g���l��reg query�𗘗p���Ď擾
@reg query %REG_KEY_TARGET% /v %REG_VAL_FILEINFOCACHELIFETIME% > nul 2>&1
if %ERRORLEVEL% == 0 (
    for /f "tokens=1,2*" %%i in ('reg query %REG_KEY_TARGET% /v %REG_VAL_FILEINFOCACHELIFETIME%') do (
        set FILEINFOCACHELIFETIME=%%k
    )
)

@reg query %REG_KEY_TARGET% /v %REG_VAL_FILENOTFOUNDCACHELIFETIME% > nul 2>&1
if %ERRORLEVEL% == 0 (
    for /f "tokens=1,2*" %%i in ('reg query %REG_KEY_TARGET% /v %REG_VAL_FILENOTFOUNDCACHELIFETIME%') do (
        set FILENOTFOUNDCACHELIFETIME=%%k
    )
)

@reg query %REG_KEY_TARGET% /v %REG_VAL_DIRECTORYCACHELIFETIME% > nul 2>&1
if %ERRORLEVEL% == 0 (
    for /f "tokens=1,2*" %%i in ('reg query %REG_KEY_TARGET% /v %REG_VAL_DIRECTORYCACHELIFETIME%') do (
        set DIRECTORYCACHELIFETIME=%%k
    )
)

rem �擾�������W�X�g���l��\��
@echo ===============================================================
@echo registory key: 
@echo   %REG_KEY_TARGET%
@echo registory value: 
if %FILEINFOCACHELIFETIME% == 0 (
    @echo   %REG_VAL_FILEINFOCACHELIFETIME%     = ^(not set^) ^=^> 0x%SET_VAL_FILEINFOCACHELIFETIME%
) else (
    @echo   %REG_VAL_FILEINFOCACHELIFETIME%     = %FILEINFOCACHELIFETIME% ^=^> 0x%SET_VAL_FILEINFOCACHELIFETIME%
)
if %FILENOTFOUNDCACHELIFETIME% == 0 (
    @echo   %REG_VAL_FILENOTFOUNDCACHELIFETIME% = ^(not set^) ^=^> 0x%SET_VAL_FILENOTFOUNDCACHELIFETIME%
) else (
    @echo   %REG_VAL_FILENOTFOUNDCACHELIFETIME% = %FILENOTFOUNDCACHELIFETIME% ^=^> 0x%SET_VAL_FILENOTFOUNDCACHELIFETIME%
)
if %DIRECTORYCACHELIFETIME% == 0 (
    @echo   %REG_VAL_DIRECTORYCACHELIFETIME%    = ^(not set^) ^=^> 0x%SET_VAL_DIRECTORYCACHELIFETIME%
) else (
    @echo   %REG_VAL_DIRECTORYCACHELIFETIME%    = %DIRECTORYCACHELIFETIME% ^=^> 0x%SET_VAL_DIRECTORYCACHELIFETIME%
)
@echo ===============================================================

rem �ύX�̕K�v�������ꍇ�͉������Ȃ�
if %FILEINFOCACHELIFETIME% == 0x0 IF %FILENOTFOUNDCACHELIFETIME% == 0x0 IF %DIRECTORYCACHELIFETIME% == 0x0 (
    @echo ���W�X�g���̕ύX���s���K�v������܂���ł����B
    exit /b 0
)

@echo ���W�X�g���l��ύX���܂��B
set /P USER_INPUT=��낵���ł���^? [Yes/No]
if not %USER_INPUT%==Yes (
    @echo �����𒆒f���܂��B
    exit /b 0
)

rem registory backup
@echo ���W�X�g���ݒ�̃o�b�N�A�b�v���s���܂��B
rem ���W�X�g���o�b�N�A�b�v�t�@�C��
set REG_BUCKUP_FILE=reg_buckup_%date:~0,4%%date:~5,2%%date:~8,2%.hiv
reg save %REG_KEY_TARGET% %REG_BUCKUP_FILE%

rem set registory
reg add %REG_KEY_TARGET% /v %REG_VAL_FILEINFOCACHELIFETIME% /t REG_DWORD /d 0
if not %ERRORLEVEL% == 0 (
    @echo ���W�X�g��^(%REG_VAL_FILEINFOCACHELIFETIME%^)�̕ύX�Ɏ��s���܂����B
    exit /b 1
)
reg add %REG_KEY_TARGET% /v %REG_VAL_FILENOTFOUNDCACHELIFETIME% /t REG_DWORD /d 0
if not %ERRORLEVEL% == 0 (
    @echo ���W�X�g��^(%REG_VAL_FILENOTFOUNDCACHELIFETIME%^)�̕ύX�Ɏ��s���܂����B
    exit /b 1
)
reg add %REG_KEY_TARGET% /v %REG_VAL_DIRECTORYCACHELIFETIME% /t REG_DWORD /d 0
if not %ERRORLEVEL% == 0 (
    @echo ���W�X�g��^(%REG_VAL_DIRECTORYCACHELIFETIME%^)�̕ύX�Ɏ��s���܂����B
    exit /b 1
)

@echo.
@echo ���W�X�g���̕ύX���������܂����B
@echo ���f�ɂ͍ċN�����K�v�ł��B

endlocal
exit /b 0

rem ----------------------------------------------------------
rem restore����
rem ----------------------------------------------------------
:restore
setlocal enabledelayedexpansion

rem restore�ŗ��p����t�@�C���̑��݃`�F�b�N
if "%2" == "" (
    goto usage
    exit /b 1
)

set RESTORE_FILE_PATH=%2
if not exist %RESTORE_FILE_PATH% (
    @echo �o�b�N�A�b�v�t�@�C�������݂��܂���B^("%2"^)
    exit /b 1
)

@echo �o�b�N�A�b�v�t�@�C���𗘗p���ă��W�X�g����Ԃ𕜌����܂��B
set /P USER_INPUT=��낵���ł���^? [Yes/No]
if not %USER_INPUT%==Yes (
    @echo �����𒆒f���܂��B
    exit /b 0
)

rem ���W�X�g���̕���
reg restore %REG_KEY_TARGET% %RESTORE_FILE_PATH%
if not %ERRORLEVEL% == 0 (
    @echo ���W�X�g���̕����Ɏ��s���܂����B
    exit /b 1
)

@echo.
@echo ���W�X�g���̕������������܂����B
@echo ���f�ɂ͍ċN�����K�v�ł��B

endlocal
exit /b 0

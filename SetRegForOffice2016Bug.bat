@echo off
rem �Ǘ��Ҍ����Ŏ��s���Ă��Ȃ��ꍇ�͎��s�s��
openfiles > nul
if %ERRORLEVEL% == 1 (
    echo ���̃o�b�`�t�@�C���𗘗p����ɂ͊Ǘ��Ҍ����Ŏ��s����K�v������܂��B
    exit /b 0
)

rem �I�v�V���������̉��
for %%f in (%*) do (
    rem usage
    if "%%f" == "/h" (
        goto usage
    )
)
goto main

rem ----------------------------------------------------------
rem �g�����̕\��
rem ----------------------------------------------------------
:usage
@echo %0 [/h]
@echo   Office2016 Access�ɂ����ăl�b�g���[�N��̃t�@�C�������s�����ꍇ��
@echo   �f�[�^�x�[�X���j�����錻�ۂ̈ꎞ�Ώ��Ƃ��āA���W�X�g���ύX���s���܂��B
@echo   �ڂ����͈ȉ��̃X���b�h�̏����m�F���Ă��������B
@echo   https://answers.microsoft.com/en-us/msoffice/forum/all/access-database-is-getting-corrupt-again-and-again/d3fcc0a2-7d35-4a09-9269-c5d93ad0031d?page=15
@echo   /h    ���̎g������\�����܂��B
exit /b 0

rem ----------------------------------------------------------
rem ���C������
rem ----------------------------------------------------------
:main
setlocal enabledelayedexpansion

rem �Ώۃ��W�X�g���l�̏ꏊ��`
set REG_KEY_TARGET=HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters
set REG_VAL_FILEINFOCACHELIFETIME=FileInfoCacheLifetime
set REG_VAL_FILENOTFOUNDCACHELIFETIME=FileNotFoundCacheLifetime
set REG_VAL_DIRECTORYCACHELIFETIME=DirectoryCacheLifetime

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
    @echo   %REG_VAL_FILEINFOCACHELIFETIME%     = ^(not set^) ^=^> %SET_VAL_FILEINFOCACHELIFETIME%
) else (
    @echo   %REG_VAL_FILEINFOCACHELIFETIME%     = %FILEINFOCACHELIFETIME% ^=^> %SET_VAL_FILEINFOCACHELIFETIME%
)
if %FILENOTFOUNDCACHELIFETIME% == 0 (
    @echo   %REG_VAL_FILENOTFOUNDCACHELIFETIME% = ^(not set^) ^=^> %SET_VAL_FILENOTFOUNDCACHELIFETIME%
) else (
    @echo   %REG_VAL_FILENOTFOUNDCACHELIFETIME% = %FILENOTFOUNDCACHELIFETIME% ^=^> %SET_VAL_FILENOTFOUNDCACHELIFETIME%
)
if %DIRECTORYCACHELIFETIME% == 0 (
    @echo   %REG_VAL_DIRECTORYCACHELIFETIME%    = ^(not set^) ^=^> %SET_VAL_DIRECTORYCACHELIFETIME%
) else (
    @echo   %REG_VAL_DIRECTORYCACHELIFETIME%    = %DIRECTORYCACHELIFETIME% ^=^> %SET_VAL_DIRECTORYCACHELIFETIME%
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
set REG_BUCKUP_FILE=reg_buckup_%date:~0,4%%date:~5,2%%date:~8,2%
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

@echo ���W�X�g���̕ύX���������܂����B
@echo ���f�ɂ͍ċN�����K�v�ł��B

endlocal
exit /b 0

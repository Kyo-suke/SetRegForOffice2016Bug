@echo off
rem �萔��`
rem �Ώۃ��W�X�g���l�̏ꏊ��`
set REG_KEY_LANMANWORKSTATION=HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters
set REG_VAL_FILEINFOCACHELIFETIME=FileInfoCacheLifetime
set REG_VAL_FILENOTFOUNDCACHELIFETIME=FileNotFoundCacheLifetime
set REG_VAL_DIRECTORYCACHELIFETIME=DirectoryCacheLifetime

set REG_KEY_LANMANSERVER=HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters
set REG_VAL_DISABLELEASING=DisableLeasing

rem �o�b�N�A�b�v�t�@�C������
set REG_BACKUP_FILENAME_LANMANWORKSTATION=reg_backup_lanmanworkstation_parameter_
set REG_BACKUP_FILENAME_LANMANSERVER=reg_backup_lanmanserver_parameter_

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
@echo %0 [/?] [/restore �t�@�C���p�X...]
@echo.
@echo   Office2016 Access�ɂ����ăl�b�g���[�N��̃t�@�C�������s�����ꍇ��
@echo   �f�[�^�x�[�X���j�����錻�ۂ̈ꎞ�Ώ��Ƃ��āA���W�X�g���ύX���s���܂��B
@echo   �ڂ����͈ȉ��̃X���b�h�̏����m�F���Ă��������B
@echo   https://answers.microsoft.com/en-us/msoffice/forum/all/access-database-is-getting-corrupt-again-and-again/d3fcc0a2-7d35-4a09-9269-c5d93ad0031d?page=15
@echo.
@echo   ������̋L���̓��e�ɂ��Ă��A�Ή����܂��B
@echo   https://support.office.com/ja-jp/article/access-%E3%81%A7%E3%83%87%E3%83%BC%E3%82%BF%E3%83%99%E3%83%BC%E3%82%B9%E3%81%8C-%E7%9F%9B%E7%9B%BE%E3%81%8C%E3%81%82%E3%82%8B%E7%8A%B6%E6%85%8B-%E3%81%AB%E3%81%82%E3%82%8B%E3%81%A8%E5%A0%B1%E5%91%8A%E3%81%95%E3%82%8C%E3%82%8B-7ec975da-f7a9-4414-a306-d3a7c422dc1d
@echo.
@echo   /restore �t�@�C���p�X...
@echo     �o�b�N�A�b�v�t�@�C�����烌�W�X�g���𕜌����܂��B
@echo     �t�@�C���p�X�ɂ͕����ɗ��p����t�@�C�����w�肵�܂��B 
@echo.
exit /b 0

rem ----------------------------------------------------------
rem ���C������
rem ----------------------------------------------------------
:main
setlocal enabledelayedexpansion

@echo �{�o�b�`�����ł̓��W�X�g���̕ύX���s���܂��B
@echo �V�X�e���ɉe��������\��������A���p�͎��ȐӔC�ƂȂ�܂��B
@echo �����|�C���g�̍쐬���𗘗p���āA���S����S�ۂ��Ă��������B
set /P USER_INPUT=��낵���ł���^? [Yes/No] ^> 
if not %USER_INPUT%==Yes (
    @echo �����𒆒f���܂��B
    exit /b 0
)

rem REG_DWORD
set /a FILEINFOCACHELIFETIME=-1
set /a FILENOTFOUNDCACHELIFETIME=-1
set /a DIRECTORYCACHELIFETIME=-1
set /a DISABLELEASING=-1

set /a SET_VAL_FILEINFOCACHELIFETIME=0x0
set /a SET_VAL_FILENOTFOUNDCACHELIFETIME=0x0
set /a SET_VAL_DIRECTORYCACHELIFETIME=0x0
set /a SET_VAL_DISABLELEASING=0x1

rem ���W�X�g���l�������ꍇ��reg query���G���[��f���ׁA�\�ߊm�F
rem ���W�X�g���l��reg query�𗘗p���Ď擾
@reg query %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_FILEINFOCACHELIFETIME% > nul 2>&1 && (
    for /f "tokens=3 skip=1" %%i in ('reg query %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_FILEINFOCACHELIFETIME%') do (
        set /a FILEINFOCACHELIFETIME=%%i
    )
)

@reg query %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_FILENOTFOUNDCACHELIFETIME% > nul 2>&1 && (
    for /f "tokens=3 skip=1" %%i in ('reg query %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_FILENOTFOUNDCACHELIFETIME%') do (
        set /a FILENOTFOUNDCACHELIFETIME=%%i
    )
)

@reg query %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_DIRECTORYCACHELIFETIME% > nul 2>&1 && (
    for /f "tokens=3 skip=1" %%i in ('reg query %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_DIRECTORYCACHELIFETIME%') do (
        set /a DIRECTORYCACHELIFETIME=%%i
    )
)

@reg query %REG_KEY_LANMANSERVER% /v %REG_VAL_DISABLELEASING% > nul 2>&1 && (
    for /f "tokens=3 skip=1" %%i in ('reg query %REG_KEY_LANMANSERVER% /v %REG_VAL_DISABLELEASING%') do (
        set /a DISABLELEASING=%%i
    )
)

rem �擾�������W�X�g���l��\��
@echo ===============================================================
@echo registory key: 
@echo   %REG_KEY_LANMANWORKSTATION%
@echo registory value: 
if %FILEINFOCACHELIFETIME% == -1 (
    @echo   %REG_VAL_FILEINFOCACHELIFETIME%     = ^(not set^) ^=^> 0x%SET_VAL_FILEINFOCACHELIFETIME%
) else (
    @echo   %REG_VAL_FILEINFOCACHELIFETIME%     = 0x%FILEINFOCACHELIFETIME% ^=^> 0x%SET_VAL_FILEINFOCACHELIFETIME%
)
if %FILENOTFOUNDCACHELIFETIME% == -1 (
    @echo   %REG_VAL_FILENOTFOUNDCACHELIFETIME% = ^(not set^) ^=^> 0x%SET_VAL_FILENOTFOUNDCACHELIFETIME%
) else (
    @echo   %REG_VAL_FILENOTFOUNDCACHELIFETIME% = 0x%FILENOTFOUNDCACHELIFETIME% ^=^> 0x%SET_VAL_FILENOTFOUNDCACHELIFETIME%
)
if %DIRECTORYCACHELIFETIME% == -1 (
    @echo   %REG_VAL_DIRECTORYCACHELIFETIME%    = ^(not set^) ^=^> 0x%SET_VAL_DIRECTORYCACHELIFETIME%
) else (
    @echo   %REG_VAL_DIRECTORYCACHELIFETIME%    = 0x%DIRECTORYCACHELIFETIME% ^=^> 0x%SET_VAL_DIRECTORYCACHELIFETIME%
)
@echo.
@echo registory key:
@echo   %REG_KEY_LANMANSERVER%
@echo registory value:
if %DISABLELEASING% == -1 (
    @echo   %REG_VAL_DISABLELEASING%            = ^(not set^) ^=^> 0x%SET_VAL_DISABLELEASING%
) else (
    @echo   %REG_VAL_DISABLELEASING%            = 0x%DISABLELEASING% ^=^> 0x%SET_VAL_DISABLELEASING%
)
@echo ===============================================================

rem �ύX�̕K�v�������ꍇ�͉������Ȃ�
set /a IS_CHANGE_NECESSARY=0
if not %FILEINFOCACHELIFETIME% == %SET_VAL_FILEINFOCACHELIFETIME% ( set /a IS_CHANGE_NECESSARY=1 )
if not %FILENOTFOUNDCACHELIFETIME% == %SET_VAL_FILENOTFOUNDCACHELIFETIME% ( set /a IS_CHANGE_NECESSARY=1 )
if not %DIRECTORYCACHELIFETIME% == %SET_VAL_DIRECTORYCACHELIFETIME% ( set /a IS_CHANGE_NECESSARY=1 )
if not %DISABLELEASING% == %SET_VAL_DISABLELEASING% ( set /a IS_CHANGE_NECESSARY=1 )
if %IS_CHANGE_NECESSARY% == 0 (
    @echo ���W�X�g���̕ύX���s���K�v������܂���ł����B
    exit /b 0
)

@echo ���W�X�g���l��ύX���܂��B
set /P USER_INPUT=��낵���ł���^? [Yes/No] ^> 
if not %USER_INPUT%==Yes (
    @echo �����𒆒f���܂��B
    exit /b 0
)

rem registory backup
@echo ���W�X�g���ݒ�̃o�b�N�A�b�v���s���܂��B
rem ���W�X�g���o�b�N�A�b�v�t�@�C��
set REG_BACKUP_FILE_LANMANWORKSTATION=%REG_BACKUP_FILENAME_LANMANWORKSTATION%%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%.hiv
reg save %REG_KEY_LANMANWORKSTATION% %REG_BACKUP_FILE_LANMANWORKSTATION% || (
    @echo ���W�X�g���̃o�b�N�A�b�v�Ɏ��s���܂����B^(%REG_KEY_LANMANWORKSTATION%^)
    exit /b 1
)
set REG_BACKUP_FILE_LANMANSERVER=%REG_BACKUP_FILENAME_LANMANSERVER%%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%.hiv
reg save %REG_KEY_LANMANSERVER% %REG_BACKUP_FILE_LANMANSERVER% || (
    @echo ���W�X�g���̃o�b�N�A�b�v�Ɏ��s���܂����B^(%REG_KEY_LANMANSERVER%^)
    exit /b 1
)

rem set registory
reg add %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_FILEINFOCACHELIFETIME% /t REG_DWORD /d %SET_VAL_FILEINFOCACHELIFETIME% || (
    @echo ���W�X�g��^(%REG_VAL_FILEINFOCACHELIFETIME%^)�̕ύX�Ɏ��s���܂����B
    exit /b 1
)
reg add %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_FILENOTFOUNDCACHELIFETIME% /t REG_DWORD /d %SET_VAL_FILENOTFOUNDCACHELIFETIME% || (
    @echo ���W�X�g��^(%REG_VAL_FILENOTFOUNDCACHELIFETIME%^)�̕ύX�Ɏ��s���܂����B
    exit /b 1
)
reg add %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_DIRECTORYCACHELIFETIME% /t REG_DWORD /d %SET_VAL_DIRECTORYCACHELIFETIME% || (
    @echo ���W�X�g��^(%REG_VAL_DIRECTORYCACHELIFETIME%^)�̕ύX�Ɏ��s���܂����B
    exit /b 1
)
reg add %REG_KEY_LANMANSERVER% /v %REG_VAL_DISABLELEASING% /t REG_DWORD /d %SET_VAL_DISABLELEASING% || (
    @echo ���W�X�g��^(%REG_VAL_DISABLELEASING%^)�̕ύX�Ɏ��s���܂����B
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

rem �^�[�Q�b�g�t�@�C�����w�肳��Ă��Ȃ��ꍇ��NG
if "%2" == "" (
    goto usage
    exit /b 1
)

for %%f in (%*) do (
    rem �R�}���h������ǂݔ�΂�
    if not %%f == ^/restore (
        rem �o�b�N�A�b�v�t�@�C���̑��݃`�F�b�N
        if not exist %%f (
            @echo �o�b�N�A�b�v�t�@�C�������݂��܂���B^("%%f"^)
            exit /b 1
        )
        rem �o�b�N�A�b�v�t�@�C�����̃`�F�b�N
        @echo %%f | findstr /r /b "%REG_BACKUP_FILENAME_LANMANWORKSTATION%" > nul 2>&1 || (
            @echo %%f | findstr /b "%REG_BACKUP_FILENAME_LANMANSERVER%" > nul 2>&1 || (
                @echo �o�b�N�A�b�v�t�@�C�������s���ł��B^("%%f"^)
                exit /b 1
            )
        )
    )
)

@echo �o�b�N�A�b�v�t�@�C���𗘗p���ă��W�X�g����Ԃ𕜌����܂��B
set /P USER_INPUT=��낵���ł���^? [Yes/No] ^> 
if not %USER_INPUT%==Yes (
    @echo �����𒆒f���܂��B
    exit /b 0
)

rem ���W�X�g���̕���
for %%f in (%*) do (
    if not %%f == ^/restore (
        rem �t�@�C�����ŕ����L�[������ӂ�
        @echo %%f | findstr /r /b "%REG_BACKUP_FILENAME_LANMANWORKSTATION%" > nul 2>&1 && (
            @echo restore %REG_KEY_LANMANWORKSTATION%
            reg restore %REG_KEY_LANMANWORKSTATION% %%f
            if not %ERRORLEVEL% == 0 (
                @echo ���W�X�g���̕����Ɏ��s���܂����B
                exit /b 1
            )
        )
        @echo %%f | findstr /r /b "%REG_BACKUP_FILENAME_LANMANSERVER%" > nul 2>&1 && (
            @echo restore %REG_KEY_LANMANSERVER%
            reg restore %REG_KEY_LANMANSERVER% %%f
            if not %ERRORLEVEL% == 0 (
                @echo ���W�X�g���̕����Ɏ��s���܂����B
                exit /b 1
            )
        )
    )
)

@echo.
@echo ���W�X�g���̕������������܂����B
@echo ���f�ɂ͍ċN�����K�v�ł��B

endlocal
exit /b 0

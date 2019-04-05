@echo off
rem 管理者権限で実行していない場合は実行不可
openfiles > nul
if %ERRORLEVEL% == 1 (
    echo このバッチファイルを利用するには管理者権限で実行する必要があります。
    exit /b 0
)

rem オプション引数の解析
for %%f in (%*) do (
    rem usage
    if "%%f" == "/h" (
        goto usage
    )
)
goto main

rem ----------------------------------------------------------
rem 使い方の表示
rem ----------------------------------------------------------
:usage
@echo %0 [/h]
@echo   Office2016 Accessにおいてネットワーク上のファイルを実行した場合に
@echo   データベースが破損する現象の一時対処として、レジストリ変更を行います。
@echo   詳しくは以下のスレッドの情報を確認してください。
@echo   https://answers.microsoft.com/en-us/msoffice/forum/all/access-database-is-getting-corrupt-again-and-again/d3fcc0a2-7d35-4a09-9269-c5d93ad0031d?page=15
@echo   /h    この使い方を表示します。
exit /b 0

rem ----------------------------------------------------------
rem メイン処理
rem ----------------------------------------------------------
:main
setlocal enabledelayedexpansion

rem 対象レジストリ値の場所定義
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

rem レジストリ値が無い場合はreg queryがエラーを吐く為、予め確認
rem レジストリ値をreg queryを利用して取得
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

rem 取得したレジストリ値を表示
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

rem 変更の必要が無い場合は何もしない
if %FILEINFOCACHELIFETIME% == 0x0 IF %FILENOTFOUNDCACHELIFETIME% == 0x0 IF %DIRECTORYCACHELIFETIME% == 0x0 (
    @echo レジストリの変更を行う必要がありませんでした。
    exit /b 0
)

@echo レジストリ値を変更します。
set /P USER_INPUT=よろしいですか^? [Yes/No]
if not %USER_INPUT%==Yes (
    @echo 処理を中断します。
    exit /b 0
)

rem registory backup
@echo レジストリ設定のバックアップを行います。
rem レジストリバックアップファイル
set REG_BUCKUP_FILE=reg_buckup_%date:~0,4%%date:~5,2%%date:~8,2%
reg save %REG_KEY_TARGET% %REG_BUCKUP_FILE%

rem set registory
reg add %REG_KEY_TARGET% /v %REG_VAL_FILEINFOCACHELIFETIME% /t REG_DWORD /d 0
if not %ERRORLEVEL% == 0 (
    @echo レジストリ^(%REG_VAL_FILEINFOCACHELIFETIME%^)の変更に失敗しました。
    exit /b 1
)
reg add %REG_KEY_TARGET% /v %REG_VAL_FILENOTFOUNDCACHELIFETIME% /t REG_DWORD /d 0
if not %ERRORLEVEL% == 0 (
    @echo レジストリ^(%REG_VAL_FILENOTFOUNDCACHELIFETIME%^)の変更に失敗しました。
    exit /b 1
)
reg add %REG_KEY_TARGET% /v %REG_VAL_DIRECTORYCACHELIFETIME% /t REG_DWORD /d 0
if not %ERRORLEVEL% == 0 (
    @echo レジストリ^(%REG_VAL_DIRECTORYCACHELIFETIME%^)の変更に失敗しました。
    exit /b 1
)

@echo レジストリの変更が完了しました。
@echo 反映には再起動が必要です。

endlocal
exit /b 0

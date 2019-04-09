@echo off
rem 定数定義
rem 対象レジストリ値の場所定義
set REG_KEY_LANMANWORKSTATION=HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters
set REG_VAL_FILEINFOCACHELIFETIME=FileInfoCacheLifetime
set REG_VAL_FILENOTFOUNDCACHELIFETIME=FileNotFoundCacheLifetime
set REG_VAL_DIRECTORYCACHELIFETIME=DirectoryCacheLifetime

set REG_KEY_LANMANSERVER=HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters
set REG_VAL_DISABLELEASING=DisableLeasing

rem バックアップファイル名称
set REG_BACKUP_FILENAME_LANMANWORKSTATION=reg_backup_lanmanworkstation_parameter_
set REG_BACKUP_FILENAME_LANMANSERVER=reg_backup_lanmanserver_parameter_

rem 管理者権限で実行していない場合は実行不可
openfiles > nul
if %ERRORLEVEL% == 1 (
    echo このバッチファイルを利用するには管理者権限で実行する必要があります。
    exit /b 0
)

rem コマンドオプション引数の解析
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
rem 使い方の表示
rem ----------------------------------------------------------
:usage
@echo.
@echo %0 [/?] [/restore ファイルパス...]
@echo.
@echo   Office2016 Accessにおいてネットワーク上のファイルを実行した場合に
@echo   データベースが破損する現象の一時対処として、レジストリ変更を行います。
@echo   詳しくは以下のスレッドの情報を確認してください。
@echo   https://answers.microsoft.com/en-us/msoffice/forum/all/access-database-is-getting-corrupt-again-and-again/d3fcc0a2-7d35-4a09-9269-c5d93ad0031d?page=15
@echo.
@echo   こちらの記事の内容についても、対応します。
@echo   https://support.office.com/ja-jp/article/access-%E3%81%A7%E3%83%87%E3%83%BC%E3%82%BF%E3%83%99%E3%83%BC%E3%82%B9%E3%81%8C-%E7%9F%9B%E7%9B%BE%E3%81%8C%E3%81%82%E3%82%8B%E7%8A%B6%E6%85%8B-%E3%81%AB%E3%81%82%E3%82%8B%E3%81%A8%E5%A0%B1%E5%91%8A%E3%81%95%E3%82%8C%E3%82%8B-7ec975da-f7a9-4414-a306-d3a7c422dc1d
@echo.
@echo   /restore ファイルパス...
@echo     バックアップファイルからレジストリを復元します。
@echo     ファイルパスには復元に利用するファイルを指定します。 
@echo.
exit /b 0

rem ----------------------------------------------------------
rem メイン処理
rem ----------------------------------------------------------
:main
setlocal enabledelayedexpansion

@echo 本バッチ処理ではレジストリの変更を行います。
@echo システムに影響がある可能性があり、利用は自己責任となります。
@echo 復元ポイントの作成等を利用して、安全性を担保してください。
set /P USER_INPUT=よろしいですか^? [Yes/No] ^> 
if not %USER_INPUT%==Yes (
    @echo 処理を中断します。
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

rem レジストリ値が無い場合はreg queryがエラーを吐く為、予め確認
rem レジストリ値をreg queryを利用して取得
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

rem 取得したレジストリ値を表示
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

rem 変更の必要が無い場合は何もしない
set /a IS_CHANGE_NECESSARY=0
if not %FILEINFOCACHELIFETIME% == %SET_VAL_FILEINFOCACHELIFETIME% ( set /a IS_CHANGE_NECESSARY=1 )
if not %FILENOTFOUNDCACHELIFETIME% == %SET_VAL_FILENOTFOUNDCACHELIFETIME% ( set /a IS_CHANGE_NECESSARY=1 )
if not %DIRECTORYCACHELIFETIME% == %SET_VAL_DIRECTORYCACHELIFETIME% ( set /a IS_CHANGE_NECESSARY=1 )
if not %DISABLELEASING% == %SET_VAL_DISABLELEASING% ( set /a IS_CHANGE_NECESSARY=1 )
if %IS_CHANGE_NECESSARY% == 0 (
    @echo レジストリの変更を行う必要がありませんでした。
    exit /b 0
)

@echo レジストリ値を変更します。
set /P USER_INPUT=よろしいですか^? [Yes/No] ^> 
if not %USER_INPUT%==Yes (
    @echo 処理を中断します。
    exit /b 0
)

rem registory backup
@echo レジストリ設定のバックアップを行います。
rem レジストリバックアップファイル
set REG_BACKUP_FILE_LANMANWORKSTATION=%REG_BACKUP_FILENAME_LANMANWORKSTATION%%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%.hiv
reg save %REG_KEY_LANMANWORKSTATION% %REG_BACKUP_FILE_LANMANWORKSTATION% || (
    @echo レジストリのバックアップに失敗しました。^(%REG_KEY_LANMANWORKSTATION%^)
    exit /b 1
)
set REG_BACKUP_FILE_LANMANSERVER=%REG_BACKUP_FILENAME_LANMANSERVER%%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%.hiv
reg save %REG_KEY_LANMANSERVER% %REG_BACKUP_FILE_LANMANSERVER% || (
    @echo レジストリのバックアップに失敗しました。^(%REG_KEY_LANMANSERVER%^)
    exit /b 1
)

rem set registory
reg add %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_FILEINFOCACHELIFETIME% /t REG_DWORD /d %SET_VAL_FILEINFOCACHELIFETIME% || (
    @echo レジストリ^(%REG_VAL_FILEINFOCACHELIFETIME%^)の変更に失敗しました。
    exit /b 1
)
reg add %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_FILENOTFOUNDCACHELIFETIME% /t REG_DWORD /d %SET_VAL_FILENOTFOUNDCACHELIFETIME% || (
    @echo レジストリ^(%REG_VAL_FILENOTFOUNDCACHELIFETIME%^)の変更に失敗しました。
    exit /b 1
)
reg add %REG_KEY_LANMANWORKSTATION% /v %REG_VAL_DIRECTORYCACHELIFETIME% /t REG_DWORD /d %SET_VAL_DIRECTORYCACHELIFETIME% || (
    @echo レジストリ^(%REG_VAL_DIRECTORYCACHELIFETIME%^)の変更に失敗しました。
    exit /b 1
)
reg add %REG_KEY_LANMANSERVER% /v %REG_VAL_DISABLELEASING% /t REG_DWORD /d %SET_VAL_DISABLELEASING% || (
    @echo レジストリ^(%REG_VAL_DISABLELEASING%^)の変更に失敗しました。
)

@echo.
@echo レジストリの変更が完了しました。
@echo 反映には再起動が必要です。

endlocal
exit /b 0

rem ----------------------------------------------------------
rem restore処理
rem ----------------------------------------------------------
:restore
setlocal enabledelayedexpansion

rem ターゲットファイルが指定されていない場合はNG
if "%2" == "" (
    goto usage
    exit /b 1
)

for %%f in (%*) do (
    rem コマンド部分を読み飛ばし
    if not %%f == ^/restore (
        rem バックアップファイルの存在チェック
        if not exist %%f (
            @echo バックアップファイルが存在しません。^("%%f"^)
            exit /b 1
        )
        rem バックアップファイル名のチェック
        @echo %%f | findstr /r /b "%REG_BACKUP_FILENAME_LANMANWORKSTATION%" > nul 2>&1 || (
            @echo %%f | findstr /b "%REG_BACKUP_FILENAME_LANMANSERVER%" > nul 2>&1 || (
                @echo バックアップファイル名が不正です。^("%%f"^)
                exit /b 1
            )
        )
    )
)

@echo バックアップファイルを利用してレジストリ状態を復元します。
set /P USER_INPUT=よろしいですか^? [Yes/No] ^> 
if not %USER_INPUT%==Yes (
    @echo 処理を中断します。
    exit /b 0
)

rem レジストリの復元
for %%f in (%*) do (
    if not %%f == ^/restore (
        rem ファイル名で復元キーを割りふる
        @echo %%f | findstr /r /b "%REG_BACKUP_FILENAME_LANMANWORKSTATION%" > nul 2>&1 && (
            @echo restore %REG_KEY_LANMANWORKSTATION%
            reg restore %REG_KEY_LANMANWORKSTATION% %%f
            if not %ERRORLEVEL% == 0 (
                @echo レジストリの復元に失敗しました。
                exit /b 1
            )
        )
        @echo %%f | findstr /r /b "%REG_BACKUP_FILENAME_LANMANSERVER%" > nul 2>&1 && (
            @echo restore %REG_KEY_LANMANSERVER%
            reg restore %REG_KEY_LANMANSERVER% %%f
            if not %ERRORLEVEL% == 0 (
                @echo レジストリの復元に失敗しました。
                exit /b 1
            )
        )
    )
)

@echo.
@echo レジストリの復元が完了しました。
@echo 反映には再起動が必要です。

endlocal
exit /b 0

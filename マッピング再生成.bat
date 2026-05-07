@echo off
:: コマンドプロンプトの文字コードをUTF-8に設定
chcp 65001 > nul

:: 実行ディレクトリを必ずこのバッチファイルがあるフォルダに固定する
cd /d "%~dp0"

echo ============================================
echo  pool_mapping.json 自動生成・配置ツール
echo ============================================
echo.

:: 入力処理: 引数（ドラッグ＆ドロップ）があればそれを、なければ入力を促す
set "TARGET_DIR=%~1"
if "%TARGET_DIR%"=="" (
    set /p TARGET_DIR="Excelフォルダのパスを入力（またはドラッグ＆ドロップ）: "
)

:: 未入力チェック
if "%TARGET_DIR%"=="" (
    echo [エラー] フォルダパスが入力されていません。
    pause
    exit /b 1
)

:: パスに含まれるダブルクォーテーションを一旦すべて削除し、安全な状態にする
set "TARGET_DIR=%TARGET_DIR:"=%"

echo.
echo 対象フォルダ: "%TARGET_DIR%"
echo.
echo [1/3] pool_mapping.json を生成中...
python setup_mapping.py "%TARGET_DIR%"
if errorlevel 1 (
    echo [エラー] Pythonスクリプトの実行に失敗しました。パスやPythonの環境を確認してください。
    pause
    exit /b 1
)

echo [2/3] dist\ (PoolWriter用) へコピー中...
if not exist "dist" mkdir "dist"
copy /y "pool_mapping.json" "dist\pool_mapping.json" > nul
if errorlevel 1 (
    echo [エラー] dist\ へのコピーに失敗しました。
    pause
    exit /b 1
)

echo [3/3] public\ (スマホアプリ配信用) へコピー中...
if not exist "public" mkdir "public"
copy /y "pool_mapping.json" "public\pool_mapping.json" > nul
if errorlevel 1 (
    echo [エラー] public\ へのコピーに失敗しました。
    pause
    exit /b 1
)

echo.
echo ============================================
echo  ✓ すべての処理が正常に完了しました
echo ============================================
echo   - 元データ: pool_mapping.json
echo   - PC書込用: dist\pool_mapping.json
echo   - スマホ用: public\pool_mapping.json
echo.
echo 次のステップ: 「スマホアプリ更新.bat」を実行してFirebaseへ反映してください。
echo.
pause
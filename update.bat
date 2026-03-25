@echo off
cd /d "%~dp0"

echo === Updating standings ===
uv run python main.py
if errorlevel 1 (
    echo ERROR: Script failed. Aborting push.
    pause
    exit /b 1
)

echo === Pushing to GitHub ===
git add index.html
git diff --cached --quiet && (
    echo No changes to push.
    pause
    exit /b 0
)
git commit -m "Update standings %date:~10,4%-%date:~4,2%-%date:~7,2%"
git push

echo === Done! Live at https://hasan-burak-uzun.github.io/dynasty-league/ ===
pause

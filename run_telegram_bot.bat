@echo off
chcp 65001 > nul
title Team CVS BHX Sales Tracker - Telegram Bot
cd /d "%~dp0"

echo ========================================================
echo   KHOI DONG TELEGRAM BOT THEO DOI DOANH SO TEAM
echo ========================================================
echo.

python telegram_bot.py

pause

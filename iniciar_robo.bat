@echo off
title Robô Jardim Equipamentos - AUTO RESTART
:inicio
echo ---------------------------------------------------
echo [%date% %time%] Iniciando Robô Jardim Equipamentos...
echo ---------------------------------------------------
python telegram_bot.py
echo.
echo ---------------------------------------------------
echo [AVISO] O robô parou ou caiu! 
echo Tentando reiniciar automaticamente em 5 segundos...
echo ---------------------------------------------------
timeout /t 5
goto inicio

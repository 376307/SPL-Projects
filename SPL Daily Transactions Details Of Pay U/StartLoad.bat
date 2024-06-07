@echo off
sqlldr 'RPA_RO/Robot#123@MAFPRD' control='Control.txt' log='Results.log' direct='true'
pause
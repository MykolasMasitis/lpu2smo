*-- Общие константы приложения
*-- Данный файл должен быть включен в ЛЮБОЙ класс приложения

#INCLUDE 	STRINGS.H
#INCLUDE	KEYBOARD.H
#INCLUDE	VB_CONSTANT.H

# DEFINE DEBUGMODE .F.

*-- Задержка вывода WAIT
#DEFINE		WAIT_TIMEOUT		1

#DEFINE CRLF  CHR(13) + CHR(10)
#DEFINE CR    CHR(13)
#DEFINE TAB    CHR(9)

*-- Эти константы используются в формах 
*-- для индикации статуса текущего алиаса
#DEFINE FILE_OK    0
#DEFINE FILE_BOF  1
#DEFINE FILE_EOF  2
#DEFINE FILE_CANCEL  3



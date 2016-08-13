PROCEDURE ImpBankCLient

 IF MESSAGEBOX(CHR(13)+CHR(10)+;
  'ВЫ ХОТИТЕ ЗАГРУЗИТЬ ПЛАТЕЖИ?'+CHR(13)+CHR(10),;
   4+32,'ИМПОРТ ИЗ БАНК-КЛИЕНТ')==7
  RETURN 
 ENDIF 

 odir = SYS(5)+SYS(2003)
 SET DEFAULT TO (pBin)
 BCFile = GETFILE('Text:txt','','',0,'Укажите на файл!')
 SET DEFAULT TO (odir)
 
 oBCFile = fso.GetFile(BCFile)
 IF oBCFile.size >= 20
  fhandl = oBCFile.OpenAsTextStream
  lcHead = fhandl.Read(20)
  fhandl.Close
 ENDIF 
 
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')

 IF lcHead == '1CClientBankExchange'
  MESSAGEBOX('НАЧАЛО ОБРАБОТКИ!',0+64,'')
 ELSE 
  MESSAGEBOX('НЕИЗВЕСТНЫЙ ФОРМАТ ФАЙЛА!',0+64,'')
 ENDIF 

 fhandl = oBCFile.OpenAsTextStream

 nPayOrdersAll = 0
 nPayOrdersToLpu = 0

 DO WHILE !fhandl.AtEndOfStream
  cLine = fhandl.ReadLine
  IF cLine = 'СекцияДокумент=Платежное поручение'
   nPayOrdersAll = nPayOrdersAll + 1
  ENDIF 
  IF cLine = 'НазначениеПлатежа=' AND SEEK(SUBSTR(cLine, AT('=', cLine)+1, 7), 'sprlpu')
   nPayOrdersToLpu = nPayOrdersToLpu + 1
  ENDIF 
 ENDDO 
 fhandl.Close
 
 USE IN sprlpu
 
 MESSAGEBOX('ДОКУМЕНТ СОДЕРЖИТ '+STR(nPayOrdersAll,3)+ ' ПЛАТЕЖЕК, '+CHR(13)+CHR(10)+;
  'ИЗ НИХ В АДРЕС ЛПУ '+STR(nPayOrdersToLpu,3), 0+64, '')
  
RETURN 
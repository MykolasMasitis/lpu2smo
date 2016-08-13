PROCEDURE Formn1
 IF MESSAGEBOX('ЗАПОЛНИТЬ ФАЙЛ ФОРМЫ №1?',4+32,'')==7
  RETURN 
 ENDIF 

 OutDirPeriod = pOut + '\' + gcPeriod

 FN1File = 'f'+ALLTRIM(STR(m.qObjId))
 IF !fso.FileExists(pout+'\'+gcPeriod+'\'+FN1File+'.dbf')
  MESSAGEBOX('ФАЙЛ '+FN1File+' НЕ НАЙДЕН!',0+16,'')
  RETURN 
 ENDIF 
 
 FinFile = 'f'+m.qcod
 IF !fso.FileExists(pout+'\'+gcPeriod+'\'+FinFile+'.dbf')
  MESSAGEBOX('ФАЙЛ '+FinFile+' НЕ НАЙДЕН!',0+16,'')
  RETURN 
 ENDIF 

 IF tMonth > 1
  FOR nMonth = tMonth-1 TO 1 STEP -1 
   lcPeriod = STR(tYear,4) + PADL(nMonth,2,'0')
   IF !fso.FileExists(pout+'\'+lcPeriod+'\'+FN1File+'.dbf')
    MESSAGEBOX('ФАЙЛ '+FN1File+' ЗА '+NameOfMonth(nMonth)+' НЕ НАЙДЕН!',0+16,'')
   ENDIF 
  NEXT 
 ENDIF 

 IF OpenFile(pout+'\'+gcPeriod+'\'+FN1File, 'fnfile', 'excl')>0
  RETURN 
 ENDIF 
 
 IF OpenFile(pout+'\'+gcPeriod+'\'+FinFile, 'finfile', 'excl')>0
  USE IN fnfile
  RETURN 
 ENDIF 
 
 SELECT FinFile
 INDEX ON lpu_id TAG lpuid
 SET ORDER TO lpuid

 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi' + '\sprlpuxx', 'sprlpu', 'shared', 'lpu_id')
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi' + '\tarifn', 'tarif', 'shared', 'cod')
 
 SELECT fnfile
 SET RELATION TO mo_code INTO FinFile
 
 m.gnDelta = 0
 
 SCAN 
  m.lpuid = mo_code
  m.mcod = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
  IF !fso.FileExists(OutDirPeriod+'\l'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(OutDirPeriod+'\l'+m.mcod, 'lfile', 'shared')>0
   LOOP 
  ENDIF 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  m.r07_c3 = FinFile.s_pred-FinFile.s_pr_opl-FinFile.s_dop2-FinFile.s_sank && не забыть вычесть МЭЭ и ЭКМП!
  
  m.r11_c8 = 0
  m.r12_c8 = 0
  m.r13_c8 = 0
  m.r16_c8 = 0
  m.r18_c8 = 0
  m.r19_c8 = 0
  
  SELECT lfile
  SCAN
   m.cod = cod 
   m.pr_all = pr_all
   m.vmp = IIF(SEEK(m.cod, 'tarif'), tarif.vmp, 0)
   DO CASE 
    CASE m.vmp == 1 && первичная медико-санитарная амбулаторная помощь
     m.r11_c8 = m.r11_c8 + m.pr_all
    CASE m.vmp == 2 && первичная медико-санитарная стоматологическая помощь
     m.r12_c8 = m.r12_c8 + m.pr_all
    CASE m.vmp == 3 && первичная медико-санитарная амбулаторная помощь, оказанная в условиях дневных стационаров всех типов         
     m.r13_c8 = m.r13_c8 + m.pr_all
    CASE m.vmp == 4 && Специализированная  медицинская амбулаторная  помощь 
     m.r16_c8 = m.r16_c8 + m.pr_all
    CASE m.vmp == 5 && Специализированная  медицинская помощь, оказанная в условиях дневных стационаров всех типов         
     m.r18_c8 = m.r18_c8 + m.pr_all
    CASE m.vmp == 6 && Специализированная  медицинская стационарная  помощь              
     m.r19_c8 = m.r19_c8 + m.pr_all
    OTHERWISE 
   ENDCASE 
  ENDSCAN 
  USE 

  SELECT fnfile 
  
  m.r10_c8 = m.r11_c8 + m.r12_c8 + m.r13_c8
  
  m.r15_c8 = m.r16_c8 + m.r18_c8 + m.r19_c8
  
  REPLACE r07_c3 WITH m.r07_c3, ;
          r10_c8 WITH m.r10_c8, R11_C8 WITH m.r11_c8, R12_C8 WITH m.r12_c8, R13_C8 WITH m.r13_c8, ;
          r15_c8 WITH m.r15_c8, R16_C8 WITH m.r16_c8, R18_C8 WITH m.r18_c8, R19_C8 WITH m.r19_c8 
   
 ENDSCAN 
 
 WAIT CLEAR 
 
 USE 
 USE IN tarif 
 USE IN sprlpu
 
 SELECT FinFile
 SET ORDER TO 
 DELETE TAG ALL 
 USE 
 
  MESSAGEBOX('РАСЧЕТ ЗАКОНЧЕН!',0+64,'')
 
RETURN 
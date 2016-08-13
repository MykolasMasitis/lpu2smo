PROCEDURE FormPGMEKS6

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ'+CHR(13)+CHR(10)+;
  'ФОРМУ ПГ ПО МЭК?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 

 m.dotname = ptempl+'\pgmek.xls'
 IF !fso.FileExists(m.dotname)
  MESSAGEBOX('ОТСУТСТВУЕТ ШАБЛОН'+CHR(13)+CHR(10)+m.dotname,0+64,'')
  RETURN 
 ENDIF 

 m.pgdat1 = m.tdat1
 m.pgdat2 = m.tdat2
 m.ischecked = .f.
 DO FORM SelPeriod
 IF m.ischecked = .f.
  RETURN 
 ENDIF 

 m.docname = pMee+'\PgMEK_'+DTOS(m.pgdat1)+'_'+DTOS(m.pgdat2)
 IF fso.FileExists(m.docname+'.xls')
  fso.DeleteFile(m.docname+'.xls')
 ENDIF 

 DIMENSION dimdata(15,8)
 dimdata=0

 m.curdat = m.pgdat1-1
 m.curmonth = 0
 DO WHILE  m.curdat<m.pgdat2
  m.curdat = m.curdat + 1
  IF MONTH(m.curdat)!=m.curmonth
   m.curmonth = MONTH(m.curdat)
  ELSE 
   LOOP 
  ENDIF 
  
  lcperiod =  STR(YEAR(m.curdat),4)+PADL(m.curmonth,2,'0')
  IF !fso.FolderExists(pbase+'\'+lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN', 'Tarif', 'SHARED', 'cod') > 0
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   USE IN aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', 'sookod', 'SHARED', 'er_c') > 0
   IF USED('sookod')
    USE IN sookod
   ENDIF 
   USE IN tarif
   USE IN aisoms 
   LOOP 
  ENDIF 
  
  WAIT lcperiod WINDOW NOWAIT 

  SELECT aisoms
  SCAN 
   m.mcod = mcod
   m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
  
   IF !fso.FolderExists(pBase+'\'+gcPeriod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\people.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
  
   tnresult = 0
   tnresult = tnresult + OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\people', 'people', 'shar')
   tnresult = tnresult + OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')
   tnresult = tnresult + OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'serrors', 'shar', 'rid')
   tnresult = tnresult + OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'rerrors', 'shar', 'rrid', 'again')
  
   IF tnresult>0
    IF USED('rerrors')
     USE IN rerrors
    ENDIF 
    IF USED('serrors')
     USE IN serrors
    ENDIF 
    IF USED('people')
     USE IN people
    ENDIF 
    IF USED('talon')
     USE IN talon
    ENDIF 
    LOOP 
   ENDIF 
   
*   SELECT people 
*   SET ORDER TO sn_pol
*   SET RELATION TO recid INTO rerrors
  
   SELECT talon 
*   SET RELATION TO sn_pol INTO people 
   SET RELATION TO recid INTO serrors
  
   SCAN
    m.cod    = cod
    m.sn_pol = sn_pol
    m.c_i    = c_i
    DO CASE 
     CASE IsPlk(m.cod) && амбулаторно-поликлиническая помощь
      dimdata(1,1) = dimdata(1,1) + 1
      dimdata(1,5) = dimdata(1,5) + s_all
      IF !EMPTY(serrors.c_err) AND !INLIST(LEFT(serrors.c_err,2),'PN','PK') && Если МЭК
       dimdata(2,1) = dimdata(2,1) + 1
       dimdata(2,5) = dimdata(2,5) + s_all
      ENDIF 

     CASE IsGsp(m.cod)
      dimdata(1,2) = dimdata(1,2) + 1
      dimdata(1,6) = dimdata(1,6) + s_all
      IF !EMPTY(serrors.c_err) AND !INLIST(LEFT(serrors.c_err,2),'PN','PK')
       dimdata(2,2) = dimdata(2,2) + 1
       dimdata(2,6) = dimdata(2,2) + s_all
      ENDIF 

     CASE IsDst(m.cod)
      dimdata(1,3) = dimdata(1,3) + 1
      dimdata(1,7) = dimdata(1,7) + s_all
      IF !EMPTY(serrors.c_err) AND !INLIST(LEFT(serrors.c_err,2),'PN','PK')
       dimdata(2,3) = dimdata(2,3) + 1
       dimdata(2,7) = dimdata(2,7) + s_all
      ENDIF 

     OTHERWISE 

    ENDCASE 
   ENDSCAN 
   
   SET RELATION OFF INTO serrors
   
   SELECT talon 
   SET ORDER TO recid
   SELECT serrors
   SET RELATION TO rid INTO talon 
   SET RELATION TO LEFT(c_err,2) INTO sookod ADDITIVE 

   SCAN
    m.cod    = cod
    m.sn_pol = sn_pol
    m.c_i    = c_i
    DO CASE 
     CASE IsPlk(m.cod) && амбулаторно-поликлиническая помощь
     
      DO CASE  
       CASE BETWEEN(VAL(sookod.refreason),50,55) && 3.1
        dimdata(4,1) = dimdata(4,1) + 1
        dimdata(4,5) = dimdata(4,5) + talon.s_all
       CASE c_err='PKA' && 3.2
        dimdata(5,1) = dimdata(5,1) + 1
        dimdata(5,5) = dimdata(5,5) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),61,63) && 3.3
        dimdata(6,1) = dimdata(6,1) + 1
        dimdata(6,5) = dimdata(6,5) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),64,65) && 3.4
        dimdata(7,1) = dimdata(7,1) + 1
        dimdata(7,5) = dimdata(7,5) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),66,69) && 3.5
        dimdata(8,1) = dimdata(8,1) + 1
        dimdata(8,5) = dimdata(8,5) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),70,75) && 3.6
        dimdata(9,1) = dimdata(9,1) + 1
        dimdata(9,5) = dimdata(9,5) + talon.s_all
       CASE LEFT(c_err,2)='UM' && 3.6.1
        dimdata(10,1) = dimdata(10,1) + 1
        dimdata(10,5) = dimdata(10,5) + talon.s_all
       CASE LEFT(c_err,2)='MM' && 3.6.2
        dimdata(11,1) = dimdata(11,1) + 1
        dimdata(11,5) = dimdata(11,5) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),70,72) && 3.6.3
        dimdata(12,1) = dimdata(12,1) + 1
        dimdata(12,5) = dimdata(12,5) + talon.s_all
       OTHERWISE && 3.7
        dimdata(13,1) = dimdata(13,1) + 1
        dimdata(12,5) = dimdata(13,5) + talon.s_all
       
      ENDCASE 

     CASE IsGsp(m.cod)
      DO CASE  
       CASE BETWEEN(VAL(sookod.refreason),50,55) && 3.1
        dimdata(4,2) = dimdata(4,2) + 1
        dimdata(4,6) = dimdata(4,6) + talon.s_all
       CASE c_err='PKA' && 3.2
        dimdata(5,2) = dimdata(5,2) + 1
        dimdata(5,6) = dimdata(5,6) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),61,63) && 3.3
        dimdata(6,2) = dimdata(6,2) + 1
        dimdata(6,6) = dimdata(6,6) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),64,65) && 3.4
        dimdata(7,2) = dimdata(7,2) + 1
        dimdata(7,6) = dimdata(7,6) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),66,69) && 3.5
        dimdata(8,2) = dimdata(8,2) + 1
        dimdata(8,6) = dimdata(8,6) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),70,75) && 3.6
        dimdata(9,2) = dimdata(9,2) + 1
        dimdata(9,6) = dimdata(9,6) + talon.s_all
       CASE LEFT(c_err,2)='UM' && 3.6.1
        dimdata(10,2) = dimdata(10,2) + 1
        dimdata(10,6) = dimdata(10,6) + talon.s_all
       CASE LEFT(c_err,2)='MM' && 3.6.2
        dimdata(11,2) = dimdata(11,2) + 1
        dimdata(11,6) = dimdata(11,6) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),70,72) && 3.6.3
        dimdata(12,2) = dimdata(12,2) + 1
        dimdata(12,6) = dimdata(12,6) + talon.s_all
       OTHERWISE && 3.7
        dimdata(13,2) = dimdata(13,2) + 1
        dimdata(12,6) = dimdata(13,6) + talon.s_all
       
      ENDCASE 


     CASE IsDst(m.cod)
      DO CASE  
       CASE BETWEEN(VAL(sookod.refreason),50,55) && 3.1
        dimdata(4,3) = dimdata(4,3) + 1
        dimdata(4,7) = dimdata(4,7) + talon.s_all
       CASE c_err='PKA' && 3.2
        dimdata(5,3) = dimdata(5,3) + 1
        dimdata(5,7) = dimdata(5,7) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),61,63) && 3.3
        dimdata(6,3) = dimdata(6,3) + 1
        dimdata(6,7) = dimdata(6,7) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),64,65) && 3.4
        dimdata(7,3) = dimdata(7,3) + 1
        dimdata(7,7) = dimdata(7,7) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),66,69) && 3.5
        dimdata(8,3) = dimdata(8,3) + 1
        dimdata(8,7) = dimdata(8,7) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),70,75) && 3.6
        dimdata(9,3) = dimdata(9,3) + 1
        dimdata(9,7) = dimdata(9,7) + talon.s_all
       CASE LEFT(c_err,2)='UM' && 3.6.1
        dimdata(10,3) = dimdata(10,3) + 1
        dimdata(10,7) = dimdata(10,7) + talon.s_all
       CASE LEFT(c_err,2)='MM' && 3.6.2
        dimdata(11,3) = dimdata(11,3) + 1
        dimdata(11,7) = dimdata(11,7) + talon.s_all
       CASE BETWEEN(VAL(sookod.refreason),70,72) && 3.6.3
        dimdata(12,3) = dimdata(12,3) + 1
        dimdata(12,7) = dimdata(12,7) + talon.s_all
       OTHERWISE && 3.7
        dimdata(13,3) = dimdata(13,3) + 1
        dimdata(12,7) = dimdata(13,7) + talon.s_all
       
      ENDCASE 

     OTHERWISE 
    ENDCASE 
   ENDSCAN 


   SET RELATION OFF INTO serrors
   SET RELATION OFF INTO people
   USE IN serrors
   USE 
   SELECT people 
   SET RELATION OFF INTO rerrors 
   USE IN rerrors
   USE 
   
   IF USED('rerrors')
    USE IN rerrors
   ENDIF 
   IF USED('serrors')
    USE IN serrors
   ENDIF 
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
  ENDSCAN 
  USE 
  USE IN tarif
  USE IN sookod
  
  WAIT CLEAR 
 
 ENDDO 

 CREATE CURSOR curdata (recid i)
 m.llResult = X_Report(m.dotname, m.docname+'.xls', .t.)
 USE IN curdata 

RETURN 
 


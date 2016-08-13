PROCEDURE PresS7
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ'+CHR(13)+CHR(10)+;
  'ÎÒ×ÅÒ ÄËß ÏÐÅÇÅÍÒÀÖÈÈ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\press7.xls')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ PRESS7.XLS!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdata (kusl n(7),susl n(11,2),kmes n(7),smes n(11,2),kkd n(7),skd n(11,2),;
  kuslmek n(7), suslmek n(11,2), kmesmek n(7), smesmek n(11,2), kkdmek n(7), skdmek n(11,2),;
  kuslw023 n(7),kmesw023 n(7),nkdw023 n(7), kuslw0456 n(7),kmesw0456 n(7),nkdw0456 n(7),;
  kuslww23 n(7),suslww23 n(11,2),kmesww23 n(7),smesww23 n(11,2),nkdww23 n(7),skdww23 n(11,2),;
  kuslww456 n(7),suslww456 n(11,2),kmesww456 n(7),smesww456 n(11,2),nkdww456 n(7),skdww456 n(11,2))
 
 CREATE CURSOR curmee (tip n(1),er15 n(5),er311 n(5),er313 n(5),er314 n(5),er321 n(5),er322 n(5),er331 n(5),;
  er34 n(5),er35 n(5),er37 n(5),er41 n(5),er42 n(5),er43 n(5),er45 n(5),er461 n(5),er462 n(5),er514 n(5),;
  er533 n(5),er552 n(5),er573 n(5),er575 n(5))
  
 STORE 0 TO kusl,susl,kmes,smes,kkd,skd,kuslw023,kmesw023,kkdw023,kuslw0456,kmesw0456,kkdw0456,;
  kuslww23,suslww23,kmesww23,smesww23,kkdww23,skdww23,kuslww456,suslww456,kmesww456,smesww456,;
  kkdww456,skdww456,kuslmek,suslmek,kmesmek,smesmek,kkdmek,skdmek

 STORE 0 TO m.usl15,m.usl311,m.usl313,m.usl314,m.usl321,m.usl322,m.usl331,m.usl34,m.usl35,m.usl37,;
  m.usl41,m.usl42,m.usl43,m.usl45,m.usl461,m.usl462,m.usl514,m.usl533,m.usl552,m.usl573,m.usl575,;
  m.mes15,m.mes311,m.mes313,m.mes314,m.mes321,m.mes322,m.mes331,m.mes34,m.mes35,m.mes37

 STORE 0 TO  m.mes41,m.mes42,m.mes43,m.mes45,m.mes461,m.mes462,m.mes514,m.mes533,m.mes552,m.mes573,m.mes575,;
  m.kd15,m.kd311,m.kd313,m.kd314,m.kd321,m.kd322,m.kd331,m.kd34,m.kd35,m.kd37,;
  m.kd41,m.kd42,m.kd43,m.kd45,m.kd461,m.kd462,m.kd514,m.kd533,m.kd552,m.kd573,m.kd575
  
 STORE 0 TO m.uslb14,m.uslb15,m.uslb311,m.uslb312,m.uslb313,m.uslb314,m.uslb321,m.uslb322,m.uslb323,m.uslb331,m.uslb332,;
  m.uslb34,m.uslb36,m.uslb37,m.uslb38,m.uslb39,m.uslb41,m.uslb42,m.uslb43,m.uslb443,m.uslb461,m.uslb462,m.uslb514,;
  m.uslb533,m.uslb573

 STORE 0 TO m.mesb14,m.mesb15,m.mesb311,m.mesb312,m.mesb313,m.mesb314,m.mesb321,m.mesb322,m.mesb323,m.mesb331,m.mesb332,;
  m.mesb34,m.mesb36,m.mesb37,m.mesb38,m.mesb39,m.mesb41,m.mesb42,m.mesb43,m.mesb443,m.mesb461,m.mesb462,m.mesb514,;
  m.mesb533,m.mesb573

 STORE 0 TO m.kdb14,m.kdb15,m.kdb311,m.kdb312,m.kdb313,m.kdb314,m.kdb321,m.kdb322,m.kdb323,m.kdb331,m.kdb332,;
  m.kdb34,m.kdb36,m.kdb37,m.kdb38,m.kdb39,m.kdb41,m.kdb42,m.kdb43,m.kdb443,m.kdb461,m.kdb462,m.kdb514,;
  m.kdb533,m.kdb573

 FOR nm=1 TO 3
  m.lcperiod = STR(tyear,4)+PADL(nm,2,'0')
  IF !fso.FolderExists(m.pbase+'\'+m.lcperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.pbase+'\'+m.lcperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 

  WAIT m.lcperiod+'...' WINDOW NOWAIT 

  SELECT aisoms
  SCAN 
   m.s_pred = s_pred
   IF m.s_pred<=0
    LOOP 
   ENDIF 
   m.mcod = mcod
   IF !fso.FolderExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon.dbf') OR ;
    !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf') OR ;
    !fso.FileExists(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
    LOOP  
   ENDIF 
   IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\e'+m.mcod, 'eerror', 'shar', 'rid')>0
    IF USED('eerror')
     USE IN eerror
    ENDIF 
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   IF OpenFile(m.pbase+'\'+m.lcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('eerror')
     USE IN eerror
    ENDIF 
    IF USED('merror')
     USE IN merror
    ENDIF 
    SELECT aisoms
    LOOP 
   ENDIF 
   
   SELECT talon 
   SET RELATION TO recid INTO eerror
   SCAN 
    m.cod = cod
    m.k_u = k_u
    m.s_all = s_all
    DO CASE 
     CASE IsUsl(m.cod)
      m.kusl = m.kusl + m.k_u
      m.susl = m.susl + m.s_all
      IF !EMPTY(eerror.rid)
       m.kuslmek = m.kuslmek + m.k_u
       m.suslmek = m.suslmek + m.s_all
      ENDIF 
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.kmes = m.kmes + 1
      m.smes = m.smes + m.s_all
      IF !EMPTY(eerror.rid)
       m.kmesmek = m.kmesmek + 1
       m.smesmek = m.smesmek + m.s_all
      ENDIF 
     CASE IsKd(m.cod)
      m.kkd = m.kkd + m.k_u
      m.skd = m.skd + m.s_all
      IF !EMPTY(eerror.rid)
       m.kkdmek = m.kkdmek + m.k_u
       m.skdmek = m.skdmek + m.s_all
      ENDIF 
     OTHERWISE 
    ENDCASE 
   ENDSCAN 
   SET RELATION OFF INTO eerror
   USE IN eerror
   SET ORDER TO recid
   
   SELECT merror
   SET RELATION TO recid INTO talon 
   SCAN 
    m.cod = cod
    m.k_u = k_u
    m.s_all = s_all
    m.s_def = s_1
    m.et   = et
    m.c_err = UPPER(LEFT(err_mee,2))
    m.osn230 = osn230
    IF INLIST(m.et,'2','3')
     DO CASE 
      CASE IsUsl(m.cod)
       IF m.c_err='W0'
        m.kuslw023 = m.kuslw023 + m.k_u
       ELSE 
        m.kuslww23 = m.kuslww23 + m.k_u
        m.suslww23 = m.suslww23 + m.s_def
        DO CASE 
         CASE m.osn230 = '1.5'
          m.usl15 = m.usl15 + m.k_u
         CASE m.osn230 = '3.1.1'
          m.usl311 = m.usl311 + m.k_u
         CASE m.osn230 = '3.1.3'
          m.usl313 = m.usl313 + m.k_u
         CASE m.osn230 = '3.1.4'
          m.usl314 = m.usl314 + m.k_u
         CASE m.osn230 = '3.2.1'
          m.usl321 = m.usl321 + m.k_u
         CASE m.osn230 = '3.2.2'
          m.usl322 = m.usl322 + m.k_u
         CASE m.osn230 = '3.3.1'
          m.usl331 = m.usl331 + m.k_u
         CASE m.osn230 = '3.4'
          m.usl34 = m.usl34 + m.k_u
         CASE m.osn230 = '3.5'
          m.usl35 = m.usl35 + m.k_u
         CASE m.osn230 = '3.7'
          m.usl37 = m.usl37 + m.k_u
         CASE m.osn230 = '4.1'
          m.usl41 = m.usl41 + m.k_u
         CASE m.osn230 = '4.2'
          m.usl42 = m.usl42 + m.k_u
         CASE m.osn230 = '4.3'
          m.usl43 = m.usl43 + m.k_u
         CASE m.osn230 = '4.5'
          m.usl45 = m.usl45 + m.k_u
         CASE m.osn230 = '4.6.1'
          m.usl461 = m.usl461 + m.k_u
         CASE m.osn230 = '4.6.2'
          m.usl462 = m.usl462 + m.k_u
         CASE m.osn230 = '5.1.4'
          m.usl514 = m.usl514 + m.k_u
         CASE m.osn230 = '5.3.3'
          m.usl533 = m.usl533 + m.k_u
         CASE m.osn230 = '5.5.2'
          m.usl552 = m.usl552 + m.k_u
         CASE m.osn230 = '5.7.3'
          m.usl573 = m.usl573 + m.k_u
         CASE m.osn230 = '5.7.5'
          m.usl575 = m.usl575 + m.k_u
        ENDCASE 
       ENDIF 
      CASE IsMes(m.cod) OR IsVMP(m.cod)
       IF m.c_err='W0'
        m.kmesw023 = m.kmesw023 + 1
       ELSE 
        m.kmesww23 = m.kmesww23 + 1
        m.smesww23 = m.smesww23 + m.s_def
        DO CASE 
         CASE m.osn230 = '1.5'
          m.mes15 = m.mes15 + 1
         CASE m.osn230 = '3.1.1'
          m.mes311 = m.mes311 + 1
         CASE m.osn230 = '3.1.3'
          m.mes313 = m.mes313 + 1
         CASE m.osn230 = '3.1.4'
          m.mes314 = m.mes314 + 1
         CASE m.osn230 = '3.2.1'
          m.mes321 = m.mes321 + 1
         CASE m.osn230 = '3.2.2'
          m.mes322 = m.mes322 + 1
         CASE m.osn230 = '3.3.1'
          m.mes331 = m.mes331 + 1
         CASE m.osn230 = '3.4'
          m.mes34 = m.mes34 + 1
         CASE m.osn230 = '3.5'
          m.mes35 = m.mes35 + 1
         CASE m.osn230 = '3.7'
          m.mes37 = m.mes37 + 1
         CASE m.osn230 = '4.1'
          m.mes41 = m.mes41 + 1
         CASE m.osn230 = '4.2'
          m.mes42 = m.mes42 + 1
         CASE m.osn230 = '4.3'
          m.mes43 = m.mes43 + 1
         CASE m.osn230 = '4.5'
          m.mes45 = m.mes45 + 1
         CASE m.osn230 = '4.6.1'
          m.mes461 = m.mes461 + 1
         CASE m.osn230 = '4.6.2'
          m.mes462 = m.mes462 + 1
         CASE m.osn230 = '5.1.4'
          m.mes514 = m.mes514 + 1
         CASE m.osn230 = '5.3.3'
          m.mes533 = m.mes533 + 1
         CASE m.osn230 = '5.5.2'
          m.mes552 = m.mes552 + 1
         CASE m.osn230 = '5.7.3'
          m.mes573 = m.mes573 + 1
         CASE m.osn230 = '5.7.5'
          m.mes575 = m.mes575 + 1
        ENDCASE 
       ENDIF 
      CASE IsKd(m.cod)
       IF m.c_err='W0'
        m.kkdw023 = m.kkdw023 + m.k_u
       ELSE 
        m.kkdww23 = m.kkdww23 + m.k_u
        m.skdww23 = m.skdww23 + m.s_def
        DO CASE 
         CASE m.osn230 = '1.5'
          m.kd15 = m.kd15 + m.k_u
         CASE m.osn230 = '3.1.1'
          m.kd311 = m.kd311 + m.k_u
         CASE m.osn230 = '3.1.3'
          m.kd313 = m.kd313 + m.k_u
         CASE m.osn230 = '3.1.4'
          m.kd314 = m.kd314 + m.k_u
         CASE m.osn230 = '3.2.1'
          m.kd321 = m.kd321 + m.k_u
         CASE m.osn230 = '3.2.2'
          m.kd322 = m.kd322 + m.k_u
         CASE m.osn230 = '3.3.1'
          m.kd331 = m.kd331 + m.k_u
         CASE m.osn230 = '3.4'
          m.kd34 = m.kd34 + m.k_u
         CASE m.osn230 = '3.5'
          m.kd35 = m.kd35 + m.k_u
         CASE m.osn230 = '3.7'
          m.kd37 = m.kd37 + m.k_u
         CASE m.osn230 = '4.1'
          m.kd41 = m.kd41 + m.k_u
         CASE m.osn230 = '4.2'
          m.kd42 = m.kd42 + m.k_u
         CASE m.osn230 = '4.3'
          m.kd43 = m.kd43 + m.k_u
         CASE m.osn230 = '4.5'
          m.kd45 = m.kd45 + m.k_u
         CASE m.osn230 = '4.6.1'
          m.kd461 = m.kd461 + m.k_u
         CASE m.osn230 = '4.6.2'
          m.kd462 = m.kd462 + m.k_u
         CASE m.osn230 = '5.1.4'
          m.kd514 = m.kd514 + m.k_u
         CASE m.osn230 = '5.3.3'
          m.kd533 = m.kd533 + m.k_u
         CASE m.osn230 = '5.5.2'
          m.kd552 = m.kd552 + m.k_u
         CASE m.osn230 = '5.7.3'
          m.kd573 = m.kd573 + m.k_u
         CASE m.osn230 = '5.7.5'
          m.kd575 = m.kd575 + m.k_u
        ENDCASE 
       ENDIF 
      OTHERWISE 
     ENDCASE 
    ELSE 
     DO CASE 
      CASE IsUsl(m.cod)
       IF m.c_err='W0'
        m.kuslw0456 = m.kuslw0456 + m.k_u
       ELSE 
        m.kuslww456 = m.kuslww456 + m.k_u
        m.suslww456 = m.suslww456 + m.s_def
        DO CASE 
         CASE m.osn230 = '1.4'
          m.uslb14 = m.uslb14 + m.k_u
         CASE m.osn230 = '1.5'
          m.uslb15 = m.uslb15 + m.k_u
         CASE m.osn230 = '3.1.1'
          m.uslb311 = m.uslb311 + m.k_u
         CASE m.osn230 = '3.1.2'
          m.uslb312 = m.uslb312 + m.k_u
         CASE m.osn230 = '3.1.3'
          m.uslb313 = m.uslb313 + m.k_u
         CASE m.osn230 = '3.1.4'
          m.uslb314 = m.uslb314 + m.k_u
         CASE m.osn230 = '3.2.1'
          m.uslb321 = m.uslb321 + m.k_u
         CASE m.osn230 = '3.2.2'
          m.uslb322 = m.uslb322 + m.k_u
         CASE m.osn230 = '3.2.3'
          m.uslb323 = m.uslb323 + m.k_u
         CASE m.osn230 = '3.3.1'
          m.uslb331 = m.uslb331 + m.k_u
         CASE m.osn230 = '3.3.2'
          m.uslb332 = m.uslb332 + m.k_u
         CASE m.osn230 = '3.4'
          m.uslb34 = m.uslb34 + m.k_u
         CASE m.osn230 = '3.6'
          m.uslb36 = m.uslb36 + m.k_u
         CASE m.osn230 = '3.7'
          m.uslb37 = m.uslb37 + m.k_u
         CASE m.osn230 = '3.8'
          m.uslb38 = m.uslb38 + m.k_u
         CASE m.osn230 = '3.9'
          m.uslb39 = m.uslb39 + m.k_u
         CASE m.osn230 = '4.1'
          m.uslb41 = m.uslb41 + m.k_u
         CASE m.osn230 = '4.2'
          m.uslb42 = m.uslb42 + m.k_u
         CASE m.osn230 = '4.3'
          m.uslb43 = m.uslb43 + m.k_u
         CASE m.osn230 = '4.4.3'
          m.uslb443 = m.uslb443 + m.k_u
         CASE m.osn230 = '4.6.1'
          m.uslb461 = m.uslb461 + m.k_u
         CASE m.osn230 = '4.6.2'
          m.uslb462 = m.uslb462 + m.k_u
         CASE m.osn230 = '5.1.4'
          m.uslb514 = m.uslb514 + m.k_u
         CASE m.osn230 = '5.3.3'
          m.uslb533 = m.uslb533 + m.k_u
         CASE m.osn230 = '5.7.3'
          m.uslb573 = m.uslb573 + m.k_u
        ENDCASE 
       ENDIF 
      CASE IsMes(m.cod) OR IsVMP(m.cod)
       IF m.c_err='W0'
        m.kmesw0456 = m.kmesw0456 + 1
       ELSE 
        m.kmesww456 = m.kmesww456 + 1
        m.smesww456 = m.smesww456 + m.s_def
        DO CASE 
         CASE m.osn230 = '1.4'
          m.mesb14 = m.mesb14 + m.k_u
         CASE m.osn230 = '1.5'
          m.mesb15 = m.mesb15 + m.k_u
         CASE m.osn230 = '3.1.1'
          m.mesb311 = m.mesb311 + m.k_u
         CASE m.osn230 = '3.1.2'
          m.mesb312 = m.mesb312 + m.k_u
         CASE m.osn230 = '3.1.3'
          m.mesb313 = m.mesb313 + m.k_u
         CASE m.osn230 = '3.1.4'
          m.mesb314 = m.mesb314 + m.k_u
         CASE m.osn230 = '3.2.1'
          m.mesb321 = m.mesb321 + m.k_u
         CASE m.osn230 = '3.2.2'
          m.mesb322 = m.mesb322 + m.k_u
         CASE m.osn230 = '3.2.3'
          m.mesb323 = m.mesb323 + m.k_u
         CASE m.osn230 = '3.3.1'
          m.mesb331 = m.mesb331 + m.k_u
         CASE m.osn230 = '3.3.2'
          m.mesb332 = m.mesb332 + m.k_u
         CASE m.osn230 = '3.4'
          m.mesb34 = m.mesb34 + m.k_u
         CASE m.osn230 = '3.6'
          m.mesb36 = m.mesb36 + m.k_u
         CASE m.osn230 = '3.7'
          m.mesb37 = m.mesb37 + m.k_u
         CASE m.osn230 = '3.8'
          m.mesb38 = m.mesb38 + m.k_u
         CASE m.osn230 = '3.9'
          m.mesb39 = m.mesb39 + m.k_u
         CASE m.osn230 = '4.1'
          m.mesb41 = m.mesb41 + m.k_u
         CASE m.osn230 = '4.2'
          m.mesb42 = m.mesb42 + m.k_u
         CASE m.osn230 = '4.3'
          m.mesb43 = m.mesb43 + m.k_u
         CASE m.osn230 = '4.4.3'
          m.mesb443 = m.mesb443 + m.k_u
         CASE m.osn230 = '4.6.1'
          m.mesb461 = m.mesb461 + m.k_u
         CASE m.osn230 = '4.6.2'
          m.mesb462 = m.mesb462 + m.k_u
         CASE m.osn230 = '5.1.4'
          m.mesb514 = m.mesb514 + m.k_u
         CASE m.osn230 = '5.3.3'
          m.mesb533 = m.mesb533 + m.k_u
         CASE m.osn230 = '5.7.3'
          m.mesb573 = m.mesb573 + m.k_u
        ENDCASE 
       ENDIF 
      CASE IsKd(m.cod)
       IF m.c_err='W0'
        m.kkdw0456 = m.kkdw0456 + m.k_u
       ELSE 
        m.kkdww456 = m.kkdww456 + m.k_u
        m.skdww456 = m.skdww456 + m.s_def
        DO CASE 
         CASE m.osn230 = '1.4'
          m.kdb14 = m.kdb14 + m.k_u
         CASE m.osn230 = '1.5'
          m.kdb15 = m.kdb15 + m.k_u
         CASE m.osn230 = '3.1.1'
          m.kdb311 = m.kdb311 + m.k_u
         CASE m.osn230 = '3.1.2'
          m.kdb312 = m.kdb312 + m.k_u
         CASE m.osn230 = '3.1.3'
          m.kdb313 = m.kdb313 + m.k_u
         CASE m.osn230 = '3.1.4'
          m.kdb314 = m.kdb314 + m.k_u
         CASE m.osn230 = '3.2.1'
          m.kdb321 = m.kdb321 + m.k_u
         CASE m.osn230 = '3.2.2'
          m.kdb322 = m.kdb322 + m.k_u
         CASE m.osn230 = '3.2.3'
          m.kdb323 = m.kdb323 + m.k_u
         CASE m.osn230 = '3.3.1'
          m.kdb331 = m.kdb331 + m.k_u
         CASE m.osn230 = '3.3.2'
          m.kdb332 = m.kdb332 + m.k_u
         CASE m.osn230 = '3.4'
          m.kdb34 = m.kdb34 + m.k_u
         CASE m.osn230 = '3.6'
          m.kdb36 = m.kdb36 + m.k_u
         CASE m.osn230 = '3.7'
          m.kdb37 = m.kdb37 + m.k_u
         CASE m.osn230 = '3.8'
          m.kdb38 = m.kdb38 + m.k_u
         CASE m.osn230 = '3.9'
          m.kdb39 = m.kdb39 + m.k_u
         CASE m.osn230 = '4.1'
          m.kdb41 = m.kdb41 + m.k_u
         CASE m.osn230 = '4.2'
          m.kdb42 = m.kdb42 + m.k_u
         CASE m.osn230 = '4.3'
          m.kdb43 = m.kdb43 + m.k_u
         CASE m.osn230 = '4.4.3'
          m.kdb443 = m.kdb443 + m.k_u
         CASE m.osn230 = '4.6.1'
          m.kdb461 = m.kdb461 + m.k_u
         CASE m.osn230 = '4.6.2'
          m.kdb462 = m.kdb462 + m.k_u
         CASE m.osn230 = '5.1.4'
          m.kdb514 = m.kdb514 + m.k_u
         CASE m.osn230 = '5.3.3'
          m.kdb533 = m.kdb533 + m.k_u
         CASE m.osn230 = '5.7.3'
          m.kdb573 = m.kdb573 + m.k_u
        ENDCASE 
       ENDIF 
      OTHERWISE 
     ENDCASE 
    ENDIF 
   ENDSCAN 
   SET RELATION OFF INTO talon
   USE IN merror
   USE IN talon 
   
   SELECT aisoms
   
  ENDSCAN 
  USE IN aisoms
  WAIT CLEAR 
 ENDFOR

 INSERT INTO curdata FROM MEMVAR 
 SELECT curdata 
 COPY TO &pout\present
 
 m.stot  = m.smesmek+m.smesww23+m.smesww456+m.suslmek+m.suslww23+m.suslww456+m.skdmek+m.skdww23+m.skdww456
 m.smek  = m.smesmek+m.suslmek+m.skdmek
 m.smee  = m.smesww23+m.suslww23+m.skdww23
 m.sekmp = m.smesww456+m.suslww456+m.skdww456

 m.kmek   = m.kmesmek+m.kuslmek+m.kkdmek
 m.kw023  = m.kmesw023+m.kuslw023+m.kkdw023+m.kmesww23+m.kuslww23+m.kkdww23
 m.kww23  = m.kmesww23+m.kuslww23+m.kkdww23
 m.kw0456 = m.kmesw0456+m.kuslw0456+m.kkdw0456+m.kmesww456+m.kuslww456+m.kkdww456
 m.kww456 = m.kmesww456+m.kuslww456+m.kkdww456
 
 m.mes23pr  = ROUND(((m.kmesw023+m.kmesww23)/m.kmes)*100,2)
 m.usl23pr  = ROUND(((m.kuslw023+m.kuslww23)/m.kusl)*100,2)
 m.kd23pr   = ROUND(((m.kkdw023+m.kkdww23)/m.kkd)*100,2)
 m.mes456pr = ROUND(((m.kmesw0456+m.kmesww456)/m.kmes)*100,2)
 m.usl456pr = ROUND(((m.kuslw0456+m.kuslww456)/m.kusl)*100,2)
 m.kd456pr  = ROUND(((m.kkdw0456+m.kkdww456)/m.kkd)*100,2)
 
 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl + "\press7.xls"
 m.lcRepName = pOut + "\PresS7.xls"

 m.llResult = X_Report(m.lcTmpName, m.lcRepName, .T.)
 USE IN curdata 
 
 USE IN curmee

RETURN 
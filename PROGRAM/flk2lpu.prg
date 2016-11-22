PROCEDURE Flk2Lpu

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar')>0
  RETURN .f. 
 ENDIF 
 
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', 'sprabo', 'shar')
  
 lcPeriod = STR(tYear,4) + PADL(tMonth,2,'0')
 lcMmy = SUBSTR(lcPeriod,5,2)+SUBSTR(lcPeriod,4,1)
* lcExt = IIF(m.gcFormat='DOC', 'doc', 'pdf')
 lcExt = 'pdf'
 
 SELECT AisOms
 IF gcUser!='OMS'
  SET FILTER TO Usr == PADL(ALLTRIM(gcUser),6)
 ENDIF 

 SCAN

  WAIT mcod WINDOW NOWAIT 

  lcPath = pBase+'\'+m.gcperiod+'\'+mcod
  IF !fso.FolderExists(lcPath)
   LOOP 
  ENDIF 
  IF !fso.FileExists(lcPath+'\people.dbf') OR !fso.FileExists(lcPath+'\talon.dbf')
   LOOP 
  ENDIF 
  
  DocName = pbase+'\'+m.gcperiod+'\'+mcod+'\b_flk_' + mcod

  DDDocNamec = "DD" + STR(lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
  DSDocNamec = "DS" + STR(lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
  DDDocName = lcpath + '\' + DDDocNamec
  DSDocName = lcpath + '\' + DSDocNamec


  IF fso.FileExists(DocName)
   LOOP 
  ENDIF 

  =MakeCtrl(lcPath)
  SELECT aisoms 

  IF !fso.FileExists(lcPath + '\Pr' + m.qcod + lcMmy + '.' + lcExt) AND ; 
   !fso.FileExists(lcPath + '\Pr' + m.qcod + lcMmy + '.xls') 
   =McPrn(lcPath, .f., .f.)
*   IF aisoms.tpn = .f.
*    =oms6cword(lcPath, .t., .t.)
*   ELSE 
*    =oms6cwordtpn(lcPath, .t., .t.)
*   ENDIF 
*   =Oms6cpdf(lcPath, .f., .f.)
  ENDIF 
  
  SELECT AisOms

  lcLpuID = lpuid
*  m.cTO = ALLTRIM(cfrom)
  m.cTO  = IIF(!EMPTY(ALLTRIM(cfrom)), ALLTRIM(cfrom), ;
  IIF(SEEK(lcLpuID, 'sprabo', 'lpu_id'), 'OMS@'+ALLTRIM(sprabo.abn_name), ''))
  m.un_id    = SYS(3)
  m.bansfile = 'b_flk_'  + mcod
  m.tansfile = 't_flk_'  + mcod
  m.d1file   = 'd1_flk_' + mcod
  m.d2file   = 'd2_flk_' + mcod
  m.d3file   = 'dd_' + mcod
  m.d4file   = 'ds_' + mcod
  m.d5file   = 'ud_' + mcod
*  m.mmid     = m.un_id+'.'+m.gcUser+'@'+m.qmail
  m.mmid     = m.un_id+'.'+m.usrmail+'@'+m.qmail
  m.csubj    = 'OMS#'+lcPeriod+'#'+STR(lcLpuID,4)+'##1'

  poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

  poi.WriteLine('To: '+m.cTO)
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.csubj)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Resent-Message-Id: '+ALLTRIM(cmessage))
  poi.WriteLine('Attachment: '+m.d1file+' Ctrl'+m.qcod+'.dbf')
  poi.WriteLine('Attachment: '+m.d2file+' Pr'+m.qcod+lcMmy+'.'+lcExt)
  IF fso.FileExists(DDDocName+'.'+lcext)
   poi.WriteLine('Attachment: '+m.d3file+' '+DDDocNamec+'.'+lcext)
  ENDIF 
  IF fso.FileExists(DSDocName+'.'+lcext)
   poi.WriteLine('Attachment: '+m.d3file+' '+DSDocNamec+'.'+lcext)
  ENDIF 
 
  poi.Close
  
  IF fso.FileExists(lcPath+'\'+'Ctrl'+m.qcod+'.dbf') 
   fso.CopyFile(lcPath+'\'+'Ctrl'+m.qcod+'.dbf', pAisOms+'\oms\output\'+m.d1file)
  ENDIF 
  IF fso.FileExists(lcPath+'\'+'Pr'+m.qcod+lcMmy+'.'+lcExt) 
   fso.CopyFile(lcPath+'\'+'Pr'+m.qcod+lcMmy+'.'+lcExt, pAisOms+'\oms\output\'+m.d2file)
  ENDIF 
  IF fso.FileExists(DDDocName+'.'+lcext)
   fso.CopyFile(DDDocName+'.'+lcext, pAisOms+'\oms\output\'+m.d3file)
  ENDIF 
  IF fso.FileExists(DSDocName+'.'+lcext)
   fso.CopyFile(DSDocName+'.'+lcext, pAisOms+'\oms\output\'+m.d4file)
  ENDIF 

  fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

  fso.CopyFile(lcPath+'\'+m.tansfile, lcPath+'\'+m.bansfile)
  fso.DeleteFile(lcPath+'\'+m.tansfile)
  
  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 
 
 ENDSCAN 

 WAIT CLEAR 

 USE
 USE IN sprabo
 
 SET ESCAPE &OldEscStatus
RETURN 
 

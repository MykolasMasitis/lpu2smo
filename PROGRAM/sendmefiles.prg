PROCEDURE SendMEFiles

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar', 'mcod')>0
  RETURN .f. 
 ENDIF 

 mmy      = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 
 SELECT AisOms

 SCAN
*  IF s_pred-sum_flk <= 0
*   LOOP 
*  ENDIF 

  m.mcod = mcod
  m.lpu_id = lpuid

  WAIT m.mcod WINDOW NOWAIT 

  MEFile  = 'ME'+UPPER(m.qcod)+STR(lpuid,4)
  MEFilep = pOut+'\'+gcPeriod+'\ME'+UPPER(m.qcod)+STR(lpuid,4)
  
  IF !fso.FileExists(MEFilep+'.zip')
   LOOP 
  ENDIF 

  m.cTO  = 'oms@mgf.msk.oms'
  
  m.un_id    = SYS(3)
  m.bansfile = 'b_me_' + m.mcod
  m.tansfile = 't_me_' + m.mcod
  m.dfile    = 'd_me_' + m.mcod
  m.mmid     = m.un_id+'.'+m.gcUser+'@'+m.qmail
  m.csubj    = 'OMS#'+gcPeriod+'#'+STR(lpu_id,4)+'##RM'

  poi = fso.CreateTextFile(pOut+'\'+gcPeriod+'\'+m.tansfile)

  poi.WriteLine('To: '+m.cTO)
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.csubj)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Attachment: '+m.dfile+' '+MEFile+'.zip')
 
  poi.Close
 
  fso.CopyFile(MEFilep+'.zip', pAisOms+'\oms\output\'+m.dfile)
  fso.CopyFile(pOut+'\'+gcPeriod+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

  SELECT AisOms
  
 ENDSCAN 
 WAIT CLEAR 
 USE 
 
 RETURN 

PROCEDURE SendByAis(cmcod, cdocname, csubject, liszip)
 PRIVATE mcod, docname, subject, iszip

 m.mcod     = cmcod
 m.subject  = csubject
 m.docname  = cdocname && œÓÎÌ˚È ÔÛÚ¸ 
 m.fpath    = SUBSTR(m.docname, 1, RAT('\',m.docname)-1) && “ÓÎ¸ÍÓ ÔÛÚ¸ Í Ù‡ÈÎÛ
 m.fname    = SUBSTR(m.docname, RAT('\',m.docname)+1) && “ÓÎ¸ÍÓ ËÏˇ Ù‡ÈÎ‡
 m.iszip    = liszip

 m.un_id = SYS(3)
 m.bfile = 'b' + m.un_id
 m.tfile = 't' + m.un_id
 m.dfile = 'd' + m.un_id
 m.mmid  = m.un_id+'.OMS@'+m.qmail
 
 ZipOpen(m.fpath + '\' + m.dfile + '.zip')
 ZipFile(m.docname)
 ZipClose()
 
 m.lWasUsedSprLpu = .T.
 m.lWasUsedSprAbo = .T.
 IF !USED('sprlpu')
  IF OpenFile(pbase+'\'+gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
   RETURN 
  ENDIF 
  m.lWasUsedSprLpu = .F.
 ENDIF 
 IF !USED('sprabo')
  IF OpenFile(pbase+'\'+gcperiod+'\nsi\spraboxx', 'sprabo', 'shar', 'lpu_id')>0
   IF m.lWasUsedSprLpu = .F.
    USE IN sprlpu
   ENDIF 
   RETURN 
  ENDIF 
  m.lWasUsedSprAbo = .F.
 ENDIF 
 
 m.lpuid = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
 
 IF m.lpuid==0
  IF m.lWasUsedSprAbo = .F.
   USE IN sprabo
  ENDIF 
  IF m.lWasUsedSprLpu = .F.
   USE IN sprlpu
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈ Õ¿…ƒ≈Õ LPU_ID!'+CHR(13)+CHR(10),0+64,m.mcod)
  RETURN 
 ENDIF 
 
 m.address = IIF(SEEK(m.lpuid, 'sprabo'), sprabo.abn_name, '')
 
 IF EMPTY(m.address)
  IF m.lWasUsedSprAbo = .F.
   USE IN sprabo
  ENDIF 
  IF m.lWasUsedSprLpu = .F.
   USE IN sprlpu
  ENDIF 
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈ Õ¿…ƒ≈Õ ABN_NAME!'+CHR(13)+CHR(10),0+64,m.lpuid)
  RETURN 
 ENDIF 

 m.cTo   = 'USR010@'+m.address

 poi = fso.CreateTextFile(m.fpath + '\' + m.tfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.subject)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Attachment: '+m.dfile+' '+STRTRAN(LOWER(m.fname),'.doc','.zip'))

 poi.Close

 fso.CopyFile(m.fpath + '\' + m.dfile + '.zip', pAisOms+'\usr010\output\'+m.dfile)
 fso.CopyFile(m.fpath + '\'+m.tfile, pAisOms+'\usr010\output\'+m.bfile)
 fso.DeleteFile(m.fpath+'\'+m.tfile)
 
 IF m.lWasUsedSprAbo = .F.
  USE IN sprabo
 ENDIF 
 IF m.lWasUsedSprLpu = .F.
  USE IN sprlpu
 ENDIF 

 MESSAGEBOX('ƒŒ ”Ã≈Õ“ Œ“œ–¿¬À≈Õ!',0+64, '')

RETURN 
 

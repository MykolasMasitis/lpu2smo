PROCEDURE SendVzvFile

 VzvFile = 'vzv13'+m.qcod+'.dbf'
 VzvFileZip = 'vzv13'+m.qcod+'.zip'

 IF !fso.FileExists(pOut+'\'+VzvFile)
  MESSAGEBOX('VZV-‘¿…À Õ≈ —‘Œ–Ã»–Œ¬¿Õ!', 0+16, '')
  RETURN
 ENDIF 
 
 IF fso.FileExists(pOut+'\'+VzvFileZip)
  fso.DeleteFile(pOut+'\'+VzvFileZip)
 ENDIF 
 
 ZipOpen(pOut+'\'+VzvFileZip)
 ZipFile(pOut+'\'+VzvFile)
 ZipClose()

 m.cTo   = 'oms@mgf.msk.oms'
 m.un_id = SYS(3)
 m.bfile = 'b_vzv_'+m.un_id
 m.tfile = 't_vzv_'+m.un_id
 m.dfile = 'd_vzv_' + m.un_id
 m.mmid  = m.un_id+'.OMS@'+m.qmail
 m.csubj = 'OMS#'+gcperiod+'#'+UPPER(m.qcod)+'##VM'

 poi = fso.CreateTextFile(pOut + '\' + m.tfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Attachment: '+m.dfile+' '+VzvFileZip)

 poi.Close

 fso.CopyFile(pOut+'\'+VzvFileZip, pAisOms+'\oms\output\'+m.dfile)
 fso.CopyFile(pOut+'\'+m.tfile, pAisOms+'\oms\output\'+m.bfile)
 
 fso.CopyFile(pOut+'\'+m.tfile, pOut+'\'+m.bfile)
 fso.DeleteFile(pOut+'\'+m.tfile)
 
 MESSAGEBOX('VZV-‘¿…À Œ“œ–¿¬À≈Õ!',0+64, '')

RETURN 
 

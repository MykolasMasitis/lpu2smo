PROCEDURE SendFinFile
 lcPeriod = STR(tYear,4) + PADL(tMonth,2,'0')

 OutDirPeriod = pOut + '\' + lcPeriod
 
 FinFile = 'f13'+m.qcod+'.dbf'
 FinFileZip = 'f13'+m.qcod+'.zip'

 IF !fso.FolderExists(OutDirPeriod)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+OutDirPeriod+'!', 0+16, '')
  RETURN 
 ENDIF 

 IF !fso.FileExists(OutDirPeriod+'\'+FinFile)
  MESSAGEBOX('‘»Õ-‘¿…À Õ≈ —‘Œ–Ã»–Œ¬¿Õ!', 0+16, '')
  RETURN
 ENDIF 
 
 IF fso.FileExists(OutDirPeriod+'\'+FinFileZip)
  fso.DeleteFile(OutDirPeriod+'\'+FinFileZip)
 ENDIF 
 
 ZipOpen(OutDirPeriod+'\'+FinFileZip)
 ZipFile(OutDirPeriod+'\'+FinFile)
 ZipClose()

 m.cTo   = 'oms@mgf.msk.oms'
 m.un_id = SYS(3)
 m.bfile = 'b_fin_'+m.un_id
 m.tfile = 't_fin_'+m.un_id
 m.dfile = 'd_fin_' + m.un_id
 m.mmid  = m.un_id+'.OMS@'+m.qmail
 m.csubj = 'OMS#'+lcPeriod+'#'+UPPER(m.qcod)+'##TM'

 poi = fso.CreateTextFile(OutDirPeriod + '\' + m.tfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Attachment: '+m.dfile+' '+FinFileZip)

 poi.Close

 fso.CopyFile(OutDirPeriod+'\'+FinFileZip, pAisOms+'\oms\output\'+m.dfile)
 fso.CopyFile(OutDirPeriod+'\'+m.tfile, pAisOms+'\oms\output\'+m.bfile)
 
 fso.CopyFile(OutDirPeriod+'\'+m.tfile, OutDirPeriod+'\'+m.bfile)
 fso.DeleteFile(OutDirPeriod+'\'+m.tfile)
 
 MESSAGEBOX('‘»Õ-‘¿…À Œ“œ–¿¬À≈Õ!',0+64, '')

RETURN 
 

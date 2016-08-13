PROCEDURE SendZapr2LpuByAis(m.mcod, m.tip) && m.tip=0 - текущий период, 1 - произвольный период

 m.lpuid   = IIF(SEEK(m.mcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)
 m.aisadr  = IIF(SEEK(m.lpuid, 'sprabo', 'lpu_id'), ALLTRIM(sprabo.abn_name), '')
 m.pDir    = IIF(m.Tip=0, m.pMee, m.pMee)+'\'+m.gcperiod+IIF(m.tip=0,'\','\0000000\')+m.mcod
 m.DocName = m.pDir+'\Rq'+IIF(m.Tip=0,'',flcod)+'.xls'
 
 IF !fso.FileExists(m.DocName)
  MESSAGEBOX('По выбранному ЛПУ'+CHR(13)+CHR(10)+;
   'запрос на подбор карт не сформирован!'+CHR(13)+CHR(10),0+16,m.Docname)
  RETURN 
 ENDIF 
 
 ZipReq = 'Rq'+m.mcod
 IF fso.FileExists(pDir+'\'+ZipReq+'.zip')
  fso.DeleteFile(pDir+'\'+ZipReq+'.zip')
 ENDIF 
 
 ZipOpen(pDir+'\'+ZipReq+'.zip')
 ZipFile(m.docname)
 ZipClose()
 
 m.cTo   = 'usr010@'+m.aisadr
 m.un_id = SYS(3)
 m.bfile = 'brq'+m.un_id
 m.tfile = 'trq'+m.un_id
 m.dfile = 'drq' + m.un_id
 m.mmid  = m.un_id+'.OMS@'+m.qmail
 m.csubj = 'Запрос на экспертизу ('+ALLTRIM(m.qname)+')'

 poi = fso.CreateTextFile(pDir + '\' + m.tfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Attachment: '+m.dfile+' '+ZipReq+'.zip')

 poi.Close

 fso.CopyFile(pDir+'\'+ZipReq+'.zip', pAisOms+'\usr010\output\'+m.dfile)
 fso.CopyFile(pDir+'\'+m.tfile, pAisOms+'\usr010\output\'+m.bfile)
 
 fso.CopyFile(pDir+'\'+m.tfile, pDir+'\'+m.bfile)
 fso.DeleteFile(pDir+'\'+m.tfile)
 
 MESSAGEBOX('ЗАПРОС ОТПРАВЛЕН!',0+64, '')

RETURN 
 

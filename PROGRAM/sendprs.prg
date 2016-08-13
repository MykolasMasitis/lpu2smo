FUNCTION SendPrs(mcod)
 lcPath = pBase+'\'+m.gcperiod+'\'+m.mcod
 IF !fso.FolderExists(lcPath)
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер дхпейрнпхъ кос!'+CHR(13)+CHR(10),0+64,mcod)
  RETURN 
 ENDIF 

lcPeriod  = STR(tYear,4) + PADL(tMonth,2,'0')
 mmy      = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 m.mcod   = mcod
 m.lpu_id = lpuid

 ZipPath = lcPath
 MmyName = 'D'+m.qcod+STR(m.lpu_id,4)+'.'+mmy
  
 IF !fso.FileExists(lcPath+'\'+MmyName)
  MESSAGEBOX(CHR(13)+CHR(10)+'он бшапюммнлс кос оепянрвер ме ятнплхпнбюм!'+CHR(13)+CHR(10),0+64,mcod)
  RETURN 
 ENDIF 

 WAIT m.mcod WINDOW NOWAIT 

 m.cTO  = 'oms@mgf.msk.oms'
  
 m.un_id    = SYS(3)
 m.bansfile = 'b_y_' + m.mcod
 m.tansfile = 't_y_' + m.mcod
 m.dfile    = 'd_y_' + m.mcod
 m.mmid     = m.un_id+'.'+m.usrmail+'@'+m.qmail
 m.csubj    = 'OMS#'+lcPeriod+'#'+STR(lpu_id,4)+'##1'

 poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Attachment: '+m.dfile+' '+mmyname)
 
 poi.Close
 
 fso.CopyFile(lcPath+'\'+mmyname, pAisOms+'\oms\output\'+m.dfile)
 fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)
  
 WAIT CLEAR 

 MESSAGEBOX(CHR(13)+CHR(10)+'нрвер нропюбкем!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 
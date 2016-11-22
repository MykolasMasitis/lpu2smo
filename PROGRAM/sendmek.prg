FUNCTION SendMEK(lcPath)

 lcPeriod = STR(tYear,4) + PADL(tMonth,2,'0')
 lcMcod  = SUBSTR(lcPath, RAT('\',lcPath)+1)
 lcLpuID = IIF(SEEK(lcMcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)
 m.cTO   = IIF(!EMPTY(ALLTRIM(cfrom)), ALLTRIM(cfrom), ;
  IIF(SEEK(lcLpuID, 'sprabo', 'lpu_id'), 'OMS@'+ALLTRIM(sprabo.abn_name), ''))
 
 m.Mmy = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 m.lcext = m.gcformat

 PrFile    = 'Pr' + m.qcod + m.mmy + '.'+m.lcext
 McFile    = 'Mc' + STR(lpuid,4) + m.qcod + m.mmy + '.'+m.lcext
 MkFile    = 'Mk' + STR(lpuid,4) + m.qcod + m.mmy + '.'+m.lcext
 MtFile    = 'Mt' + STR(lpuid,4) + m.qcod + m.mmy + '.'+m.lcext
 PdfFile   = 'pdf' + m.qcod + m.mmy + '.'+m.lcext
 UDFile    = 'ud'+m.qcod+STR(lclpuid,4)+'.dbf'
 UPFile    = 'up'+m.qcod+STR(lclpuid,4)+'.dbf'
 DDDocName = 'DD' + STR(lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)+'.'+m.lcext
 DSDocName = 'DS' + STR(lpuid,4)+LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)+'.'+m.lcext
 PcFile    = 'S_' + m.qcod + m.mmy + '.'+m.lcext

 IF !fso.FileExists(lcPath + '\' + PrFile)
  WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ œ–Œ“Œ ŒÀ¿..." WINDOW NOWAIT 
  IF aisoms.tpn = .f.
   =oms6cword(lcPath, .f., .t.)
  ELSE 
   =oms6cwordtpn(lcPath, .f., .t.)
  ENDIF 
  SELECT AisOms
  WAIT CLEAR 
 ENDIF 

 IF !fso.FileExists(lcpath+'\'+'ctrl'+m.qcod+'.dbf')
  WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ ‘¿…À¿ Œÿ»¡Œ ..." WINDOW NOWAIT 
  =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
  SELECT AisOms
  WAIT CLEAR 
 ENDIF 

 IF !fso.FileExists(lcPath + '\' + McFile)
  WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ ¿ “Œ¬ Ã› ..." WINDOW NOWAIT 
  =McPrn(lcPath, .f., .t.)
  SELECT AisOms
  WAIT CLEAR 
 ENDIF 

 IF !fso.FileExists(lcPath + '\' + MkFile)
  WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ –≈≈—“–¿ ¿ “Œ¬..." WINDOW NOWAIT 
  =MkPrn(lcPath, .f., .t.)
  SELECT AisOms
  WAIT CLEAR 
 ENDIF 

 IF !fso.FileExists(lcPath + '\' + MtFile)
  WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ “¿¡À»◊ÕŒ… ‘Œ–Ã€ ¿ “Œ¬..." WINDOW NOWAIT 
*  =MtPrn(lcPath, .f., .t.)
  =MtPrn2(lcPath, .f., .t.)
  SELECT AisOms
  WAIT CLEAR 
 ENDIF 

 IF fso.FileExists(lcPath + '\' + PdfFile)
  m.lIsPdfFile = .t.
 ELSE 
  m.lIsPdfFile = .f.
 ENDIF 

 IF fso.FileExists(lcPath + '\' + UDFile)
  m.lIsUDFile = .t.
 ELSE 
  m.lIsUDFile = .f.
 ENDIF 

 IF fso.FileExists(lcPath + '\' + UPFile)
  m.lIsUPFile = .t.
 ELSE 
  m.lIsUPFile = .f.
 ENDIF 

 IF fso.FileExists(lcPath + '\' + PcFile)
  m.lIsPcFile = .t.
 ELSE 
  m.lIsPcFile = .f.
 ENDIF 
 
 ZipPath = lcPath
 ZipName = 'ot'+m.qcod+STR(lcLpuID,4)+'.zip'
 MmyName = 'ot'+m.qcod+STR(lcLpuID,4)+'.'+mmy

 m.un_id    = SYS(3)
 m.bansfile = 'b_mek_' + mcod
 m.tansfile = 't_mek_' + mcod
 m.dfile    = 'd_mek_' + mcod
 m.mmid     = m.un_id+'.'+m.usrmail+'@'+m.qmail
* m.csubj    = 'OMS#'+m.gcperiod+'###1'
 m.csubj    = 'OMS#'+lcPeriod+'#'+PADL(lcLpuID,4,'0')+'##1'

 poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: '+ALLTRIM(cmessage))
 poi.WriteLine('Attachment: '+m.dfile+' '+MmyName)
 poi.Close

 IF fso.FileExists(lcpath+'\'+ZipName)
  fso.DeleteFile(lcpath+'\'+ZipName)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+MmyName)
  fso.DeleteFile(lcpath+'\'+MmyName)
 ENDIF 

 SET DEFAULT TO (lcpath)
 ZipOpen(MmyName, lcPath+'\')
 IF fso.FileExists(lcpath+'\'+'ctrl'+m.qcod+'.dbf')
  ZipFile('ctrl'+m.qcod+'.dbf', .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+PrFile)
  ZipFile(PrFile, .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+McFile)
  ZipFile(McFile, .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+MkFile)
  ZipFile(MkFile, .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+MtFile)
  ZipFile(MtFile, .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+PdfFile)
  ZipFile(PdfFile, .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+UdFile)
  ZipFile(UdFile, .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+UpFile)
  ZipFile(UpFile, .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+DDDocName)
  ZipFile(DDDocName, .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+DSDocName)
  ZipFile(DSDocName, .T.)
 ENDIF 
 IF fso.FileExists(lcpath+'\'+PCFile)
  ZipFile(PCFile, .T.)
 ENDIF 
 ZipClose()
 SET DEFAULT TO (pBin)

 fso.CopyFile(lcpath+'\'+MmyName, pAisOms+'\oms\output\'+m.dfile)
 fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)
 fso.CopyFile(lcPath+'\'+m.tansfile, lcPath+'\'+m.bansfile)
 fso.DeleteFile(lcPath+'\'+m.tansfile)
 
RETURN
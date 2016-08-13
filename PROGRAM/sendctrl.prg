FUNCTION SendCtrl(lcPath)
 lcPeriod = SUBSTR(lcPath, RAT('\',lcPath,2)+1, RAT('\',lcPath,1)-RAT('\',lcPath,2)-1)

 lcMcod  = SUBSTR(lcPath, RAT('\',lcPath)+1)
 lcLpuID = IIF(SEEK(lcMcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)
* m.cTO  = IIF(SEEK(lcLpuID, 'sprabo', 'lpu_id'), ALLTRIM(sprabo.abn_name), '')
 m.cTO  = IIF(!EMPTY(ALLTRIM(cfrom)), ALLTRIM(cfrom), ;
  IIF(SEEK(lcLpuID, 'sprabo', 'lpu_id'), 'OMS@'+ALLTRIM(sprabo.abn_name), ''))
 
 lcMmy = SUBSTR(lcPeriod,5,2)+SUBSTR(lcPeriod,4,1)

 IF !fso.FileExists(lcPath + '\Pr' + m.qcod + lcMmy + '.pdf')
  IF aisoms.tpn = .f.
   =oms6cword(lcPath, .t., .t.)
  ELSE 
   =oms6cwordtpn(lcPath, .t., .t.)
  ENDIF 
*  =oms6cpdf(lcPath, .f.)
  SELECT AisOms
 ENDIF 
 IF !fso.FileExists(lcpath+'\'+'ctrl'+m.qcod+'.dbf')
  =MakeCtrl(pBase+'\'+m.gcperiod+'\'+mcod)
  SELECT AisOms
 ENDIF 
 
 m.un_id    = SYS(3)
 m.bansfile = 'b' + m.un_id
 m.tansfile = 't' + m.un_id
 m.d1file   = 'd1' + m.un_id
 m.d2file   = 'd2' + m.un_id
 m.mmid     = m.un_id+'.'+m.usrmail+'@'+m.qmail
 m.csubj    = 'OMS#'+lcPeriod+'###1'

 poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

* poi.WriteLine('To: oms@'+m.cTO)
 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: '+ALLTRIM(cmessage))
 poi.WriteLine('Attachment: '+m.d1file+' Ctrl'+m.qcod+'.dbf')
 poi.WriteLine('Attachment: '+m.d2file+' Pr'+m.qcod+lcMmy+'.pdf')
 
 poi.Close
 
 fso.CopyFile(lcPath+'\'+'Ctrl'+m.qcod+'.dbf', pAisOms+'\oms\output\'+m.d1file)
 fso.CopyFile(lcPath+'\'+'Pr'+m.qcod+lcMmy+'.pdf', pAisOms+'\oms\output\'+m.d2file)
 fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)
 
 MESSAGEBOX('‘¿…À€ Œ“œ–¿¬À≈Õ€!',0+64,'')

RETURN  
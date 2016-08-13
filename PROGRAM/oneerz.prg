FUNCTION OneERZ(lcDir, IsOK)
 IF !fso.FolderExists(lcDir)
  RETURN .F.
 ENDIF 

 IF !fso.FileExists(lcDir + '\People.dbf')
  RETURN .F.
 ENDIF 


 IF OpenFile("&lcDir\People", "People", "SHARE")>0
  RETURN .F.
 ENDIF 
 
 IF RECCOUNT('people')<=0
  USE IN people 
  RETURN .F.
 ENDIF 
 
* m.mcod = RIGHT(ALLTRIM(lcDir),7)
 
 fso.CopyFile(ptempl+'\'+'\Zapros.dbf', lcDir+'\Zapros.dbf', .t.)
 oSettings.CodePage(lcDir+'\Zapros.dbf', 866, .t.)
 
 IF OpenFile("&lcDir\Zapros", "Zapros", "SHARE")>0
  USE IN people
  RETURN .F.
 ENDIF 

 SELECT people
 SCAN 
  SCATTER MEMVAR 
  
  IF m.mcod = '0371001'
   m.date_in  = m.d_beg
   m.date_out = m.d_end
  ELSE 
   m.date_in  = MIN(m.d_beg, m.tdat1)
   m.date_out = m.tdat2
  ENDIF 

  m.recid    = PADL(m.recid,6,'0')
  m.fam      = m.fam
  m.im       = m.im
  m.ot       = m.ot
  m.q        = m.qcod
  IF OCCURS(' ', ALLTRIM(m.sn_pol))>0
   m.s_pol    = SUBSTR(m.sn_pol, 1, AT(' ',m.sn_pol)-1)
   m.n_pol    = SUBSTR(m.sn_pol, AT(' ',m.sn_pol)+1)
  ELSE 
   m.s_pol    = ''
   m.n_pol    = m.sn_pol
  ENDIF 
  DO CASE 
   CASE ISALPHA(m.sn_pol)
    m.tip_d='¬'
   CASE OCCURS(' ', ALLTRIM(m.sn_pol))==0
    m.tip_d='œ'
   OTHERWISE 
    m.tip_d='—'
  ENDCASE 
  m.dr       = DToS(m.dr)

  INSERT INTO Zapros FROM MEMVAR 
    
 ENDSCAN 
 USE IN Zapros
 USE IN people 

 SELECT AisOms
   
 ChVal  = SYS(3)
 ID     = ALLTRIM(ChVal+'.'+m.usrmail+'@'+m.qmail)
  
 TFile  = 'terz_' + mcod
 BFile  = 'berz_' + mcod
 DFile  = 'derz_' + mcod

 iii = 1
 DO WHILE fso.FileExists(pAisOms+'\'+m.usrmail+'\OUTPUT\'+m.bfile)
  m.tfile  = 'terz_' + m.mcod + '_' + PADL(iii,2,'0')
  m.bfile  = 'berz_' + m.mcod + '_' + PADL(iii,2,'0')
  m.dfile  = 'derz_' + m.mcod + '_' + PADL(iii,2,'0')
  iii = iii + 1
 ENDDO 
   
 fso.CopyFile(lcDir+'\Zapros.dbf', PAisOms+'\'+m.usrmail+'\OutPut\'+DFile)
 fso.DeleteFile(lcDir+'\Zapros.dbf')
   
 poi = FCREATE('&PAisOms\&usrmail\OutPut\&TFile')
 IF poi != -1
  =FPUTS(poi,'To: erz@mgf.msk.oms')
  =FPUTS(poi,'Message-Id: &ID')
  =FPUTS(poi,'Subject: ERZ_sverka4n')
  fzz = 'q_' + PADL(MONTH(DATE()),2,'0')+RIGHT(ALLTRIM(STR(YEAR(DATE()))),1)+'.dbf'
  =FPUTS(poi,'Attachment: &DFile &Fzz')
 ENDIF 
 =FCLOSE(poi)
 
 oTFile = fso.GetFile('&PAisOms\&usrmail\OutPut\&TFile')
 oTFile.Move('&PAisOms\&usrmail\OutPut\&BFile')

 REPLACE erz_id WITH m.id, erz_status WITH 1

 IF IsOk==.t.
  MESSAGEBOX('«¿œ–Œ— Œ“œ–¿¬À≈Õ!', 0+64, '')
 ENDIF 
RETURN .T.
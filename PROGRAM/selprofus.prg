PROCEDURE SelProfUs

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÎÁÐÀÒÜ ÌÅÄÈÖÈÍÑÊÓÞ ÏÎÌÎÙÜ ÏÎ ÏÐÎÔÈËÞ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 m.profil = ''
 DO FORM selprofil
 
 IF EMPTY(m.profil)
  RETURN 
 ENDIF 

 m.dotname = pTempl+'\selguests.xls'
 IF !fso.FileExists(m.dotname)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË ØÀÁËÎÍÀ '+CHR(13)+CHR(10)+;
   m.dotname,0+64,'')
  RETURN 
 ENDIF 

 m.RepDir = fso.GetParentFolderName(pbin) + '\REPS'
 IF !fso.FolderExists(m.RepDir)
  fso.CreateFolder(m.RepDir)
 ENDIF 

 m.BookName = m.repdir+'\se'+m.gcperiod
 m.IsOpDoc = IsOpenExcelDoc('se'+m.gcperiod)
 IF m.IsOpDoc
  IF !CloseExcelDoc('se'+m.gcperiod)
   MESSAGEBOX('ÔÀÉË '+'se'+m.gcperiod+' ÎÒÊÐÛÒ!',0+64,'')
   RETURN .f. 
  ENDIF 
 ENDIF 

 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.gcperiod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN aisoms
  USE IN tarif
  RETURN 
 ENDIF 
 IF OpenFile(m.pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'mcod')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  USE IN aisoms
  USE IN sprlpu
  USE IN tarif
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi'+'\profus', 'profus', 'shar', 'cod')>0
  IF USED('profus')
   USE IN profus
  ENDIF 
  USE IN aisoms
  USE IN sprlpu
  USE IN tarif
  RETURN 
 ENDIF 
  
 CREATE CURSOR allguests (nrec i AUTOINC, period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(30), ;
  fam c(25), im c(25), ot c(25), dr d, w n(1), d_u d, ds c(6), otd c(4), pcod c(10), ;
  cod n(6), tip c(1), k_u n(3), n_kd n(3), s_all n(11,2), d_beg d, d_end d, codname c(40), lpuname c(40), q c(2),;
  isexp l, et c(1), osn230 c(5), s_1 n(11,2), s_2 n(11,2))

 CREATE CURSOR curbooks (bkname c(13))

 =SelAlienOne()
 
 USE IN allguests
 USE IN sprlpu
 USE IN tarif
 USE IN pilot
 USE IN profus

 PUBLIC oExcel AS Excel.Application
 WAIT "Çàïóñê MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 
 
 oexcel.SheetsInNewWorkbook = 1
 AllBook = oExcel.Workbooks.Add()
 SELECT curbooks
 SCAN 
  m.lpucod = SUBSTR(bkname,3,7)
  m.bkname = m.RepDir+'\'+bkname
  IF !fso.FileExists(m.bkname)
   LOOP 
  ENDIF 
  OneBook  = oExcel.Workbooks.Add(m.bkname)
  OneSheet = OneBook.WorkSheets(1).Copy(,AllBook.ActiveSheet)
  OneBook.Close
  fso.DeleteFile(m.bkname)
  AllBook.ActiveSheet.Name = m.lpucod
 ENDSCAN 
 USE IN curbooks 

 IF fso.FileExists(m.repdir+'\se'+m.gcperiod+'.xls')
  fso.DeleteFile(m.repdir+'\se'+m.gcperiod+'.xls')
 ENDIF  
 AllBook.SaveAs(m.repdir+'\se'+m.gcperiod,18)
 oExcel.Visible=.t.

RETURN 


FUNCTION SelAlienOne()

 SELECT aisoms
 SCAN 
  m.lpuid = lpuid
  m.mcod = mcod
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon
   ENDIF 
   USE IN people
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'error', 'shar', 'rid')>0
   IF USED('error')
    USE IN error
   ENDIF 
   USE IN people
   USE IN talon
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar', 'rid')>0
   IF USED('merror')
    USE IN merror
   ENDIF 
   USE IN error
   USE IN people
   USE IN talon
   LOOP 
  ENDIF 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  m.docname = m.repdir+'\se'+m.mcod+'.XLS'
  INSERT INTO curbooks (bkname) VALUES ('SE'+m.mcod+'.XLS')
  IF fso.FileExists(m.docname)
   fso.DeleteFile(m.docname)
  ENDIF 

  CREATE CURSOR curguests (nrec i AUTOINC, period c(7), lpuid i(4), mcod c(7), sn_pol c(25), c_i c(30), ;
   fam c(25), im c(25), ot c(25), dr d, w n(1),d_u d, ds c(6), otd c(4), pcod c(10), cod n(6), tip c(1), k_u n(3), n_kd n(3),;
   s_all n(11,2), d_beg d, d_end d, codname c(40), lpuname c(40), q c(2),;
   isexp l, et c(1), osn230 c(5), s_1 n(11,2), s_2 n(11,2))

  m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.name, '')
 
  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO recid INTO error ADDITIVE 
  SET RELATION TO recid INTO merror ADDITIVE 
  SET RELATION TO cod INTO profus ADDITIVE 
  SCAN 
   IF !EMPTY(error.rid)
    LOOP 
   ENDIF 
   IF profus.profil!=m.profil
    LOOP 
   ENDIF 

   m.cod = cod
   m.sn_pol = sn_pol
   m.c_i    = c_i
   m.fam    = people.fam
   m.im     = people.im
   m.ot     = people.ot
   m.dr     = people.dr
   m.w      = people.w
   m.d_beg  = people.d_beg
   m.d_end  = people.d_end
   m.d_u    = d_u
   m.ds     = ds
   m.otd    = otd
   m.pcod   = pcod
   m.cod    = cod
   m.tip    = tip
   m.k_u    = k_u
   m.n_kd   = n_kd
   m.s_all  = s_all
   m.isexp  = IIF(!EMPTY(merror.osn230), .t., .f.)
   m.et     = merror.et
   m.osn230 = merror.osn230
   m.s_1    = merror.s_1
   m.s_2    = merror.s_2
   
   m.codname = IIF(SEEK(m.cod, 'tarif'), tarif.comment, '')

   INSERT INTO curguests (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,;
    d_u,ds,otd,pcod, cod, tip, k_u, n_kd, s_all, d_beg, d_end, codname,lpuname,;
    isexp,et,osn230,s_1,s_2) VALUES ;
    (m.gcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,;
    m.d_u,m.ds,m.otd,m.pcod,m.cod,m.tip, m.k_u, m.n_kd, m.s_all, ;
    m.d_beg, m.d_end, m.codname,m.lpuname,m.isexp,m.et,m.osn230,m.s_1,m.s_2) 

   INSERT INTO allguests (period,lpuid,mcod,sn_pol,c_i,fam,im,ot,dr,w,;
    d_u,ds,otd,pcod, cod,tip,k_u, n_kd, s_all, d_beg, d_end, codname,lpuname,;
    isexp,et,osn230,s_1,s_2) VALUES ;
    (m.gcperiod,m.lpuid,m.mcod,m.sn_pol,m.c_i,m.fam,m.im,m.ot,m.dr,m.w,;
    m.d_u,m.ds,m.otd,m.pcod,m.cod,m.tip,m.k_u, m.n_kd, m.s_all, ;
    m.d_beg, m.d_end, m.codname,m.lpuname,m.isexp,m.et,m.osn230,m.s_1,m.s_2) 

  ENDSCAN 

  SET RELATION OFF INTO error
  SET RELATION OFF INTO merror
  SET RELATION OFF INTO people
  SET RELATION OFF INTO profus
  USE IN people
  USE IN talon
  USE IN error
  USE IN merror
  
  SELECT curguests
  IF RECCOUNT()>0
   m.llResult = X_Report(m.dotname, m.docname, .f.)
  ENDIF 
  USE IN curguests
 
  SELECT aisoms
  
  WAIT CLEAR 

 ENDSCAN 
 USE IN aisoms

RETURN 


FUNCTION MtPrn2(pathparam, IsVisible, IsQuit)
 m.lcpath  = pathparam
 m.mcod    = mcod
 m.lpuid   = lpuid
 m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.koplate = s_pred - sum_flk 
 m.period = NameOfMonth(tMonth)+ ' ' + STR(tYear,4)
 m.mmy    = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)
 
 IF OpFiles()>0
  =ClFiles()
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdata (n_rec n(5), enp c(16), ds c(6), d_beg c(10), d_end c(10), osn230 c(6), ername c(100), s_all n(11,2))
 
 SELECT talon
 SET RELATION TO sn_pol INTO people
 SELECT serror
 SET RELATION TO rid INTO talon
 SET RELATION TO LEFT(c_err,2) INTO sookod ADDITIVE 
 
 m.n_rec = 1
 SCAN
  m.f = UPPER(f)
  IF m.f!='S'
   LOOP 
  ENDIF 

  m.enp    = talon.sn_pol
  m.ds     = talon.ds
*  m.d_beg  = DTOC(people.d_beg)
*  m.d_end  = DTOC(people.d_end)
  m.d_beg  = DTOC(talon.d_u-talon.k_u+1)
  m.d_end  = DTOC(talon.d_u)
  m.osn230 = ALLTRIM(sookod.osn230)
  m.ername = ALLTRIM(sookod.comment)
  m.s_all  = talon.s_all
  
  INSERT INTO curdata FROM MEMVAR 
  
  m.n_rec = m.n_rec + 1

 ENDSCAN 
 
 IF RECCOUNT('curdata')<=0
  m.n_rec = 1
  m.enp    = ''
  m.ds     = ''
  m.d_beg  = ''
  m.d_end  = ''
  m.osn230 = ''
  m.ername = ''
  m.s_all  = 0
  INSERT INTO curdata FROM MEMVAR 
 ENDIF 
 
 SET RELATION OFF INTO sookod
 SET RELATION OFF INTO talon 
 SELECT talon 
 SET RELATION OFF INTO people
 
 =ClFiles()

 CREATE CURSOR curdata2 (profil c(3), pr_name c(100), k_pred n(5), s_pred n(11,2), ;
  k_def n(5), s_def n(11,2), k_opl n(5), s_opl n(11,2))
 INDEX on profil TAG profil 
 SET ORDER TO profil

 =OpenFile(m.lcpath+'\Talon', 'Talon', 'Shar')
 =OpenFile(pcommon+'\prv002xx', 'prv002', 'shar', 'profil')
 =OpenFile(m.lcpath+'\e'+m.mcod, 'sError', 'Shar', 'rid')
 SELECT talon 
 SET RELATION TO recid INTO serror
 SET RELATION TO PADL(ALLTRIM(profil),3,'0') INTO prv002 ADDITIVE 

 SCAN 

  m.profil = profil
  m.k_u    = k_u
  m.s_all  = s_all
  m.pr_name = m.profil+', '+ALLTRIM(prv002.pr_name)

  IF EMPTY(sError.rid)
   m.k_opl = k_u
   m.s_opl = s_all
   m.k_def = 0
   m.s_def = 0
  ELSE 
   m.k_opl = 0
   m.s_opl = 0
   m.k_def = k_u
   m.s_def = s_all
  ENDIF 

  IF SEEK(m.profil, 'curdata2')
   UPDATE curdata2 SET k_pred = k_pred+m.k_u, s_pred = s_pred + m.s_all,;
    k_opl = k_opl+m.k_opl, s_opl = s_opl + m.s_opl, ;
    k_def = k_def+m.k_def, s_def = s_def + m.s_def ;
    WHERE profil=m.profil
  ELSE 
   INSERT INTO curdata2 (profil,pr_name,k_pred,s_pred,k_def,s_def,k_opl,s_opl) VALUES ;
   (m.profil,m.pr_name,m.k_u,m.s_all,m.k_def,m.s_def,m.k_opl,m.s_opl)
  ENDIF 
  
 ENDSCAN 
 SET RELATION OFF INTO prv002
 SET RELATION OFF INTO serror
 USE IN talon 
 USE IN prv002
 USE IN serror
 
* SELECT curdata2
* BROWSE 
* SELECT curdata

 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl + "\MtxxxxQQmmy.xls"
 m.lcRepName = lcPath + "\Mt" + STR(m.lpuid,4) + m.qcod + m.mmy+'.xls'
 m.lcRepName2 = lcPath + "\Mt" + STR(m.lpuid,4) + m.qcod + m.mmy
 m.lcDbfName = 'curdata'

 m.n_akt = mcod+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 m.d_akt = DTOC(DATE())
 m.n_mek = PADL(tMonth,2,'0') + '/' + STR(tYear,4)
 m.d_mek = DTOC(TTOD(aisoms.sent))

 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)

 USE IN curdata
 USE IN curdata2

 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 IF fso.FileExists(m.lcRepName2+'.pdf')
  fso.DeleteFile(m.lcRepName2+'.pdf')
 ENDIF 
 oDoc = oExcel.Workbooks.Add(m.lcRepName)
 TRY 
  odoc.SaveAs(m.lcRepName2,57)
 CATCH 
 ENDTRY 
 odoc.close 

 SELECT aisoms
 
RETURN 
 

FUNCTION OpFiles
 tn_rslt = 0 
 tn_rslt = tn_rslt + OpenFile(lcpath+'\talon', 'talon', 'shar', 'recid')
 tn_rslt = tn_rslt + OpenFile(lcpath+'\people', 'people', 'shar', 'sn_pol')
 tn_rslt = tn_rslt + OpenFile(lcpath+'\e'+m.mcod, 'serror', 'shar')
 tn_rslt = tn_rslt + OpenFile(pcommon+'\prv002xx', 'prv002', 'shar', 'profil')
 tn_rslt = tn_rslt + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', 'sookod', 'shar', 'er_c')
RETURN tn_rslt

FUNCTION ClFiles
 IF USED('talon')
  USE IN talon
 ENDIF 
 IF USED('people')
  USE IN people
 ENDIF 
 IF USED('serror')
  USE IN serror
 ENDIF 
 IF USED('prv002')
  USE IN prv002
 ENDIF 
 IF USED('sookod')
  USE IN sookod
 ENDIF 
RETURN 
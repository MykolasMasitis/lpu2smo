FUNCTION McPrn(para1, IsVisible, IsQuit)
 
 m.lcpath  = para1
 m.mcod    = mcod
 m.lpuid   = lpuid
 IF USED('sprlpu')
  m.lpu_name = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 ELSE 
  m.lpu_name = ''
 ENDIF 
 
 IF OpFiles()>0
  =ClFiles()
  RETURN 
 ENDIF 

 m.mmy    = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)

 m.n_akt    = mcod+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 m.d_akt    = DTOC(DATE())
 m.akt_mon  = NameOfMonth(tMonth)
 m.akt_year = STR(tYear,4)
 m.nr_akt   = m.mcod+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 
 m.k_predst = 0
 m.s_predst = 0
 m.k_bad    = 0
 m.s_bad    = 0
 m.s_ok     = 0
 
 SELECT talon
 SET RELATION TO recid INTO serror
 SCAN 
  m.k_u   = k_u
  m.s_all = s_all
  m.k_predst = m.k_predst + m.k_u
  m.s_predst = m.s_predst + m.s_all
  IF EMPTY(sError.rid)
   m.s_ok = m.s_ok + m.s_all
  ELSE 
   m.k_bad = m.k_bad + m.k_u
   m.s_bad = m.s_bad + m.s_all
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO serror
 
 =ClFiles()

 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl + "\McxxxxQQmmy.xls"
 m.lcRepName = lcPath + "\Mc" + STR(m.lpuid,4) + m.qcod + m.mmy+'.xls'
 m.lcRepName2 = lcPath + "\Mc" + STR(m.lpuid,4) + m.qcod + m.mmy

 CREATE CURSOR curdata (recid i, cod n(6), name c(10))
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 USE IN curdata 

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

 SELECT aisoms

RETURN 

FUNCTION OpFiles
 tn_rslt = 0 
 tn_rslt = tn_rslt + OpenFile(lcpath+'\talon', 'talon', 'shar')
 tn_rslt = tn_rslt + OpenFile(lcpath+'\e'+m.mcod, 'serror', 'shar', 'rid')
RETURN tn_rslt

FUNCTION ClFiles
 IF USED('talon')
  USE IN talon
 ENDIF 
 IF USED('serror')
  USE IN serror
 ENDIF 
RETURN 
FUNCTION MkPrn2(pathparam, IsVisible, IsQuit)

 m.lcpath  = pathparam

 m.mcod     = mcod
 m.lpuid    = lpuid
 m.lpuname  = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.lpuadr   = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 
 m.period = NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 m.mmy    = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)
 
 m.kol_paz = paz
 m.kol_sch = 0
 m.summa   = s_pred

 eeFile = 'e'+m.mcod

 m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 m.d_akt = DTOC(DATE())
 m.akt = '¹ '+m.n_akt+' ÎÒ '+m.d_akt
 
 m.dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 m.dat2 = DTOC(GOMONTH(CTOD(m.dat1),1)-1)
 m.period = 'ïåðèîä ñ '+m.dat1+' ïî '+m.dat2

 m.smo_name = m.qname+', '+m.qcod
 m.smo_adr = [Ã. ÌÎÑÊÂÀ, 77]
 m.smoname = m.smo_name+', '+m.smo_adr
 
 m.moname = m.lpuname+', '+m.lpuadr


 USE &lcPath\Talon IN 0 ALIAS Talon SHARED 
 USE &lcPath\People IN 0 ALIAS People SHARED
 USE &lcPath\&eeFile IN 0 ALIAS sError SHARED ORDER rid 
 USE &lcPath\&eeFile IN 0 ALIAS rError SHARED ORDER rrid AGAIN 
 
 USE pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx' IN 0 ALIAS sookod SHARED ORDER er_c

 SELECT People 
 SET RELATION TO RecId INTO rError

 m.paz_amb   = 0
 m.usl_amb   = 0
 m.sum_amb   = 0

 m.paz_gosp  = 0
 m.usl_gosp  = 0 
 m.sum_gosp  = 0

 m.paz_dstac  = 0
 m.usl_dstac = 0
 m.sum_dstac = 0

 m.pr_sum = 0

 m.paz_all_ok  = 0
 m.paz_all_bad = 0
 m.usl_all_ok  = 0
 m.usl_all_bad = 0
 m.sum_all_ok  = 0
 m.sum_all_bad = 0

 m.usl_amb_ok  = 0
 m.usl_amb_bad = 0
 m.sum_amb_ok  = 0
 m.sum_amb_bad = 0

 m.usl_gosp_ok  = 0
 m.usl_gosp_bad = 0
 m.sum_gosp_ok  = 0
 m.sum_gosp_bad = 0

 m.usl_dstac_ok  = 0
 m.usl_dstac_bad = 0
 m.sum_dstac_ok  = 0
 m.sum_dstac_bad = 0

 SCAN 
  DO CASE 
   CASE tip_p == 1
    m.paz_amb = m.paz_amb + 1
   CASE tip_p == 2
    m.paz_gosp = m.paz_gosp + 1
   CASE tip_p == 3
    m.paz_amb  = m.paz_amb + 1
    m.paz_gosp = m.paz_gosp + 1
  ENDCASE 
  IF EMPTY(rError.rid)
   m.paz_all_ok = m.paz_all_ok + IIF(tip_p==3, 2, 1)
   DO CASE 
    CASE tip_p == 1
    CASE tip_p == 2
    CASE tip_p == 3
   ENDCASE 
  ELSE 
   m.paz_all_bad = m.paz_all_bad + IIF(tip_p==3, 2, 1)
   DO CASE 
    CASE tip_p == 1
    CASE tip_p == 2
    CASE tip_p == 3
   ENDCASE 
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO rError
 SET ORDER TO sn_pol
 *USE 
 USE IN rError

 CREATE CURSOR curstac  (rid i AUTOINC,recid c(7),otd c(8),sn_pol c(25),osn230 c(5),c_err c(3),s_all n(11,2))
 INDEX ON rid TAG rid
 SET ORDER TO rid 
 CREATE CURSOR curdstac (rid i AUTOINC,recid c(7),otd c(8),sn_pol c(25),osn230 c(5),c_err c(3),s_all n(11,2))
 INDEX ON rid TAG rid
 SET ORDER TO rid 
 CREATE CURSOR curpolk  (rid i AUTOINC,recid c(7),otd c(8),sn_pol c(25),osn230 c(5),c_err c(3),s_all n(11,2))
 INDEX ON rid TAG rid
 SET ORDER TO rid 

 SELECT sError
 SET RELATION TO LEFT(c_err,2) INTO sookod ADDITIVE 
 SELECT Talon 
 SET RELATION TO sn_pol INTO people 
 SET RELATION TO RecId INTO sError ADDITIVE 
 SCAN 
  m.cod = cod
  m.d_type = d_type

  m.pr_sum = m.pr_sum + s_all

  m.otd    = otd
  m.recid  = people.recid_lpu
  m.sn_pol = sn_pol
  m.osn230 = ALLTRIM(sookod.osn230)
  m.c_err  = LEFT(sError.c_err,2)
  m.s_all  = s_all

  DO CASE 
   CASE IsMes(m.cod) OR IsVMP(m.cod)
    m.usl_gosp = m.usl_gosp + 1
    m.sum_gosp = m.sum_gosp + s_all
   CASE IsKD(m.cod)
    m.usl_dstac = m.usl_dstac + 1
    m.sum_dstac = m.sum_dstac + s_all
   CASE IsUsl(m.cod)
    m.usl_amb = m.usl_amb + 1
    m.sum_amb = m.sum_amb + s_all
   OTHERWISE 
  ENDCASE 

  IF EMPTY(sError.rid)
   m.usl_all_ok = m.usl_all_ok + 1
   m.sum_all_ok = m.sum_all_ok + s_all
   DO CASE 
    CASE IsMes(m.cod) OR IsVMP(m.cod)
     m.usl_gosp_ok = m.usl_gosp_ok + 1
     m.sum_gosp_ok = m.sum_gosp_ok + s_all

    CASE IsKD(m.cod)
     m.usl_dstac_ok = m.usl_dstac_ok + 1
     m.sum_dstac_ok = m.sum_dstac_ok + s_all

    CASE IsUsl(m.cod)
     m.usl_amb_ok = m.usl_amb_ok + 1
     m.sum_amb_ok = m.sum_amb_ok + s_all
	
    OTHERWISE 
   ENDCASE 
  ELSE 
   m.usl_all_bad = m.usl_all_bad + 1
   m.sum_all_bad = m.sum_all_bad + s_all
   DO CASE 
    CASE IsMes(m.cod) OR IsVMP(m.cod)
     m.usl_gosp_bad = m.usl_gosp_bad + 1
     m.sum_gosp_bad = m.sum_gosp_bad + s_all

     INSERT INTO curstac FROM MEMVAR 

    CASE IsKD(m.cod)
     m.usl_dstac_bad = m.usl_dstac_bad + 1
     m.sum_dstac_bad = m.sum_dstac_bad + s_all

     INSERT INTO curdstac FROM MEMVAR 

    CASE IsUsl(m.cod)
     m.usl_amb_bad = m.usl_amb_bad + 1
     m.sum_amb_bad = m.sum_amb_bad + s_all

     INSERT INTO curpolk FROM MEMVAR 

    OTHERWISE 
   ENDCASE 
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO sError
 SET RELATION OFF INTO people 
 USE 
 SELECT sError
 SET RELATION OFF INTO sookod
 USE 
 USE IN people 
 USE IN sookod

 
 CREATE CURSOR curdata (recid i)
 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl + "\MkxxxxQQmmy.xls"
 m.lcRepName = lcPath + "\Mk" + STR(m.lpuid,4) + m.qcod + m.mmy+'.xls'
 m.lcRepName2 = lcPath + "\Mk" + STR(m.lpuid,4) + m.qcod + m.mmy
 m.lcDbfName = 'aisoms'

 m.n_akt = mcod+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 m.d_akt = DTOC(DATE())
 m.n_mek = PADL(tMonth,2,'0') + '/' + STR(tYear,4)
 m.d_mek = DTOC(TTOD(aisoms.sent))

 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 
 USE IN curdata 
 USE IN curstac
 USE IN curdstac
 USE IN curpolk 

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

 SELECT AisOms

RETURN  


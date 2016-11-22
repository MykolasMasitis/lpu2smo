FUNCTION oms6cn(lcPath, IsVisible, IsQuit)
 
 tn_result = 0
 tn_result = tn_result + OpenFile(pcommon+'\smo', 'smo', 'shar', 'code')
 tn_result = tn_result + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx', 'sprcokr', 'shar', 'cokr')
 tn_result = tn_result + OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'shar', 'cod')
 IF tn_result > 0
  IF USED('smo')
   USE IN smo
  ENDIF 
  IF USED('sprcokr')
   USE IN sprcokr
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 
 
 m.lIsPr4 = .F.
 IF fso.FileExists(pbase+'\'+gcperiod+'\pr4.dbf')
  m.lIsPr4 = .F.
  IF OpenFile(pbase+'\'+gcperiod+'\pr4', 'pr4', 'shar', 'lpuid')>0
   IF USED('pr4')
    USE IN pr4
   ENDIF 
  ELSE 
   m.lIsPr4 = .T.
  ENDIF 
 ENDIF 

 SELECT AisOms

 m.mcod       = mcod
 eeFile = 'e'+m.mcod
 tn_result = 0
 tn_result = tn_result + OpenFile(lcpath+'\Talon', 'talon', 'shar')
 tn_result = tn_result + OpenFile(lcpath+'\People', 'people', 'shar', 'sn_pol')
 tn_result = tn_result + OpenFile(lcpath+'\'+eeFile, 'serror', 'shar', 'rid')
 IF tn_result>0
  IF USED('smo')
   USE IN smo
  ENDIF 
  IF USED('sprcokr')
   USE IN sprcokr
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('talon')
   USE IN talon
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('serror')
   USE IN serror
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT aisoms

 m.mmy        = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 m.lpuid      = lpuid
 m.lpuname    = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.cokr       = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.cokr), '')
 m.cokr_name  = IIF(SEEK(m.cokr, 'sprcokr'), ALLTRIM(sprcokr.name_okr), '')
 m.smoname    = IIF(SEEK(m.qcod, 'smo'), ALLTRIM(smo.fullname), '')
 m.arcfname   = 'b'+m.mcod+'.'+m.mmy
 m.datpriemki = TTOC(Recieved)
 m.finval     = finval
 m.udsum      = 0
 m.koplate    = 0 
 m.koplpf     = 0
 m.cmessage   = ALLTRIM(cmessage)
 IF USED('pr4')
  IF SEEK(m.lpuid, 'pr4')
   m.udsum = pr4.s_others
   m.koplpf=m.finval-pr4.s_others+pr4.s_guests+pr4.s_npilot+pr4.s_empty
  ENDIF 
 ENDIF 
 
 m.IsPilot  = IIF(SEEK(m.lpuid, 'pilot'), .T., .F.)
 m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn', 'lpu_id'), .t., .f.)	
 
 m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 
 m.kol_paz = paz
 m.kol_sch = 0
 m.summa   = s_pred

 poi_file   = fso.GetFile(lcPath + '\' + arcfname)
 m.arcfdate = poi_file.DateLastModified
 
 ZipItemCount = 5

 m.DotName = pTempl + "\Prqqmmy.xls"
 DocName = lcPath + "\Pr" + LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)

 DIMENSION dimdata(4,11)
 dimdata = 0 
 
 CREATE CURSOR paz1 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol 
 SET ORDER TO sn_pol
 CREATE CURSOR paz2 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol 
 SET ORDER TO sn_pol
 CREATE CURSOR paz3 (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol 
 SET ORDER TO sn_pol
 CREATE CURSOR paz1ok (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol 
 SET ORDER TO sn_pol
 CREATE CURSOR paz2ok (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol 
 SET ORDER TO sn_pol
 CREATE CURSOR paz3ok (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol 
 SET ORDER TO sn_pol
 
 SELECT Talon 
 SET RELATION TO sn_pol INTO people
 SET RELATION TO RecId  INTO sError ADDITIVE 
 SET RELATION TO cod    INTO tarif ADDITIVE 
 REPLACE ALL Mp WITH ''
 
 SCAN
  SCATTER MEMVAR 
  m.cod       = cod
  m.sn_pol    = sn_pol
  m.IsErr     = IIF(!EMPTY(serror.rid), .T., .F.)
  m.prmcod    = people.prmcod
  m.prlpuid   = IIF(SEEK(m.prmcod, 'sprlpu', 'mcod'), sprlpu.lpu_id, '')
  m.IsPrPilot = IIF(SEEK(m.prlpuid, 'sprlpu'), .T., .F.)
  m.s_all     = s_all 
  m.rslt      = rslt
  m.fil_id    = fil_id
  m.otd       = SUBSTR(otd,2,2)
  m.d_type    = d_type 
  m.lpu_ord   = lpu_ord
  m.ord       = ord
  
*  m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod) OR IsEko(m.cod)), .T., .F.)
  m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)

  IF m.IsLpuTpn=.t.
   m.IsUslTpn = IIF(SEEK(m.fil_id, 'lputpn', 'fil_id'), .t., .f.)
  ELSE 
   m.IsUslTpn = .f.
  ENDIF 

  m.Is02      = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
  
  DO CASE 
   CASE EMPTY(m.prmcod) && неприкрепленные
    dimdata(3,2)=dimdata(3,2)+1
    dimdata(3,3)=dimdata(3,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz3')
     INSERT INTO paz3 FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz3ok')
      INSERT INTO paz3ok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s'
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'

*     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod) OR IsEko(m.cod)
     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
      dimdata(3,8) = dimdata(3,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     OTHERWISE 
*      IF m.IsLpuTpn = .T.
*       m.fil_id = fil_id
*       IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
*        dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
*        REPLACE Mp WITH '4'
*       ELSE 
*        dimdata(3,5) = dimdata(3,5) + IIF(m.IsErr,0,m.s_all)
*        IF m.Is02
*         dimdata(3,7) = dimdata(3,7) + IIF(m.IsErr,0,m.s_all)
*        ENDIF 
*       ENDIF 
*      ELSE 
       dimdata(3,5) = dimdata(3,5) + IIF(m.IsErr,0,m.s_all)
       IF m.Is02
        dimdata(3,7) = dimdata(3,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 
*      ENDIF 

    ENDCASE 
    dimdata(3,10) = dimdata(3,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(3,9) = dimdata(3,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
   
   CASE m.mcod  = m.prmcod && свои пациенты
    dimdata(1,2)=dimdata(1,2)+1
    dimdata(1,3)=dimdata(1,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz1')
     INSERT INTO paz1 FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz1ok')
      INSERT INTO paz1ok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 

     CASE m.IsTpnR = .T. OR m.d_type='s'
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'

*     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod) OR IsEko(m.cod)
     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
      dimdata(1,8) = dimdata(1,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE m.otd='93' AND IsStac(m.mcod)
      dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     OTHERWISE 
*      IF m.IsLpuTpn = .T.
*       m.fil_id = fil_id
*       IF !SEEK(m.fil_id, 'lputpn', 'fil_id') AND INLIST(m.otd,'70','73')
*        dimdata(1,11) = dimdata(1,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
*        REPLACE Mp WITH '4'
*       ELSE 
*        dimdata(1,5) = dimdata(1,5) + IIF(m.IsErr,0,m.s_all)
*        IF m.Is02
*         dimdata(1,7) = dimdata(1,7) + IIF(m.IsErr,0,m.s_all)
*        ENDIF 
*       ENDIF 
*      ELSE 
       dimdata(1,5) = dimdata(1,5) + IIF(m.IsErr,0,m.s_all)
       IF m.Is02
        dimdata(1,7) = dimdata(1,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 
*      ENDIF 

    ENDCASE 
    dimdata(1,10) = dimdata(1,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(1,9) = dimdata(1,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 
    
   CASE m.mcod != m.prmcod && чужие пациенты
    dimdata(2,2)=dimdata(2,2)+1
    dimdata(2,3)=dimdata(2,3)+m.s_all
    IF !SEEK(m.sn_pol, 'paz2')
     INSERT INTO paz2 FROM MEMVAR 
    ENDIF 
    IF m.IsErr = .F.
     IF !SEEK(m.sn_pol, 'paz2ok')
      INSERT INTO paz2ok FROM MEMVAR 
     ENDIF 
    ENDIF 

    DO CASE 

     CASE m.IsTpnR = .T. OR m.d_type='s'
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'

*     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod) OR IsEko(m.cod)
     CASE IsMes(m.cod) OR IsKdS(m.cod) OR IsVMP(m.cod)
      dimdata(2,8) = dimdata(2,8) + IIF(m.IsErr,0,m.s_all)

     CASE INLIST(m.otd,'70','73') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE INLIST(m.otd,'01','90','93') AND IsStac(m.mcod)
      dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'
     
     CASE m.ord=7 AND m.lpu_ord=7665
      dimdata(3,11) = dimdata(3,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
      REPLACE Mp WITH '4'

     OTHERWISE 
*      IF m.IsLpuTpn = .T.
*       m.fil_id = fil_id
*       IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
*        dimdata(2,11) = dimdata(2,11) + IIF(m.IsErr,0,m.s_all) && допуслуги!
*        REPLACE Mp WITH '4'
*       ELSE 
*        dimdata(2,5) = dimdata(2,5) + IIF(m.IsErr,0,m.s_all)
*        IF m.Is02
*         dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
*        ENDIF 
*        IF (m.lpu_ord>0 AND m.Is02=.F.) OR (m.lpu_ord=0 AND INLIST(m.otd,'08','92') AND m.Is02=.F.)
*         dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
*        ENDIF 
*       ENDIF 
*      ELSE 
       dimdata(2,5) = dimdata(2,5) + IIF(m.IsErr,0,m.s_all)
       IF m.Is02
        dimdata(2,7) = dimdata(2,7) + IIF(m.IsErr,0,m.s_all)
       ENDIF 
       IF (m.lpu_ord>0 AND m.Is02=.F.) OR (m.lpu_ord=0 AND INLIST(m.otd,'08','92') AND m.Is02=.F.)
*       IF (m.lpu_ord>0 AND m.lpu_ord=m.prlpuid AND m.Is02=.F.) OR (m.lpu_ord=0 AND INLIST(m.otd,'08','92') AND m.Is02=.F.)
        dimdata(2,6) = dimdata(2,6) + IIF(m.IsErr,0,m.s_all)
       ENDIF 
*      ENDIF 

    ENDCASE 
    dimdata(2,10) = dimdata(2,10) + IIF(m.IsErr,0,m.s_all)
    IF IsVMP(m.cod)
     dimdata(2,9) = dimdata(2,9) + IIF(m.IsErr,0,m.s_all)
    ENDIF 

   OTHERWISE 

  ENDCASE 
 ENDSCAN  
 
 SET RELATION OFF INTO people
 SET RELATION OFF INTO sError
 SET RELATION OFF INTO tarif
 USE 
 USE IN sError
 USE IN people 
 
 USE IN smo
 USE IN sprcokr
 USE IN tarif
 IF USED('pr4')
  USE IN pr4
 ENDIF 
 
 dimdata(1,1) = RECCOUNT('paz1')
 dimdata(2,1) = RECCOUNT('paz2')
 dimdata(3,1) = RECCOUNT('paz3')
 dimdata(1,4) = RECCOUNT('paz1ok')
 dimdata(2,4) = RECCOUNT('paz2ok')
 dimdata(3,4) = RECCOUNT('paz3ok')
 
 dimdata(4,1) = dimdata(1,1) + dimdata(2,1)  + dimdata(3,1)
 dimdata(4,2) = dimdata(1,2) + dimdata(2,2)  + dimdata(3,2)
 dimdata(4,3) = dimdata(1,3) + dimdata(2,3)  + dimdata(3,3)
 dimdata(4,4) = dimdata(1,4) + dimdata(2,4)  + dimdata(3,4)
 dimdata(4,5) = dimdata(1,5) + dimdata(2,5)  + dimdata(3,5)
 dimdata(4,6) = dimdata(1,6) + dimdata(2,6)  + dimdata(3,6)
 dimdata(4,7) = dimdata(1,7) + dimdata(2,7)  + dimdata(3,7)
 dimdata(4,8) = dimdata(1,8) + dimdata(2,8)  + dimdata(3,8)
 dimdata(4,9) = dimdata(1,9) + dimdata(2,9)  + dimdata(3,9)
 dimdata(4,10)= dimdata(1,10)+ dimdata(2,10) + dimdata(3,10)
 dimdata(4,11)= dimdata(1,11)+ dimdata(2,11) + dimdata(3,11)
 
 USE IN paz1
 USE IN paz2
 USE IN paz3
 USE IN paz1ok
 USE IN paz2ok
 USE IN paz3ok

IF m.IsPilot
  m.koplate  = m.koplpf
  m.koplate2 = m.koplate+dimdata(4,8)+dimdata(4,11)
ELSE 
 m.koplate  = dimdata(4,10)
 m.koplate2 = 0
ENDIF  

 CREATE CURSOR curdata (recid i)
 m.llResult = X_Report(m.dotname, m.docname+'.xls', m.IsVisible)
 USE IN curdata 

 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 IF fso.FileExists(m.docname+'.pdf')
  fso.DeleteFile(m.docname+'.pdf')
 ENDIF 
 oDoc = oExcel.Workbooks.Add(m.docname+'.xls')
 TRY 
  odoc.SaveAs(m.docname,57)
 CATCH 
 ENDTRY 

 SELECT AisOms

RETURN  


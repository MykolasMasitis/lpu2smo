FUNCTION MkPrn(lcPath, IsVisible, IsQuit)

 m.mcod  = mcod
 m.lpuid = lpuid
 m.lpuname  = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.lpuadr   = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.adres), '')
 
 m.period = NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 m.mmy = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)
 
 m.kol_paz = paz
 m.kol_sch = 0
 m.summa = s_pred

 DotName = pTempl + "\MkxxxxQQmmy.dot"
 DocName = lcPath + "\Mk" + STR(m.lpuid,4) + m.qcod + m.mmy

 eeFile = 'e'+m.mcod

 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 oDoc = oWord.Documents.Add(dotname)

 m.n_akt = mcod + m.qcod + PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 oDoc.Bookmarks('n_akt').Select  
 oWord.Selection.TypeText(m.n_akt)
 
 m.d_akt = DTOC(DATE())
 oDoc.Bookmarks('d_akt').Select  
 oWord.Selection.TypeText(m.d_akt)
 
 m.dat1 = '01.'+PADL(tMonth,2,'0')+'.'+STR(tYear,4)
 oDoc.Bookmarks('dat1').Select  
 oWord.Selection.TypeText(m.dat1)

 m.dat2 = DTOC(GOMONTH(CTOD(m.dat1),1)-1)
 oDoc.Bookmarks('dat2').Select  
 oWord.Selection.TypeText(m.dat2)

 m.vid_mek = '1'
 oDoc.Bookmarks('vid_mek').Select  
 oWord.Selection.TypeText(m.vid_mek)

 m.smo_name = m.qname+', '+m.qcod
 oDoc.Bookmarks('smo_name').Select  
 oWord.Selection.TypeText(m.smo_name)

 m.smo_adr = [Ã. ÌÎÑÊÂÀ, 77]
 oDoc.Bookmarks('smo_adr').Select  
 oWord.Selection.TypeText(m.smo_adr)

 oDoc.Bookmarks('lpu_name').Select  
 oWord.Selection.TypeText(m.lpuname)

 oDoc.Bookmarks('lpu_adr').Select  
 oWord.Selection.TypeText(m.lpuadr)


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

nRowGosp  = 0 
nRowDStac = 0
nRowAmb   = 0

SELECT sError
SET RELATION TO LEFT(c_err,2) INTO sookod ADDITIVE 
SELECT Talon 
SET RELATION TO sn_pol INTO people 
SET RELATION TO RecId INTO sError ADDITIVE 
SCAN 
 m.cod = cod
 m.d_type = d_type

 m.pr_sum = m.pr_sum + s_all

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
	
	oDoc.Tables(23).rows(3+nRowAmb).Select
   
   OTHERWISE 
  ENDCASE 
 ELSE 
  m.usl_all_bad = m.usl_all_bad + 1
  m.sum_all_bad = m.sum_all_bad + s_all
  DO CASE 
   CASE IsMes(m.cod) OR IsVMP(m.cod)
    m.usl_gosp_bad = m.usl_gosp_bad + 1
    m.sum_gosp_bad = m.sum_gosp_bad + s_all

    oDoc.Tables(19).rows(3+nRowGosp).Select
    oWord.Selection.InsertRows
    
    oDoc.Tables(19).Rows(3+nRowGosp).Cells(1).Select
    oWord.Selection.TypeText(otd)
    oDoc.Tables(19).Rows(3+nRowGosp).Cells(2).Select
    oWord.Selection.TypeText(otd)
    oDoc.Tables(19).Rows(3+nRowGosp).Cells(3).Select
    oWord.Selection.TypeText(people.recid_lpu)
    oDoc.Tables(19).Rows(3+nRowGosp).Cells(4).Select
    oWord.Selection.TypeText(PADL(tMonth,2,'0'))
    oDoc.Tables(19).Rows(3+nRowGosp).Cells(5).Select
    oWord.Selection.TypeText(sn_pol)
    oDoc.Tables(19).Rows(3+nRowGosp).Cells(6).Select
    oWord.Selection.TypeText(ALLTRIM(sookod.osn230))
    oDoc.Tables(19).Rows(3+nRowGosp).Cells(7).Select
    oWord.Selection.TypeText(LEFT(sError.c_err,2))
    oDoc.Tables(19).Rows(3+nRowGosp).Cells(8).Select
    oWord.Selection.TypeText(TRANSFORM(s_all, '99 999 999.99'))
	nRowGosp  = nRowGosp + 1

   CASE IsKD(m.cod)
    m.usl_dstac_bad = m.usl_dstac_bad + 1
    m.sum_dstac_bad = m.sum_dstac_bad + s_all

    oDoc.Tables(21).rows(3+nRowDStac).Select
    oWord.Selection.InsertRows
    
    oDoc.Tables(21).Rows(3+nRowDStac).Cells(1).Select
    oWord.Selection.TypeText(otd)
    oDoc.Tables(21).Rows(3+nRowDStac).Cells(2).Select
    oWord.Selection.TypeText(otd)
    oDoc.Tables(21).Rows(3+nRowDStac).Cells(3).Select
    oWord.Selection.TypeText(people.recid_lpu)
    oDoc.Tables(21).Rows(3+nRowDStac).Cells(4).Select
    oWord.Selection.TypeText(PADL(tMonth,2,'0'))
    oDoc.Tables(21).Rows(3+nRowDStac).Cells(5).Select
    oWord.Selection.TypeText(sn_pol)
    oDoc.Tables(21).Rows(3+nRowDStac).Cells(6).Select
    oWord.Selection.TypeText(ALLTRIM(sookod.osn230))
    oDoc.Tables(21).Rows(3+nRowDStac).Cells(7).Select
    oWord.Selection.TypeText(LEFT(sError.c_err,2))
    oDoc.Tables(21).Rows(3+nRowDStac).Cells(8).Select
    oWord.Selection.TypeText(TRANSFORM(s_all, '99 999 999.99'))
	nRowDStac  = nRowDStac + 1

   CASE IsUsl(m.cod)
    m.usl_amb_bad = m.usl_amb_bad + 1
    m.sum_amb_bad = m.sum_amb_bad + s_all

    oDoc.Tables(23).rows(3+nRowAmb).Select
    oWord.Selection.InsertRows
    
    oDoc.Tables(23).Rows(3+nRowAmb).Cells(1).Select
    oWord.Selection.TypeText(otd)
    oDoc.Tables(23).Rows(3+nRowAmb).Cells(2).Select
    oWord.Selection.TypeText(otd)
    oDoc.Tables(23).Rows(3+nRowAmb).Cells(3).Select
    oWord.Selection.TypeText(people.recid_lpu)
    oDoc.Tables(23).Rows(3+nRowAmb).Cells(4).Select
    oWord.Selection.TypeText(PADL(tMonth,2,'0'))
    oDoc.Tables(23).Rows(3+nRowAmb).Cells(5).Select
    oWord.Selection.TypeText(sn_pol)
    oDoc.Tables(23).Rows(3+nRowAmb).Cells(6).Select
    oWord.Selection.TypeText(ALLTRIM(sookod.osn230))
    oDoc.Tables(23).Rows(3+nRowAmb).Cells(7).Select
    oWord.Selection.TypeText(LEFT(sError.c_err,2))
    oDoc.Tables(23).Rows(3+nRowAmb).Cells(8).Select
    oWord.Selection.TypeText(TRANSFORM(s_all, '99 999 999.99'))
	nRowAmb  = nRowAmb + 1

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

oDoc.Bookmarks('pr_sum').Select  
oWord.Selection.TypeText(TRANSFORM(m.pr_sum,'99 999 999.99'))

oDoc.Bookmarks('paz_amb').Select  
oWord.Selection.TypeText(TRANSFORM(m.paz_amb,'99999'))
oDoc.Bookmarks('usl_amb').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_amb,'999999'))
oDoc.Bookmarks('sum_amb').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_amb,'99 999 999.99'))

oDoc.Bookmarks('paz_gosp').Select  
oWord.Selection.TypeText(TRANSFORM(m.paz_gosp,'99999'))
oDoc.Bookmarks('usl_gosp').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_gosp,'999999'))
oDoc.Bookmarks('sum_gosp').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_gosp,'99 999 999.99'))

oDoc.Bookmarks('paz_dstac').Select  
oWord.Selection.TypeText(TRANSFORM(m.paz_dstac,'99999'))
oDoc.Bookmarks('usl_dstac').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_dstac,'999999'))
oDoc.Bookmarks('sum_dstac').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_dstac,'99 999 999.99'))

oDoc.Bookmarks('paz_all_ok').Select  
oWord.Selection.TypeText(TRANSFORM(m.paz_all_ok,'99999'))
oDoc.Bookmarks('usl_all_ok').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_all_ok,'99999'))
oDoc.Bookmarks('sum_all_ok').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_all_ok,'99 9999 999.99'))
oDoc.Bookmarks('usl_amb_ok').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_amb_ok,'99999'))
oDoc.Bookmarks('sum_amb_ok').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_amb_ok,'99 9999 999.99'))
oDoc.Bookmarks('usl_gosp_ok').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_gosp_ok,'99999'))
oDoc.Bookmarks('sum_gosp_ok').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_gosp_ok,'99 9999 999.99'))
oDoc.Bookmarks('usl_dstac_ok').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_dstac_ok,'99999'))
oDoc.Bookmarks('sum_dstac_ok').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_dstac_ok,'99 9999 999.99'))

oDoc.Bookmarks('paz_all_bad').Select  
oWord.Selection.TypeText(TRANSFORM(m.paz_all_bad,'99999'))
oDoc.Bookmarks('usl_all_bad').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_all_bad,'99999'))
oDoc.Bookmarks('sum_all_bad').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_all_bad,'99 9999 999.99'))
oDoc.Bookmarks('usl_amb_bad').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_amb_bad,'99999'))
oDoc.Bookmarks('sum_amb_bad').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_amb_bad,'99 9999 999.99'))
oDoc.Bookmarks('usl_gosp_bad').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_gosp_bad,'99999'))
oDoc.Bookmarks('sum_gosp_bad').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_gosp_bad,'99 9999 999.99'))
oDoc.Bookmarks('usl_dstac_bad').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_dstac_bad,'99999'))
oDoc.Bookmarks('sum_dstac_bad').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_dstac_bad,'99 9999 999.99'))

oDoc.Bookmarks('usl_tot_iskl').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_all_bad,'99 9999 999.99'))
oDoc.Bookmarks('sum_tot_iskl').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_all_bad,'99 9999 999.99'))

oDoc.Bookmarks('usl_gosp_iskl').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_gosp_bad,'99 9999 999.99'))
oDoc.Bookmarks('sum_gosp_iskl').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_gosp_bad,'99 9999 999.99'))

oDoc.Bookmarks('usl_dstac_iskl').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_dstac_bad,'99 9999 999.99'))
oDoc.Bookmarks('sum_dstac_iskl').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_dstac_bad,'99 9999 999.99'))

oDoc.Bookmarks('usl_amb_iskl').Select  
oWord.Selection.TypeText(TRANSFORM(m.usl_amb_bad,'99 9999 999.99'))
oDoc.Bookmarks('sum_amb_iskl').Select  
oWord.Selection.TypeText(TRANSFORM(m.sum_amb_bad,'99 9999 999.99'))
 
 SELECT AisOms

 oDoc.SaveAs(DocName, 0)
 TRY 
  oDoc.SaveAs(DocName, 17)
 CATCH 
 ENDTRY 
 
 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  oDoc.Close
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 
 
RETURN  


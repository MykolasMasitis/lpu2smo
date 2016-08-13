PROCEDURE PrnSvTabl
 PARAMETERS IsNeedZero
 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\pr4.dbf')
  MESSAGEBOX('ÔÀÉË PR4 ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\pr4', 'pr4', 'shar', 'lpuid')>0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  RETURN 
 ENDIF 

 PUBLIC oWord as Word.Application
 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 
 
 DotName = pTempl+'\SvTabl.dot'
 DocName = m.pBase+'\'+STR(tYear,4)+PADL(tMonth,2,'0')+'\SvTabl'

 oDoc = oWord.Documents.Add(dotname)
 oTable = oDoc.Tables(1)

 npp = 1
 tot_s_pred   = 0
 tot_sum_flk  = 0
 tot_sum_prin = 0
 tot_k_opl    = 0 
 
 SELECT aisoms
 SET RELATION TO lpuid INTO pr4

 SCAN FOR IIF(!IsNeedZero, s_pred>0, 1=1)
  lpuname = IIF(SEEK(lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')

  sum_prin = s_pred-sum_flk
  
  IF !EMPTY(pr4.lpuid)
   m.koplate = pr4.finval - pr4.s_others + pr4.s_guests + (pr4.s_kompl + pr4.s_dst) + ;
    pr4.s_npilot + pr4.s_empty - (aisoms.e_mee+aisoms.e_ekmp)
  ELSE 
   m.koplate = sum_prin
  ENDIF 
  m.koplate =  IIF(m.koplate>=0, m.koplate, 0)
  
  tot_k_opl    = tot_k_opl + m.koplate

  tot_s_pred   = tot_s_pred   + s_pred
  tot_sum_flk  = tot_sum_flk  + sum_flk
  tot_sum_prin = tot_sum_prin + sum_prin

  oTable.Cell(npp+1,1).Select
  oWord.Selection.TypeText(STR(npp,3))
  oTable.Cell(npp+1,2).Select
  oWord.Selection.TypeText(STR(lpuid,4))
  oTable.Cell(npp+1,3).Select
  oWord.Selection.TypeText(mcod)
  oTable.Cell(npp+1,4).Select
  oWord.Selection.TypeText(lpuname)
  oTable.Cell(npp+1,5).Select
  oWord.Selection.TypeText(TRANSFORM(s_pred, '99999999.99'))
  oTable.Cell(npp+1,6).Select
  oWord.Selection.TypeText(TRANSFORM(0, '99999999.99'))
  oTable.Cell(npp+1,7).Select
  oWord.Selection.TypeText(TRANSFORM(sum_flk, '99999999.99'))
  oTable.Cell(npp+1,8).Select
  oWord.Selection.TypeText(TRANSFORM(sum_prin, '99999999.99'))
  oTable.Cell(npp+1,9).Select
  oWord.Selection.TypeText(TRANSFORM(m.koplate, '99999999.99'))
  
  oWord.Selection.InsertRowsBelow
  npp = npp + 1
 ENDSCAN 
 
 SET RELATION OFF INTO pr4 

 IF USED('pr4')
  USE IN pr4
 ENDIF 

 oTable.Cell(npp+2,2).Select
 oWord.Selection.TypeText(TRANSFORM(tot_s_pred, '999999999.99'))
 oTable.Cell(npp+2,3).Select
 oWord.Selection.TypeText(TRANSFORM(0, '99999999.99'))
 oTable.Cell(npp+2,4).Select
 oWord.Selection.TypeText(TRANSFORM(tot_sum_flk, '99999999.99'))
 oTable.Cell(npp+2,5).Select
 oWord.Selection.TypeText(TRANSFORM(tot_sum_prin, '999999999.99'))
 oTable.Cell(npp+2,6).Select
 oWord.Selection.TypeText(TRANSFORM(m.tot_k_opl, '999999999.99'))

 oDoc.SaveAs(DocName, 0)
 oWord.Visible = .t.

RETURN 
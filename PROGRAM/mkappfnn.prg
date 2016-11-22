PROCEDURE mkAppFnN(mcod,m.IsVisible,m.IsQuit)
 m.mcod = mcod 
 m.lpuid = lpuid
 m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
 DotName = ptempl+'\ActPFnN.xlt'
 m.mmy    = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)
 DocName = pbase+'\'+m.gcperiod+'\'+m.mcod+'\pdf'+m.qcod+m.mmy
 
 IF !fso.FileExists(dotname)
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ШАБЛОН ДОКУМЕНТА ActPFnN.xlt'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 oal = ALIAS()
 IF OpenFile(pcommon+'\svod082016', 'svod', 'shar', 'lpuid')>0
  IF USED('svod')
   USE IN svod
  ENDIF 
  SELECT (oal)
  RETURN 
 ENDIF 
 
 m.val02 = IIF(SEEK(m.lpuid, 'svod'), svod.feb, 0)
 m.val03 = IIF(SEEK(m.lpuid, 'svod'), svod.mar, 0)
 m.val04 = IIF(SEEK(m.lpuid, 'svod'), svod.apr, 0)
 m.val05 = IIF(SEEK(m.lpuid, 'svod'), svod.may, 0)
 m.delta = m.val02+m.val03+m.val04+m.val05
 
 WAIT "ЗАПУСК EXCEL..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 oDoc = oExcel.WorkBooks.Add(dotname)
 
 m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ;
  ALLTRIM(sprlpu.fullname)+', '+m.mcod+' ('+CokrName(VAL(sprlpu.cokr))+')', '')
 IF m.IsLpuTpn
  m.s_avans = IIF(SEEK(m.mcod, 'aisoms'), aisoms.s_pr_avans, 0)
 ELSE 
  m.s_avans = IIF(SEEK(m.mcod, 'aisoms'), aisoms.s_avans, 0)
 ENDIF 
 m.e_mee   = IIF(SEEK(m.mcod, 'aisoms'), aisoms.e_mee, 0)
 m.e_ekmp  = IIF(SEEK(m.mcod, 'aisoms'), aisoms.e_ekmp, 0)

 m.finval = finval
 m.str01 = m.finval + m.delta
 m.str02 = m.s_avans 
 m.str03 = s_others
 m.str04 = s_guests
 m.str05 = s_npilot + s_empty
 m.str06 = m.e_mee+m.e_ekmp
 
* m.str10 = aisoms.s_pred - IIF(m.IsLpuTpn, aisoms.pf_flk, aisoms.sum_flk)
 
 m.koplate = m.str01 - m.str02 - m.str03 + m.str04 + m.str05 - m.str06
* m.str10 = IIF(m.IsLpuTpn, (aisoms.sum_kr + aisoms.s_dop) - aisoms.pf_flk, aisoms.s_pred - aisoms.sum_flk)
 m.str10 = aisoms.s_pred - aisoms.sum_flk
 
 WITH oExcel
  .Cells(4,1) = m.lpuname
  .Cells(7,1)='со СМО '+ALLTRIM(m.qname)
  .Cells(9,1)='за  '+NameOfMonth(tMonth)+' '+STR(tYear,4)+' г.'

  .Cells(12,8).Value = m.str01
  .Cells(13,8).Value = m.str02
  .Cells(14,8).Value = m.str03
  .Cells(15,8).Value = m.str04
  .Cells(16,8).Value = m.str05
  .Cells(17,8).Value = m.str06

  .Cells(29,1)=ALLTRIM(m.qname)
 ENDWITH 

 oExcel.Cells(18,1)='Итого к оплате: (п.1-п.2-п.3+п.4+п.5-п.6)'+;
  TRANSFORM(m.koplate, '99 999 999.99')+' руб.'
 oExcel.Cells(20,1)='(с учетом результатов МЭК) '+;
  TRANSFORM(m.str10,'99999999.99')+' руб.'

 oExcel.Cells(22,1)='Расчетный объем подушевого финансирования за август 2016: '+;
  TRANSFORM(m.finval,'99999999.99')+' руб.'

 oExcel.Cells(25,1)='за февраль 2016 года на '+;
  TRANSFORM(m.val02,'99999.99')+' руб.'
 oExcel.Cells(26,1)='за март 2016 года на '+;
  TRANSFORM(m.val03,'99999.99')+' руб.'
 oExcel.Cells(27,1)='за апрель 2016 года на '+;
  TRANSFORM(m.val04,'99999.99')+' руб.'
 oExcel.Cells(28,1)='за май 2016 года на '+;
  TRANSFORM(m.val05,'99999.99')+' руб.'
 oExcel.Cells(29,1)='Итого: '+;
  TRANSFORM(m.val02+m.val03+m.val04+m.val05,'999999.99')+' руб.'


 IF fso.FileExists(docname+'.xls')
  fso.DeleteFile(docname+'.xls')
 ENDIF 
 oDoc.SaveAs(DocName, 18)
 TRY 
  IF fso.FileExists(docname+'.pdf')
   fso.DeleteFile(docname+'.pdf')
  ENDIF 
  oDoc.SaveAs(DocName, 57)
 CATCH 
 ENDTRY 
 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
*  oExcel.Interactive= .F. 
 ELSE 
  oDoc.Close(0)
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 
 
 USE IN svod
 SELECT (oal)

RETURN 
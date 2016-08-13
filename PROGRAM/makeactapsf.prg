PROCEDURE MakeActAPSF

 m.IsVisible = .t. 
 m.IsQuit    = .f.

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ ÀÊÒ ÑÂÅÐÊÈ ÐÀÑ×ÅÒÎÂ?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 m.mcod  = mcod
 m.lpuid = lpuid
 m.cokr  = cokr
 m.cokrname = IIF(SEEK(m.cokr, 'admokr'), ALLTRIM(admokr.name_okr), '')
 
 m.lIsPilot = IIF(SEEK(m.lpuid, 'pilot'), .T., .F.)
 m.IsLpuTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)

 dotname = ptempl+'\aktapsf.xlt'
 docname = pout+'\as'+mcod+m.gcperiod

 IF !fso.FileExists(dotname)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ AKTAPSF.XLT'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF fso.FileExists(docname+'.xls')
  fso.DeleteFile(docname+'.xls')
 ENDIF 

 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 
 WITH oExcel
  .ReferenceStyle= -4150  && xlR1C1
  .SheetsInNewWorkbook = 1
 ENDWITH 
 
 m.sum_flk = sum_flk 
 m.defsum = IIF(!m.lIsPilot, sum_flk + e_mee + e_ekmp, e_mee + e_ekmp)

 IF m.IsLpuTpn
  m.s_avans = IIF(SEEK(m.mcod, 'aisoms'), aisoms.s_pr_avans, 0)
 ELSE 
  m.s_avans = IIF(SEEK(m.mcod, 'aisoms'), aisoms.s_avans, 0)
 ENDIF 

 m.stroka1  = dolg_b+m.s_avans && (13,52)

 IF m.lIsPilot
*  m.str01 = IIF(SEEK(m.lpuid, 'pr4'), pr4.adnorm*pr4.adults + pr4.chnorm*pr4.childs, 0)
  m.str01 = IIF(SEEK(m.lpuid, 'pr4'), pr4.finval, 0)
  m.str33 = IIF(SEEK(m.lpuid, 'pr4'), pr4.s_own - m.s_avans, 0)
  m.str31 = IIF(SEEK(m.lpuid, 'pr4'), pr4.s_others, 0)
  m.str32 = IIF(SEEK(m.lpuid, 'pr4'), pr4.s_guests, 0)
  m.str04 = IIF(SEEK(m.lpuid, 'pr4'), pr4.s_kompl + pr4.s_dst, 0)
  m.str06 = IIF(SEEK(m.lpuid, 'pr4'), pr4.s_npilot, 0)
  m.str07 = IIF(SEEK(m.lpuid, 'pr4'), pr4.s_empty, 0)
  m.koplate = m.str01 - m.s_avans - m.str31 + m.str32 + m.str04 + m.str06 + m.str07 - (e_mee + e_ekmp)

  m.stroka2  = m.str01 - m.str31 + m.str32 + m.str04 + m.str06 + m.str07 && (14,52)
 ELSE 
  m.stroka2  = s_pred         && (14,52)
 ENDIF 

 m.stroka3  = m.defsum       && (15,52)
 m.stroka31 = IIF(!m.lIsPilot, sum_flk, 0)        && (16,52)
 m.stroka32 = e_mee          && (17,52)
 m.stroka33 = e_ekmp         && (18,52)
 m.stroka4  = s_avans2+IIF(s_pred-m.defsum-m.s_avans-dolg_b>0, s_pred-m.defsum-m.s_avans-dolg_b,0)
 m.stroka41 = s_avans2
 IF m.lIsPilot
  m.stroka42 = IIF(m.koplate>=0, m.koplate, 0)
*  m.stroka5  = s_avans2 + IIF(m.koplate<0, -1*m.koplate, 0)
*  m.stroka5  = s_avans2 + m.koplate
  m.stroka5  = -1*m.stroka1 + m.stroka2 - m.defsum - (m.stroka41+m.stroka42)
 ELSE 
  m.stroka42 = IIF(s_pred-m.defsum-m.s_avans-dolg_b>0, s_pred-m.defsum-m.s_avans-dolg_b,0)
  m.stroka5  = m.stroka2 - m.stroka3 - m.stroka4 - m.stroka1
 ENDIF 

 
 oDoc = oExcel.WorkBooks.Add(dotname)
 
 WITH oExcel
   
  .Cells(04,36).Value  = IIF(tMonth<=10, tMonth+2, tMonth-10)
  .Cells(04,49).Value  = '1 '+LOWER(NameOfMonth2(MONTH(GOMONTH(m.tdat1,2))))
*  .Cells(04,74).Value  = RIGHT(STR(YEAR(GOMONTH(m.tdat1,2)-1),4),2)
  .Cells(04,74).Value  = RIGHT(STR(IIF(tMonth<=10, tYear, tYear+1),4),2)
  .Cells(07,47).Value  = '1 '+LOWER(NameOfMonth2(MONTH(GOMONTH(m.tdat1,2))))
  .Cells(07,74).Value  = RIGHT(STR(IIF(tMonth<=10, tYear, tYear+1),4),2)
  .Cells(09,01).Value  = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
  .Cells(09,42).Value  = m.cokrname

  m.t_cell = .Cells(14,13).Value
  m.t_cell = ALLTRIM(m.t_cell) + ' ' + LOWER(NameOfMonth(tMonth))
  .Cells(14,13).Value = m.t_cell

  .Cells(13,52).Value  = TRANSFORM(-1*m.stroka1,'99999999.99')
  .Cells(13,79).Value  = TRANSFORM(-1*m.stroka1,'99999999.99')
  
  .Cells(14,52).Value  = TRANSFORM(m.stroka2,'99999999.99')
  .Cells(14,79).Value  = TRANSFORM(m.stroka2,'99999999.99')

  .Cells(15,52).Value  = TRANSFORM(m.defsum,'99999999.99')
  .Cells(15,79).Value  = TRANSFORM(m.defsum,'99999999.99')

  .Cells(16,52).Value  = TRANSFORM(m.stroka31,'99999999.99')
  .Cells(16,79).Value  = TRANSFORM(m.stroka31,'99999999.99')

  .Cells(17,52).Value  = TRANSFORM(m.stroka32,'99999999.99')
  .Cells(17,79).Value  = TRANSFORM(m.stroka32,'99999999.99')

  .Cells(18,52).Value  = TRANSFORM(m.stroka33,'99999999.99')
  .Cells(18,79).Value  = TRANSFORM(m.stroka33,'99999999.99')

  .Cells(19,52).Value  = TRANSFORM(IIF(m.lIsPilot, m.stroka41+m.stroka42, m.stroka4),'99999999.99')
  .Cells(19,79).Value  = TRANSFORM(IIF(m.lIsPilot, m.stroka41+m.stroka42, m.stroka4),'99999999.99')

  .Cells(20,52).Value  = TRANSFORM(m.stroka41,'99999999.99')
  .Cells(20,79).Value  = TRANSFORM(m.stroka41,'99999999.99')

  .Cells(21,52).Value  = TRANSFORM(m.stroka42,'99999999.99')
  .Cells(21,79).Value  = TRANSFORM(m.stroka42,'99999999.99')

  .Cells(22,52).Value  = TRANSFORM(m.stroka5,'99999999.99')
  .Cells(22,79).Value  = TRANSFORM(m.stroka5,'99999999.99')
  
  .Cells(23,1).Value  = 'Ñïðàâî÷íî, çàáðàêîâàíî ïî ÌÝÊ: '+TRANSFORM(m.sum_flk, '99999999.99')

ENDWITH 

 oDoc.SaveAs(DocName)
 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  oDoc.Close(0)
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 
 
RETURN 
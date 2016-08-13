PROCEDURE mkAppF(mcod,m.IsVisible,m.IsQuit)
 m.mcod = mcod 
 m.lpuid = lpuid
 DotName = ptempl+'\ActPF.xlt'
* DocName = pbase+'\'+m.gcperiod+'\'+m.mcod+'\apf'+m.mcod
 m.mmy    = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)
 DocName = pbase+'\'+m.gcperiod+'\'+m.mcod+'\pdf'+m.qcod+m.mmy
 
* m.IsVisible = .t.
* m.IsQuit    = .f.

 IF !fso.FileExists(dotname)
  MESSAGEBOX(CHR(13)+CHR(10)+'ќ“—”“—“¬”≈“ ЎјЅЋќЌ ƒќ ”ћ≈Ќ“ј ActPF.xlt'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 WAIT "«јѕ”—  EXCEL..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 oDoc = oExcel.WorkBooks.Add(dotname)
 
 m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 
 m.s_avans = IIF(SEEK(m.mcod, 'aisoms'), aisoms.s_avans, 0)
 m.e_mee   = IIF(SEEK(m.mcod, 'aisoms'), aisoms.e_mee, 0)
 m.e_ekmp  = IIF(SEEK(m.mcod, 'aisoms'), aisoms.e_ekmp, 0)

* m.str01 = adnorm*adults + chnorm*childs
 m.str01 = finval
 m.str02 = m.s_avans 
 m.str33 = s_own - m.str02
 m.str31 = s_others
 m.str32 = s_guests
 m.str04 = s_kompl + s_dst
 m.str06 = s_npilot
 m.str07 = s_empty
 m.str03 = m.str02 - m.str31 + m.str32 + m.str33
 m.str05 = m.str01 - IIF(m.str03<0, -1*m.str03, m.str03)
 m.str08 = m.e_mee+m.e_ekmp
 m.str09 = IIF(m.str05>=0, m.str05, 0)+m.str04+m.str06+m.str07-m.str08
 
 m.str10 = aisoms.s_pred - aisoms.sum_flk
 
 m.koplate = m.str01 - m.str02 - m.str31 + m.str32 + m.str04 + m.str06 + m.str07 - m.str08 
 
 WITH oExcel
  .Cells(4,1) = m.lpuname
  .Cells(7,1)='со —ћќ '+ALLTRIM(m.qname)
  .Cells(9,1)='за  '+NameOfMonth(tMonth)+' '+STR(tYear,4)+' г.'

  .Cells(12,8).Value = m.str01
  .Cells(13,8).Value = m.str02
  .Cells(14,8).Value = m.str03
  .Cells(15,8).Value = m.str31
  .Cells(16,8).Value = m.str32
  .Cells(17,8).Value = m.str33
  .Cells(18,8).Value = m.str04
  .Cells(19,8).Value = m.str05
  .Cells(20,8).Value = m.str06
  .Cells(21,8).Value = m.str07
  .Cells(22,8).Value = m.str08
  .Cells(23,8).Value = m.str09

  .Cells(29,1)=ALLTRIM(m.qname)
 ENDWITH 

 oExcel.Cells(24,1)='»того к оплате: (п.1-п.2-п.3.1+п.3.2+п.4+п.6+п.7-п.8)'+;
  TRANSFORM(m.koplate, '99 999 999.99')+' руб.'
 oExcel.Cells(26,1)='(с учетом результатов ћЁ ) '+;
  TRANSFORM(m.str10,'99999999.99')+' руб.'

 IF fso.FileExists(docname+'.xls')
  fso.DeleteFile(docname+'.xls')
 ENDIF 
 oDoc.SaveAs(DocName)
 TRY 
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
 

RETURN 
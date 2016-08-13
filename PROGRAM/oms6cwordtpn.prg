FUNCTION oms6cwordtpn(lcPath, IsVisible, IsQuit)
 
 =OpenFile(pcommon+'\smo', 'smo', 'shar', 'code')
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx', 'sprcokr', 'shar', 'cokr')
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\tarifn', 'tarif', 'shar', 'cod')
 IF !USED('sprlpu')
  m.WasUsedSprLpu = .f.
  =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "lpu_id")
 ELSE 
  m.WasUsedSprLpu = .t.
 ENDIF 

 SELECT AisOms


 m.mmy        = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
 m.mcod       = mcod
 m.lpuid      = lpuid
 m.lpuname    = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 m.cokr     = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.cokr), '')
 m.cokr_name = IIF(SEEK(m.cokr, 'sprcokr'), ALLTRIM(sprcokr.name_okr), '')
 m.smoname    = IIF(SEEK(m.qcod, 'smo'), ALLTRIM(smo.fullname), '')
 m.smonames   = IIF(SEEK(m.qcod, 'smo'), ALLTRIM(smo.name), '')
 m.arcfname   = 'b'+m.mcod+'.'+m.mmy
 m.message_id = ALLTRIM(cmessage)
 m.datpriemki = TTOC(Recieved)
 
 m.period = ' '+NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))+ ' '+SUBSTR(m.gcperiod,1,4)
 
 m.kol_paz = paz
 m.kol_sch = 0
 m.summa   = s_pred

 poi_file   = fso.GetFile(lcPath + '\' + arcfname)
 m.arcfdate = poi_file.DateLastModified
 
 ZipItemCount = 5

 DotName = pTempl + "\Prqqmmy_TPN.dot"
 DocName = lcPath + "\Pr" + LOWER(m.qcod) + PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 
 eeFile = 'e'+m.mcod
 =OpenFile(lcpath+'\Talon', 'talon', 'shar')
* USE &lcPath\Talon  IN 0 ALIAS Talon SHARED 
 =OpenFile(lcpath+'\People', 'people', 'shar', 'sn_pol')
* USE &lcPath\People IN 0 ALIAS people SHARED orde sn_pol 
 =OpenFile(lcpath+'\'+eeFile, 'serror', 'shar', 'rid')
* USE &lcPath\&eeFile IN 0 ALIAS sError SHARED ORDER rid 
 =OpenFile(lcpath+'\'+eeFile, 'rerror', 'shar', 'rrid', 'again')
* USE &lcPath\&eeFile IN 0 ALIAS rError SHARED ORDER rrid AGAIN 

 m.pazprikl   = 0
 m.pazneprikl = 0
 m.pazpriklok   = 0
 m.paznepriklok = 0

 SELECT people
 SET RELATION TO RecId INTO rError
 SCAN
  m.d_type = d_type
  m.IsPrikl = IIF(INLIST(m.d_type,'b','f','g','h','i','j'), .t., .f.)
  m.pazprikl   = m.pazprikl + IIF(m.IsPrikl,1,0)
  m.pazneprikl = m.pazneprikl + IIF(m.IsPrikl,0,1)
  IF EMPTY(rError.rid)
   m.pazpriklok   = m.pazpriklok + IIF(m.IsPrikl,1,0)
   m.paznepriklok = m.paznepriklok + IIF(m.IsPrikl,0,1)
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO rError
 USE IN rError

 SELECT Talon 
 SET RELATION TO sn_pol INTO people
 SET RELATION TO RecId INTO sError ADDITIVE 
 SET RELATION TO cod INTO tarif ADDITIVE 

 m.accsprikl     = 0 
 m.accsneprikl   = 0
 m.accspriklok   = 0 
 m.accsnepriklok = 0
 m.sumprikl = 0
 m.sumneprikl = 0
 m.sumpriklok = 0
 m.sumnepriklok = 0
 
 m.priklpf = 0
 m.nepriklpf = 0
 m.priklt  = 0
 m.nepriklt  = 0
 m.priklothers  = 0
 m.nepriklothers  = 0
 
 
 SCAN
  m.d_type = people.d_type
  m.IsPrikl = IIF(INLIST(m.d_type,'b','f','g','h','i','j'), .t., .f.)
  m.accsprikl     = m.accsprikl + IIF(m.IsPrikl,1,0)
  m.accsneprikl   = m.accsneprikl + IIF(m.IsPrikl,0,1)
  
  m.sumprikl     = m.sumprikl + IIF(m.IsPrikl,s_all,0)
  m.sumneprikl   = m.sumneprikl + IIF(m.IsPrikl,0,s_all)
  
  IF EMPTY(sError.rid)
   m.accspriklok   = m.accspriklok + IIF(m.IsPrikl,1,0)
   m.accsnepriklok = m.accsnepriklok + IIF(m.IsPrikl,0,1)
   m.sumpriklok   = m.sumpriklok + IIF(m.IsPrikl,s_all,0)
   m.sumnepriklok = m.sumnepriklok + IIF(m.IsPrikl,0,s_all)
   DO CASE 
    CASE tarif.tpn = 'p' && ¿œœ (ÔÙ)
     m.priklpf   = m.priklpf + IIF(m.IsPrikl,s_all,0)
     m.nepriklpf = m.nepriklpf + IIF(m.IsPrikl,0, s_all)
     
    CASE tarif.tpn = 't' && ¿œœ (Ú)
     m.priklt   = m.priklt + IIF(m.IsPrikl,s_all,0)
     m.nepriklt = m.nepriklt + IIF(m.IsPrikl,0, s_all)
    OTHERWISE 

     m.priklothers   = m.priklothers + IIF(m.IsPrikl,s_all,0)
     m.nepriklothers = m.nepriklothers + IIF(m.IsPrikl,0, s_all)
   ENDCASE 
  ENDIF 

 ENDSCAN  

 SET RELATION OFF INTO people
 SET RELATION OFF INTO sError
 SET RELATION OFF INTO tarif
 USE 
 USE IN sError
 USE IN people 
 
 USE IN smo 
 USE IN sprcokr
 IF m.WasUsedSprLpu = .f.
  USE IN SprLpu
 ENDIF 
 USE IN tarif

 SELECT AisOms

 TRY 
  oWord=GETOBJECT(,"Word.Application")
 CATCH 
  oWord=CREATEOBJECT("Word.Application")
 ENDTRY 

 oDoc = oWord.Documents.Add(dotname)
 
 oDoc.Bookmarks('DatPriemki').Select  
 oWord.Selection.TypeText(m.datpriemki)
 oDoc.Bookmarks('SmoName').Select  
 oWord.Selection.TypeText(m.smoname)
 oDoc.Bookmarks('LpuName').Select  
 oWord.Selection.TypeText(m.lpuname+', '+m.cokr_name+', '+m.mcod)
 oDoc.Bookmarks('Period').Select  
 oWord.Selection.TypeText(m.period)

 oDoc.Bookmarks('pazprikl').Select  
 oWord.Selection.TypeText(TRANSFORM(m.pazprikl,'999999'))
 oDoc.Bookmarks('pazneprikl').Select  
 oWord.Selection.TypeText(TRANSFORM(m.pazneprikl,'999999'))
 oDoc.Bookmarks('paztot').Select  
 oWord.Selection.TypeText(TRANSFORM(m.pazprikl+m.pazneprikl,'999999'))

 oDoc.Bookmarks('accsprikl').Select  
 oWord.Selection.TypeText(TRANSFORM(m.accsprikl,'999999'))
 oDoc.Bookmarks('accsneprikl').Select  
 oWord.Selection.TypeText(TRANSFORM(m.accsneprikl,'999999'))
 oDoc.Bookmarks('accstot').Select  
 oWord.Selection.TypeText(TRANSFORM(m.accsprikl+m.accsneprikl,'999999'))

 oDoc.Bookmarks('sumprikl').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sumprikl,'99999999.99'))
 oDoc.Bookmarks('sumneprikl').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sumneprikl,'99999999.99'))
 oDoc.Bookmarks('sumtot').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sumprikl+m.sumneprikl,'99999999.99'))

 oDoc.Bookmarks('pazpriklok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.pazpriklok,'999999'))
 oDoc.Bookmarks('paznepriklok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.paznepriklok,'999999'))
 oDoc.Bookmarks('paztotok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.pazpriklok+m.paznepriklok,'999999'))

 oDoc.Bookmarks('sumpriklok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sumpriklok,'99999999.99'))
 oDoc.Bookmarks('sumnepriklok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sumnepriklok,'99999999.99'))
 oDoc.Bookmarks('sumtotok').Select  
 oWord.Selection.TypeText(TRANSFORM(m.sumpriklok+m.sumnepriklok,'99999999.99'))

 oDoc.Bookmarks('priklpf').Select  
 oWord.Selection.TypeText(TRANSFORM(m.priklpf,'99999999.99'))
 oDoc.Bookmarks('nepriklpf').Select  
 oWord.Selection.TypeText(TRANSFORM(m.nepriklpf,'99999999.99'))
 oDoc.Bookmarks('totpf').Select  
 oWord.Selection.TypeText(TRANSFORM(m.priklpf+m.nepriklpf,'99999999.99'))

 oDoc.Bookmarks('priklt').Select  
 oWord.Selection.TypeText(TRANSFORM(m.priklt,'99999999.99'))
 oDoc.Bookmarks('nepriklt').Select  
 oWord.Selection.TypeText(TRANSFORM(m.nepriklt,'99999999.99'))
 oDoc.Bookmarks('tott').Select  
 oWord.Selection.TypeText(TRANSFORM(m.priklt+m.nepriklt,'99999999.99'))

 oDoc.Bookmarks('priklothers').Select  
 oWord.Selection.TypeText(TRANSFORM(m.priklothers,'99999999.99'))
 oDoc.Bookmarks('nepriklothers').Select  
 oWord.Selection.TypeText(TRANSFORM(m.nepriklothers,'99999999.99'))
 oDoc.Bookmarks('totothers').Select  
 oWord.Selection.TypeText(TRANSFORM(m.priklothers+m.nepriklothers,'99999999.99'))

 oDoc.Bookmarks('arcfname').Select  
 oWord.Selection.TypeText(m.arcfname)
 oDoc.Bookmarks('arcfdate').Select  
 oWord.Selection.TypeText(DTOC(m.arcfdate))
 oDoc.Bookmarks('KolVloz').Select  
 oWord.Selection.TypeText(ALLTRIM(STR(m.ZipItemCount)))

 TRY 
  oDoc.SaveAs(DocName, 17)
 CATCH 
*  MESSAGEBOX(CHR(13)+CHR(10)+;
   '—Œ’–¿Õ≈Õ»≈ ¬ PDF-‘Œ–Ã¿“≈'+CHR(13)+CHR(10)+;
   'Õ≈ œŒƒƒ≈–∆»¬¿≈“—ﬂ ”—“¿ÕŒ¬À≈ÕÕŒ…'+CHR(13)+CHR(10)+;
   '¬≈–—»≈… WORD!',0+64,'')
 ENDTRY 

 oDoc.SaveAs(DocName, 0)
 
 IF IsVisible == .t. 
  oWord.Visible = .t.
 ELSE 
  oDoc.Close(0)
  IF IsQuit
   oWord.Quit
  ENDIF 
 ENDIF 
 
RETURN  


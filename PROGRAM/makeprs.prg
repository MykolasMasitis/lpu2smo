FUNCTION MakePrs(para1)
 PRIVATE mcod,mmy,lpu_id,IsLpuTpn
 m.mcod   = para1
 m.mmy    = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 m.lpu_id = lpuid
 m.IsLpuTpn = IIF(SEEK(m.lpu_id, 'lputpn'), .t., .f.)
 m.IsPilot = IIF(SEEK(m.mcod, 'pilot'), .t., .f.)

 m.sum_st1 = 0  && Сумма к оплате, полученная вычитанием ФЛК из представленных
 m.sum_st1 = s_pred - sum_flk

 m.nIsDoubles = 0

 lcPath = pBase+'\'+gcperiod+'\'+mcod
 IF !fso.FolderExists(lcPath)
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ДИРЕКТОРИЯ ЛПУ!'+CHR(13)+CHR(10),0+64,mcod)
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN', 'tarif', 'shar', 'cod')>0
  IF USED('TARIF')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', 'sookod', 'shar', 'er_c')>0
  IF USED('TARIF')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 
 
 =MakeYFilesOne(lcPath)
 
 USE IN sookod
 USE IN tarif
 SELECT aisoms

 
 MESSAGEBOX("ОБРАБОТКА ЗАКОНЧЕНА!",0+64,"")
 
RETURN 
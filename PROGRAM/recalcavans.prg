PROCEDURE RecalcAvans
 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш унрхре оепеявхрюрэ юбюмяхпнбюмхе?'+CHR(13)+CHR(10), '')==7
  RETURN 
 ENDIF 

 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш сбепемш я ябнху деиярбхъу?'+CHR(13)+CHR(10), '')==7
  RETURN 
 ENDIF 

 IF tMonth>1
  m.pperiod = STR(tYear,4)+PADL(tMonth-1,2,'0')
 ELSE 
  m.pperiod = STR(tYear-1,4)+PADL(12,2,'0')
 ENDIF 
 m.lIsUsedPeriod = .F.
 IF fso.FileExists(pBase+'\'+pperiod+'\aisoms.dbf')
  IF OpenFile(pbase+'\'+pperiod+'\aisoms', 'aispr', 'shar', 'lpuid')>0
   IF USED('aispr')
    USE IN aispr
   ENDIF 
  ELSE 
   m.lIsUsedPeriod = .T.
  ENDIF 
 ENDIF 

 IF m.lIsUsedPeriod = .F.
  MESSAGEBOX(CHR(13)+CHR(10)+'мебнглнфмн нрйпшрэ опедшдсыхи оепхнд!'+CHR(13)+CHR(10),0+64,m.pperiod)
  RETURN 
 ENDIF 

 wasrec = RECNO()

 IF m.lIsUsedPeriod = .T.
  SELECT aispr
  SCAN
   m.lpu_id = lpuid
*   m.dolg = IIF(s_pred - sum_flk -  (e_mee + e_ekmp) - s_avans - dolg_b<0, ;
   -(s_pred - sum_flk - (e_mee + e_ekmp) - s_avans - dolg_b), 0)

   m.s_avans    = s_avans2
   m.s_pr_avans = pr_avans2
   
   IF m.s_avans=0 AND m.s_pr_avans=0
    LOOP 
   ENDIF 
   
   IF SEEK(m.lpu_id, 'aisoms', 'lpuid')
*    m.isdolg = aisoms.dolg_b
*    REPLACE dolg_b WITH m.isdolg+m.dolg IN aisoms
     REPLACE s_avans WITH m.s_avans, s_pr_avans WITH m.s_pr_avans IN aisoms
   ELSE
    m.headid = 0
    SELECT lpu_id as lpu_id FROM sprlpu WHERE fil_id=m.lpu_id AND lpu_id!=fil_id INTO CURSOR curabc
    IF _tally==1
     m.headid = lpu_id
     MESSAGEBOX(STR(m.headid,4),0+64,'')
    ENDIF 
    USE IN curabc
    SELECT aispr
    IF m.headid>0
     IF SEEK(m.headid, 'aisoms', 'lpuid')
*      m.isdolg = aisoms.dolg_b
*      REPLACE dolg_b WITH m.isdolg+m.dolg, s_avans WITH m.s_avans, s_pr_avans WITH m.s_pr_avans IN aisoms
*      REPLACE dolg_b WITH m.isdolg+m.dolg IN aisoms
      REPLACE s_avans WITH m.s_avans, s_pr_avans WITH m.s_pr_avans IN aisoms
     ENDIF 
    ENDIF 
   ENDIF 
  ENDSCAN 
 ENDIF 

 IF m.lIsUsedPeriod = .T.
  USE IN aispr
 ENDIF 

 SELECT AisOms
 GO (wasrec)

RETURN 
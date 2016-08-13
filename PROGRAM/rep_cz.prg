PROCEDURE rep_cz
 IF OpenFile("&pBase\&gcPeriod\aisoms", "aisoms", "shar", "mcod") > 0
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "mcod") > 0
  USE IN aisoms
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spi_cz', "spi_cz", "shar", "lpu_id") > 0
  USE IN aisoms
  USE IN sprlpu
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spi_cz_ch', "spi_czch", "shar", "lpu_id") > 0
  USE IN spi_cz
  USE IN aisoms
  USE IN sprlpu
  RETURN 
 ENDIF 

 PUBLIC oExcel AS Excel.Application
 WAIT "Запуск MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 m.period = NameOfMonth(tMonth)+ ' '+STR(tYear,4)
 ddat1 = CTOD('01.' + PADL(tMonth,2,'0') + '.' + STR(tYear,4))

 DotName = pTempl + '\Protokol.xlt'
 DocName = pOut + '\Protokol_cz'

 oBook   = oExcel.WorkBooks.Add(dotname)
 oSheet1 = oBook.WorkSheets('CZ1')
 oSheet2 = oBook.WorkSheets('CZ2')
 
 oSheet1.Select
 oSheet1.Cells(2,1).Value2 = 'за ' + m.period + ' г.'
 oSheet1.Cells(3,7).Value2 = m.qname
 
 oSheet2.Select
 oSheet2.Cells(2,1).Value2 = 'за ' + m.period + ' г.'
 oSheet2.Cells(3,8).Value2 = m.qname

 SELECT AisOms
 SET RELATION TO mcod INTO sprlpu
 IF fso.FileExists(pBin+'\aisoms.idx')
  fso.DeleteFile(pBin+'\aisoms.idx')
 ENDIF 
 INDEX ON sprlpu.cokr+mcod TO &pBin\aisoms
 SET INDEX TO &pBin\aisoms

 m.num_lpu1 = 0
 m.num_lpu2 = 0

 m.paz_tot1 = 0
 m.perv_paz_tot1 = 0
 m.usl_tot1 = 0
 m.perv_usl_tot1 = 0
 m.sum_tot1 = 0
 m.sum_perv_tot1 = 0
 m.paz_prin_tot1 = 0
 m.perv_paz_prin_tot1 = 0
 m.usl_prin_tot1 = 0
 m.perv_usl_prin_tot1 = 0
 m.usl_prin_tot1 = 0
 m.perv_usl_prin_tot1 = 0
 m.sum_prin_tot1 = 0
 m.sum_perv_prin_tot1 = 0

 m.paz_tot2 = 0
 m.perv_paz_tot2 = 0
 m.usl_tot2 = 0
 m.perv_usl_tot2 = 0
 m.sum_tot2 = 0
 m.sum_perv_tot2 = 0
 m.paz_prin_tot2 = 0
 m.perv_paz_prin_tot2 = 0
 m.usl_prin_tot2 = 0
 m.perv_usl_prin_tot2 = 0
 m.usl_prin_tot2 = 0
 m.perv_usl_prin_tot2 = 0
 m.sum_prin_tot2 = 0
 m.sum_perv_prin_tot2 = 0

 m.paz_tot217 = 0
 m.perv_paz_tot217 = 0
 m.usl_tot217 = 0
 m.perv_usl_tot217 = 0
 m.sum_tot217 = 0
 m.sum_perv_tot217 = 0
 m.paz_prin_tot217 = 0
 m.perv_paz_prin_tot217 = 0
 m.usl_prin_tot217 = 0
 m.perv_usl_prin_tot217 = 0
 m.usl_prin_tot217 = 0
 m.perv_usl_prin_tot217 = 0
 m.sum_prin_tot217 = 0
 m.sum_perv_prin_tot217 = 0

 SCAN FOR !DELETED()
  m.mcod = mcod
  m.lpu_id = lpuid
  
  IF !SEEK(m.lpu_id, 'spi_cz') AND !SEEK(m.lpu_id, 'spi_czch')
   LOOP 
  ENDIF 
  
  WAIT m.mcod WINDOW NOWAIT 
  IF fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\people.dbf') AND ;
     fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon.dbf') AND ;
     fso.FileExists(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   
   IF OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\people', 'people', 'shared', 'sn_pol') > 0
    SELECT aisoms
    LOOP 
   ENDIF 
   IF OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shared') > 0
    USE IN people
    SELECT aisoms
    LOOP 
   ENDIF 
   IF OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'errors', 'shared', 'rid') > 0
    USE IN people 
    USE IN talon 
    SELECT aisoms
    LOOP 
   ENDIF 
   
   SELECT talon 
   COUNT FOR SUBSTR(otd,2,2)='91' AND d_type!='h'
   IF _tally == 0
    USE IN talon 
    USE IN people 
    USE IN errors
    SELECT aisoms 
    LOOP 
   ENDIF 
   
   CREATE CURSOR paz_cz (sn_pol c(25)) 
   SELECT paz_cz
   INDEX ON sn_pol TAG sn_pol
   SET ORDER TO sn_pol

   CREATE CURSOR paz_czok (sn_pol c(25)) 
   SELECT paz_czok
   INDEX ON sn_pol TAG sn_pol
   SET ORDER TO sn_pol
   
   CREATE CURSOR paz_cz_kusl (sn_pol c(25)) 
   SELECT paz_cz_kusl
   INDEX ON sn_pol TAG sn_pol
   SET ORDER TO sn_pol

   CREATE CURSOR paz_cz_kuslok (sn_pol c(25)) 
   SELECT paz_cz_kuslok
   INDEX ON sn_pol TAG sn_pol
   SET ORDER TO sn_pol

   SELECT talon 
   SET RELATION TO sn_pol INTO people
 
   && если нужно считать!

*   m.tip_lpu = IIF(SUBSTR(m.mcod,2,1)='2', 2, 1)

    IF SEEK(m.lpu_id, 'spi_cz')
     m.tip_lpu = 1
    ELSE 
     m.tip_lpu = 2
    ENDIF 
   
   m.lpuname = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.name), '')
   m.cokr    = IIF(SEEK(m.mcod, 'sprlpu'), ALLTRIM(sprlpu.cokr), '')

   IF m.tip_lpu == 2
    m.num_lpu2 = m.num_lpu2 + 1
    oSheet2.Select
    oSheet2.Cells(10+m.num_lpu2,1).value2 = m.num_lpu2
    oSheet2.Cells(10+m.num_lpu2,2).value2 = m.lpuname
    oSheet2.Cells(10+m.num_lpu2,3).value2 = m.mcod
    oSheet2.Cells(10+m.num_lpu2,4).value2 = m.cokr
   ELSE 
    m.num_lpu1 = m.num_lpu1 + 1
    oSheet1.Select
    oSheet1.Cells(10+m.num_lpu1,1).value2 = m.num_lpu1
    oSheet1.Cells(10+m.num_lpu1,2).value2 = m.lpuname
    oSheet1.Cells(10+m.num_lpu1,3).value2 = m.mcod
    oSheet1.Cells(10+m.num_lpu1,4).value2 = m.cokr
   ENDIF 
   
   m.paz = 0
   m.perv_paz = 0
   m.usl = 0
   m.perv_usl = 0
   m.sum = 0
   m.sum_perv = 0
   
   m.paz_prin = 0
   m.perv_paz_prin = 0
   m.usl_prin = 0
   m.perv_usl_prin = 0
   m.sum_prin = 0
   m.sum_perv_prin = 0
   
   m.sum_h = 0
   m.sum_h_perv = 0

   m.paz17 = 0
   m.perv_paz17 = 0
   m.usl17 = 0
   m.perv_usl17 = 0
   m.sum17 = 0
   m.sum_perv17 = 0
   
   m.paz_prin17 = 0
   m.perv_paz_prin17 = 0
   m.usl_prin17 = 0
   m.perv_usl_prin17 = 0
   m.sum_prin17 = 0
   m.sum_perv_prin17 = 0
   
   m.sum_h17 = 0
   m.sum_h_perv17 = 0

   SET RELATION TO RecID INTO errors ADDITIVE 
   SCAN FOR SUBSTR(otd,2,2) == '91'
    Is17 = IIF(BETWEEN((ddat1-people.dr)/365.25, 15, 17), .T., .F.)
    m.sn_pol = sn_pol
    m.cod = cod
    m.k_u = k_u
    m.s_all = s_all
    m.d_type = d_type
    IF m.d_type != 'h'
     IF !SEEK(m.sn_pol, 'paz_cz')
      INSERT INTO paz_cz (sn_pol) VALUES (m.sn_pol)
      m.paz = m.paz + 1
*      m.perv_paz = m.perv_paz + IIF(INLIST(m.cod,15001,115001), 1, 0)
      m.paz17 = m.paz17 + IIF(Is17, 1, 0)
*      m.perv_paz17 = m.perv_paz17 + IIF(INLIST(m.cod,15001,115001), IIF(Is17, 1, 0), 0)
     ENDIF 
     IF INLIST(m.cod,15001,115001)
      IF !SEEK(m.sn_pol, 'paz_cz_kusl')
       INSERT INTO paz_cz_kusl (sn_pol) VALUES (m.sn_pol)
       m.perv_paz = m.perv_paz + 1
       m.perv_paz17 = m.perv_paz17 + IIF(Is17, 1, 0)
      ENDIF 
     ENDIF 
     IF EMPTY(errors.c_err)

      IF !SEEK(m.sn_pol, 'paz_czok')
       INSERT INTO paz_czok (sn_pol) VALUES (m.sn_pol)
       m.paz_prin = m.paz_prin + 1
*       m.perv_paz_prin = m.perv_paz_prin + IIF(INLIST(m.cod,15001,115001), 1, 0)
      ENDIF 

      IF INLIST(m.cod,15001,115001)
       IF !SEEK(m.sn_pol, 'paz_cz_kuslok')
        INSERT INTO paz_cz_kuslok (sn_pol) VALUES (m.sn_pol)
*       m.paz_prin = m.paz_prin + 1
        m.perv_paz_prin = m.perv_paz_prin + 1
       ENDIF 
      ENDIF 

     ENDIF 
     m.usl = m.usl + m.k_u
     m.perv_usl = m.perv_usl + IIF(INLIST(m.cod,15001,115001), m.k_u, 0)
     m.sum = m.sum + m.s_all
     m.sum_perv = m.sum_perv + IIF(INLIST(m.cod,15001,115001), m.s_all, 0)
     m.usl17 = m.usl17 + IIF(Is17, m.k_u, 0)
     m.perv_usl17 = m.perv_usl17 + IIF(INLIST(m.cod,15001,115001), IIF(Is17, m.k_u, 0), 0)
     m.sum17 = m.sum17 + IIF(Is17, m.s_all, 0)
     m.sum_perv17 = m.sum_perv17 + IIF(INLIST(m.cod,15001,115001), IIF(Is17, m.s_all, 0), 0)
     IF EMPTY(errors.c_err)
      m.usl_prin = m.usl_prin + m.k_u
      m.sum_prin = m.sum_prin + m.s_all
      m.usl_prin17 = m.usl_prin17 + IIF(Is17, m.k_u, 0)
      m.sum_prin17 = m.sum_prin17 + IIF(Is17, m.s_all, 0)
      IF INLIST(m.cod, 15001, 115001)
       m.perv_usl_prin = m.perv_usl_prin + m.k_u
       m.sum_perv_prin = m.sum_perv_prin + m.s_all
       m.perv_usl_prin17 = m.perv_usl_prin17 + IIF(Is17, m.k_u, 0)
       m.sum_perv_prin17 = m.sum_perv_prin17 + IIF(Is17, m.s_all, 0)
      ENDIF 
     ENDIF 
    ELSE
     m.sum_h = m.sum_h + m.s_all
     m.sum_h17 = m.sum_h17 + IIF(Is17, m.s_all, 0)
    ENDIF
   ENDSCAN 
   SET RELATION OFF INTO errors
   SET RELATION OFF INTO people 
   USE IN paz_cz
   USE IN paz_cz_kusl
   USE IN paz_czok
   USE IN paz_cz_kuslok
  
   IF m.tip_lpu == 2
    oSheet2.Select
    oSheet2.Cells(10+m.num_lpu2,5).value2  = TRANSFORM(m.paz, '999')
    oSheet2.Cells(10+m.num_lpu2,6).value2  = TRANSFORM(m.paz17, '999')
    oSheet2.Cells(10+m.num_lpu2,7).value2  = TRANSFORM(m.perv_paz, '999')
    oSheet2.Cells(10+m.num_lpu2,8).value2  = TRANSFORM(m.perv_paz17, '999')
    oSheet2.Cells(10+m.num_lpu2,9).value2  = TRANSFORM(m.usl, '9999')
    oSheet2.Cells(10+m.num_lpu2,10).value2 = TRANSFORM(m.usl17, '9999')
    oSheet2.Cells(10+m.num_lpu2,11).value2 = TRANSFORM(m.perv_usl, '9999')
    oSheet2.Cells(10+m.num_lpu2,12).value2 = TRANSFORM(m.perv_usl17, '9999')
    oSheet2.Cells(10+m.num_lpu2,13).value2 = TRANSFORM(m.sum-m.sum_h, '999999.99')
    oSheet2.Cells(10+m.num_lpu2,14).value2 = TRANSFORM(m.sum17-m.sum_h17, '999999.99')
    oSheet2.Cells(10+m.num_lpu2,15).value2 = TRANSFORM(m.sum_perv-m.sum_h, '999999.99')
    oSheet2.Cells(10+m.num_lpu2,16).value2 = TRANSFORM(m.sum_perv17-m.sum_h17, '999999.99')
    oSheet2.Cells(10+m.num_lpu2,17).value2 = TRANSFORM(m.usl_prin, '99999')
    oSheet2.Cells(10+m.num_lpu2,18).value2 = TRANSFORM(m.usl_prin17, '99999')
    oSheet2.Cells(10+m.num_lpu2,19).value2 = TRANSFORM(m.perv_usl_prin, '99999')
    oSheet2.Cells(10+m.num_lpu2,20).value2 = TRANSFORM(m.perv_usl_prin17, '99999')
    oSheet2.Cells(10+m.num_lpu2,21).value2 = TRANSFORM(m.usl_prin, '99999')
    oSheet2.Cells(10+m.num_lpu2,22).value2 = TRANSFORM(m.usl_prin17, '99999')
    oSheet2.Cells(10+m.num_lpu2,23).value2 = TRANSFORM(m.perv_usl_prin, '99999')
    oSheet2.Cells(10+m.num_lpu2,24).value2 = TRANSFORM(m.perv_usl_prin17, '99999')
    oSheet2.Cells(10+m.num_lpu2,25).value2 = TRANSFORM(m.sum_prin-m.sum_h, '999999.99')
    oSheet2.Cells(10+m.num_lpu2,26).value2 = TRANSFORM(m.sum_prin17-m.sum_h17, '999999.99')
    oSheet2.Cells(10+m.num_lpu2,27).value2 = TRANSFORM(m.sum_perv_prin-m.sum_h, '999999.99')
    oSheet2.Cells(10+m.num_lpu2,28).value2 = TRANSFORM(m.sum_perv_prin17-m.sum_h17, '999999.99')

    m.paz_tot2           = m.paz_tot2 + m.paz
    m.perv_paz_tot2      = m.perv_paz_tot2 + m.perv_paz
    m.usl_tot2           = m.usl_tot2 + m.usl
    m.perv_usl_tot2      = m.perv_usl_tot2 + m.perv_usl
    m.sum_tot2           = m.sum_tot2 + m.sum
    m.sum_perv_tot2      = m.sum_perv_tot2 + m.sum_perv
    m.usl_prin_tot2      = m.usl_prin_tot2 + m.usl_prin
    m.perv_usl_prin_tot2 = m.perv_usl_prin_tot2 + m.perv_usl_prin
    m.sum_prin_tot2      = m.sum_prin_tot2 + m.sum_prin
    m.sum_perv_prin_tot2 = m.sum_perv_prin_tot2 + m.sum_perv_prin

    m.paz_tot217           = m.paz_tot217 + m.paz17
    m.perv_paz_tot217      = m.perv_paz_tot217 + m.perv_paz17
    m.usl_tot217           = m.usl_tot217 + m.usl17
    m.perv_usl_tot217      = m.perv_usl_tot217 + m.perv_usl17
    m.sum_tot217           = m.sum_tot217 + m.sum17
    m.sum_perv_tot217      = m.sum_perv_tot217 + m.sum_perv17
    m.usl_prin_tot217      = m.usl_prin_tot217 + m.usl_prin17
    m.perv_usl_prin_tot217 = m.perv_usl_prin_tot217 + m.perv_usl_prin17
    m.sum_prin_tot217      = m.sum_prin_tot217 + m.sum_prin17
    m.sum_perv_prin_tot217 = m.sum_perv_prin_tot217 + m.sum_perv_prin17

    oSheet2.Cells(11+m.num_lpu2,1).Select
    ssel = oExcel.Selection 
    ssel.EntireRow.Insert
   ELSE 
    oSheet1.Select
    oSheet1.Cells(10+m.num_lpu1,5).value2  = TRANSFORM(m.paz, '999')
    oSheet1.Cells(10+m.num_lpu1,6).value2  = TRANSFORM(m.perv_paz, '999')
*    oSheet1.Cells(10+m.num_lpu1,5).value2  = TRANSFORM(m.usl, '999')
*    oSheet1.Cells(10+m.num_lpu1,6).value2  = TRANSFORM(m.perv_usl, '999')

    oSheet1.Cells(10+m.num_lpu1,7).value2  = TRANSFORM(m.usl, '9999')
    oSheet1.Cells(10+m.num_lpu1,8).value2  = TRANSFORM(m.perv_usl, '9999')
    oSheet1.Cells(10+m.num_lpu1,9).value2  = TRANSFORM(m.sum-m.sum_h, '999999.99')
    oSheet1.Cells(10+m.num_lpu1,10).value2 = TRANSFORM(m.sum_perv-m.sum_h, '999999.99')

    oSheet1.Cells(10+m.num_lpu1,11).value2 = TRANSFORM(m.paz_prin, '99999')
    oSheet1.Cells(10+m.num_lpu1,12).value2 = TRANSFORM(m.perv_paz_prin, '99999')
*    oSheet1.Cells(10+m.num_lpu1,11).value2 = TRANSFORM(m.usl_prin, '99999')
*    oSheet1.Cells(10+m.num_lpu1,12).value2 = TRANSFORM(m.perv_usl_prin, '99999')

    oSheet1.Cells(10+m.num_lpu1,13).value2 = TRANSFORM(m.usl_prin, '99999')
    oSheet1.Cells(10+m.num_lpu1,14).value2 = TRANSFORM(m.perv_usl_prin, '99999')
    oSheet1.Cells(10+m.num_lpu1,15).value2 = TRANSFORM(m.sum_prin-m.sum_h, '999999.99')
    oSheet1.Cells(10+m.num_lpu1,16).value2 = TRANSFORM(m.sum_perv_prin-m.sum_h, '999999.99')

    m.paz_tot1 = m.paz_tot1 + m.paz
    m.perv_paz_tot1 = m.perv_paz_tot1 + m.perv_paz
    m.usl_tot1 = m.usl_tot1 + m.usl
    m.perv_usl_tot1 = m.perv_usl_tot1 + m.perv_usl
    m.sum_tot1 = m.sum_tot1 + m.sum
    m.sum_perv_tot1 = m.sum_perv_tot1 + m.sum_perv

    m.paz_prin_tot1 = m.paz_prin_tot1 + m.paz_prin
    m.perv_paz_prin_tot1 = m.perv_paz_prin_tot1 + m.perv_paz_prin
    m.usl_prin_tot1 = m.usl_prin_tot1 + m.usl_prin
    m.perv_usl_prin_tot1 = m.perv_usl_prin_tot1 + m.perv_usl_prin
    m.sum_prin_tot1 = m.sum_prin_tot1 + m.sum_prin
    m.sum_perv_prin_tot1 = m.sum_perv_prin_tot1 + m.sum_perv_prin

    oSheet1.Cells(11+m.num_lpu1,1).Select
    ssel = oExcel.Selection 
    ssel.EntireRow.Insert
   ENDIF 

   USE IN talon 
   USE IN people
   USE IN errors
   
  ENDIF 
 ENDSCAN 
 WAIT CLEAR 

 SET RELATION OFF INTO sprlpu
 SET INDEX TO 
 fso.DeleteFile(pBin+'\aisoms.idx')
 USE 
 USE IN sprlpu
 USE IN spi_cz
 USE IN spi_czch
 
 oSheet1.Select
 oSheet1.Cells(12+m.num_lpu1,5).value2  = TRANSFORM(m.paz_tot1, '9999')
 oSheet1.Cells(12+m.num_lpu1,6).value2  = TRANSFORM(m.perv_paz_tot1, '9999')
 oSheet1.Cells(12+m.num_lpu1,7).value2  = TRANSFORM(m.usl_tot1, '9999')
 oSheet1.Cells(12+m.num_lpu1,8).value2  = TRANSFORM(m.perv_usl_tot1, '9999')
 oSheet1.Cells(12+m.num_lpu1,9).value2  = TRANSFORM(m.sum_tot1, '99999.99')
 oSheet1.Cells(12+m.num_lpu1,10).value2 = TRANSFORM(m.sum_perv_tot1, '99999.99')

 oSheet1.Cells(12+m.num_lpu1,11).value2 = TRANSFORM(m.paz_prin_tot1, '9999')
 oSheet1.Cells(12+m.num_lpu1,12).value2 = TRANSFORM(m.perv_paz_prin_tot1, '9999')
 oSheet1.Cells(12+m.num_lpu1,13).value2 = TRANSFORM(m.usl_prin_tot1, '9999')
 oSheet1.Cells(12+m.num_lpu1,14).value2 = TRANSFORM(m.perv_usl_prin_tot1, '9999')
 oSheet1.Cells(12+m.num_lpu1,15).value2 = TRANSFORM(m.sum_prin_tot1, '99999.99')
 oSheet1.Cells(12+m.num_lpu1,16).value2 = TRANSFORM(m.sum_perv_prin_tot1, '99999.99')

 oSheet2.Select
 oSheet2.Cells(12+m.num_lpu2,5).value2  = TRANSFORM(m.paz_tot2, '9999')
 oSheet2.Cells(12+m.num_lpu2,6).value2  = TRANSFORM(m.paz_tot217, '9999')
 oSheet2.Cells(12+m.num_lpu2,7).value2  = TRANSFORM(m.perv_paz_tot2, '9999')
 oSheet2.Cells(12+m.num_lpu2,8).value2  = TRANSFORM(m.perv_paz_tot217, '9999')
 oSheet2.Cells(12+m.num_lpu2,9).value2  = TRANSFORM(m.usl_tot2, '9999')
 oSheet2.Cells(12+m.num_lpu2,10).value2 = TRANSFORM(m.usl_tot217, '9999')
 oSheet2.Cells(12+m.num_lpu2,11).value2 = TRANSFORM(m.perv_usl_tot2, '9999')
 oSheet2.Cells(12+m.num_lpu2,12).value2 = TRANSFORM(m.perv_usl_tot217, '9999')
 oSheet2.Cells(12+m.num_lpu2,13).value2 = TRANSFORM(m.sum_tot2, '99999.99')
 oSheet2.Cells(12+m.num_lpu2,14).value2 = TRANSFORM(m.sum_tot217, '99999.99')
 oSheet2.Cells(12+m.num_lpu2,15).value2 = TRANSFORM(m.sum_perv_tot2, '99999.99')
 oSheet2.Cells(12+m.num_lpu2,16).value2 = TRANSFORM(m.sum_perv_tot217, '99999.99')

 oSheet2.Cells(12+m.num_lpu2,17).value2 = TRANSFORM(m.paz_tot2, '9999')
 oSheet2.Cells(12+m.num_lpu2,18).value2 = TRANSFORM(m.paz_tot217, '9999')
 oSheet2.Cells(12+m.num_lpu2,19).value2 = TRANSFORM(m.perv_usl_prin_tot2, '9999')
 oSheet2.Cells(12+m.num_lpu2,20).value2 = TRANSFORM(m.perv_usl_prin_tot217, '9999')
 oSheet2.Cells(12+m.num_lpu2,21).value2 = TRANSFORM(m.usl_prin_tot2, '9999')
 oSheet2.Cells(12+m.num_lpu2,22).value2 = TRANSFORM(m.usl_prin_tot217, '9999')
 oSheet2.Cells(12+m.num_lpu2,23).value2 = TRANSFORM(m.perv_usl_prin_tot2, '9999')
 oSheet2.Cells(12+m.num_lpu2,24).value2 = TRANSFORM(m.perv_usl_prin_tot217, '9999')
 oSheet2.Cells(12+m.num_lpu2,25).value2 = TRANSFORM(m.sum_prin_tot2, '99999.99')
 oSheet2.Cells(12+m.num_lpu2,26).value2 = TRANSFORM(m.sum_prin_tot217, '99999.99')
 oSheet2.Cells(12+m.num_lpu2,27).value2 = TRANSFORM(m.sum_perv_prin_tot2, '99999.99')
 oSheet2.Cells(12+m.num_lpu2,28).value2 = TRANSFORM(m.sum_perv_prin_tot217, '99999.99')

 DocName = pOut + '\Protokol_cz'
 IF fso.FileExists(docname+'.xls')
  fso.DeleteFile(docname+'.xls') 
 ENDIF 

 oBook.SaveAs(DocName,18)
 oBook.Close

 oExcel.Quit

RETURN 
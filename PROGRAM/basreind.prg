FUNCTION BasReind

WAIT "Переиндексация aisoms.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'excl') == 0
 SELECT AisOms
 DELETE TAG ALL 
 INDEX ON cmessage TAG cmessage
 INDEX ON mcod TAG mcod
 INDEX ON lpuid TAG lpuid
 INDEX ON TTOC(sent,1)      TAG sent 
 INDEX ON TTOC(recieved,1)  TAG recieved
 INDEX ON TTOC(processed,1) TAG processed
 INDEX ON paz TAG paz
 INDEX ON s_pred TAG s_pred
 INDEX ON sum_flk TAG sum_flk
 INDEX ON usr TAG usr
 INDEX ON e_mee TAG e_mee 
 INDEX ON e_ekmp TAG e_ekmp
 INDEX ON tpn TAG tpn
 USE 
ENDIF 
WAIT CLEAR 

WAIT "Переиндексация people.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+gcPeriod+'\people', 'people', 'excl') == 0
 SELECT people
 DELETE TAG ALL 
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON sn_pol TAG sn_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX ON dr TAG dr
 INDEX ON nlpu TAG nlpu
 INDEX on s_all TAG s_all
 USE 
ENDIF 
WAIT CLEAR 

WAIT "Переиндексация dsp.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+gcPeriod+'\dsp', 'dsp', 'excl') == 0
 SELECT dsp
 DELETE TAG ALL 
 INDEX ON mcod+PADL(recid,6,'0') TAG uniqq
 INDEX on mcod+sn_pol+PADL(cod,6,'0') TAG exptag
 USE 
ENDIF 
WAIT CLEAR 

pr4file = pBase+'\'+gcPeriod+'\pr4'
IF fso.FileExists(pr4file+'.dbf')
 WAIT "Переиндексация pr4.dbf..." WINDOW NOWAIT 
 IF OpenFile(pr4file, 'pr4', 'excl')<=0
  SELECT pr4
  DELETE TAG ALL 
  INDEX ON lpuid TAG lpuid
  INDEX on mcod TAG mcod
  USE 
 ENDIF 
 WAIT CLEAR 
ENDIF 

WAIT "Переиндексация talon.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+gcPeriod+'\talon', 'talon', 'excl') == 0
 SELECT talon
 DELETE TAG ALL 
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON brid TAG brid
 INDEX ON c_i TAG c_i
 INDEX ON sn_pol TAG sn_pol
 INDEX ON otd TAG otd
 INDEX ON ds TAG ds
 INDEX ON d_u TAG d_u
 INDEX ON cod TAG cod
 INDEX ON profil TAG profil
* INDEX ON sn_pol+STR(cod,6)+DTOS(d_u) TAG ExpTag
* INDEX ON sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u) TAG unik 
 USE 
ENDIF 
WAIT CLEAR 

WAIT "Переиндексация otdel.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+gcPeriod+'\otdel', 'otdel', 'excl') == 0
 SELECT otdel
 DELETE TAG ALL 
 INDEX ON iotd TAG iotd
 INDEX ON mcod+' '+iotd TAG unkey
 USE 
ENDIF 
WAIT CLEAR 

WAIT "Переиндексация doctor.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+gcPeriod+'\doctor', 'doctor', 'excl') == 0
 SELECT doctor
 DELETE TAG ALL 
 USE 
ENDIF 
WAIT CLEAR 

errsv  = pBase+'\'+gcPeriod+'\e'+m.gcperiod
WAIT "Переиндексация e"+m.gcperiod+".dbf..." WINDOW NOWAIT 
IF OpenFile(errsv, 'errsv', 'excl') == 0
 SELECT errsv
 DELETE TAG ALL 
 INDEX FOR UPPER(f)='R' ON rid TAG rrid
 INDEX FOR UPPER(f)='S' ON rid TAG rid
 USE 
ENDIF 
WAIT CLEAR 

*WAIT "Переиндексация stat.dbf..." WINDOW NOWAIT 
*IF OpenFile(pBase+'\'+gcPeriod+'\stat', 'stat', 'excl') == 0
* SELECT stat
* DELETE TAG ALL 
* INDEX ON mcod TAG mcod 
* USE 
*ENDIF 
*WAIT CLEAR 

WAIT "Переиндексация mee.dbf..." WINDOW NOWAIT 
IF OpenFile(pBase+'\'+gcPeriod+'\mee', 'mee', 'excl') == 0
 SELECT mee
 DELETE TAG ALL 
 INDEX ON rid TAG rid
 USE 
ENDIF 
WAIT CLEAR 

WAIT "Переиндексация aisoms.dbf..." WINDOW NOWAIT 
IF OpenFile(pDouble+'\AisOms', 'AisOms', 'excl') == 0
 SELECT AisOms
 DELETE TAG ALL 
 INDEX ON cmessage TAG cmessage
 INDEX ON mcod TAG mcod
 INDEX ON lpuid TAG lpuid
 INDEX ON TTOC(sent,1)      TAG sent 
 INDEX ON TTOC(recieved,1)  TAG recieved
 INDEX ON TTOC(processed,1) TAG processed
 INDEX ON paz TAG paz
 INDEX ON s_pred TAG s_pred
 USE
ENDIF 
WAIT CLEAR 

WAIT "Переиндексация aisoms.dbf..." WINDOW NOWAIT 
IF OpenFile(pTrash+'\AisOms', 'AisOms', 'excl') == 0
 SELECT AisOms
 DELETE TAG ALL 
 INDEX ON cmessage TAG cmessage
 INDEX ON mcod TAG mcod
 INDEX ON lpuid TAG lpuid
 INDEX ON TTOC(sent,1)      TAG sent 
 INDEX ON TTOC(recieved,1)  TAG recieved
 INDEX ON TTOC(processed,1) TAG processed
 USE
ENDIF 
WAIT CLEAR 

WAIT "Переиндексация expdetails.dbf..." WINDOW NOWAIT 
IF OpenFile(pbase+'\'+gcperiod+'\expdetails', 'expdetails', 'excl') == 0
 SELECT expdetails
 DELETE TAG ALL 
 INDEX ON period+mcod+et TAG ikey
 INDEX ON mcod TAG mcod 
 INDEX ON et TAG et
 USE
ENDIF 
WAIT CLEAR 

tn_result = 0
tn_result = tn_result + OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'shar', 'mcod')
IF tn_result > 0
 RETURN 
ENDIF 

SELECT AisOms
SCAN  
 WAIT mcod WINDOW NOWAIT 
 lcPath = pbase+'\'+m.gcperiod+'\'+mcod
 IF fso.FileExists(lcPath+'\People.dbf') AND ;
    fso.FileExists(lcPath+'\Talon.dbf') AND ;
    fso.FileExists(lcPath+'\Otdel.dbf')
  =OneReind(lcPath)
 ENDIF 
 WAIT CLEAR 
ENDSCAN 
USE 

RETURN 


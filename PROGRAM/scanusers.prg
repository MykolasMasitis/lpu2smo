PROCEDURE ScanUsers
 FOR nUser=1 TO 10
  lcUser = 'USR'+PADL(nUser,3,'0')
  IF fso.FolderExists(pAisOms+'\'+lcUser)
   WAIT '咽劳刃温劳扰 娜信室涡热 ' + lcUser WINDOW NOWAIT 
   =CheckMail(lcUser)
   WAIT CLEAR 
  ENDIF 
 ENDFOR  
RETURN 
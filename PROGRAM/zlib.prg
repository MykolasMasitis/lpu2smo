Procedure CompressString  
  Lparameters tcString  
  	If Len(tcString)=0  
  		Return ""  
  	EndIf  
  	DECLARE INTEGER compress2 IN zlib1 STRING @cDest, LONG@ nDest,String cSrc, LONG nSrc,integer nLevel  
  	DECLARE INTEGER compressBound IN zlib1 integer nSrc  
  	  
  	Local lnDest,lcDest  
  	lnDest = compressBound(Len(tcString))  
  	lcDest = Space(lnDest)  
  	If compress2(@lcDest,@lnDest,tcString,Len(tcString),9) = 0  
  		Return BinToC(Len(tcString),"4RS") + Left(lcDest,lnDest)  
  	Else  
  		Error 'Error compressing string'  
  		Return .NULL.  
  	EndIf  
  EndProc  
    
  Procedure DecompressString  
  Lparameters tcString  
  	If Len(tcString)=0  
  		Return ""  
  	EndIf  
  	  
  	Local lnDest,lcDest  
  	Try   
  		lnDest = CToBin(Left(tcString,4),"4RS")	  
  	Catch  
  		lnDest = -1  
  	EndTry  
  	  
  	If lnDest<=0  
  		Error "String is corrupted or has invalid format"  
  		Return .NULL.  
  	EndIf  
    
  	lcDest = Space(lnDest)  
  	DECLARE INTEGER uncompress IN zlib1 STRING @cDest, LONG@ nDest,String cSrc, LONG nSrc  
  	If uncompress(@lcDest,@lnDest,Substr(tcString,5),Len(tcString)-4) = 0  
  		Return Left(lcDest,lnDest)  
  	Else  
  		Error 'Error decompressing string'  
  		Return .NULL.  
  	EndIf  
  EndProc
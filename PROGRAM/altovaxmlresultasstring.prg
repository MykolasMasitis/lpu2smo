FUNCTION AltovaXMLResultAsString(para1, para2)
 m.nparams = PARAMETERS()
 
 IF m.nparams!=2
  MESSAGEBOX('вхякн оепедюммшу оюпюлерпнб ме пюбмн 2!',0+16,'')
  RETURN .f.
 ENDIF 

 fso = CREATEOBJECT('Scripting.FileSystemObject')

 m.xslpath = ALLTRIM(para1)
 m.xmlpath = ALLTRIM(para2)
 
 IF !fso.FileExists(m.xslpath)
  MESSAGEBOX('ме мюидем тюик'+CHR(13)+CHR(10)+m.xslpath,0+16,'')
  RELEASE fso 
  RETURN .f.
 ENDIF 
 IF !fso.FileExists(m.xmlpath)
  MESSAGEBOX('ме мюидем тюик'+CHR(13)+CHR(10)+m.xmlpath,0+16,'')
  RELEASE fso 
  RETURN .f.
 ENDIF 
 
 m.xslt = FILETOSTR(m.xslpath)
 m.xml  = FILETOSTR(m.xmlpath)
 
 m.rspath = fso.GetParentFolderName(m.xmlpath)+'\result.001'
 
 oxml  = CREATEOBJECT("AltovaXML.Application")
 oxslt = oxml.XSLT2
 oxslt.InputXMLFromText = m.xml
 oxslt.XSLFromText      = m.xslt
 m.rsstring = oxslt.ExecuteAndGetResultAsString
 STRTOFILE(m.rsstring, m.rspath)
 
 RELEASE fso 
RETURN .t.
FUNCTION SendModByMail(lcmcod)
 
 m.mcod = lcmcod
 lcPeriod = STR(tYear,4) + PADL(tMonth,2,'0')
 m.DocName   = pout+'\'+lcPeriod+'\Pm' + m.mcod
 
 IF !fso.FolderExists(pOut)
  MESSAGEBOX('Директория '+pOut+' отсутствует!',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pOut+'\'+lcPeriod)
  MESSAGEBOX('Директория '+pOut+'\'+lcPeriod+' отсутствует!',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.DocName+'.doc')
  MESSAGEBOX('По выбранному ЛПУ'+CHR(13)+CHR(10)+;
   'уведомление о модернизации не сформировано!'+CHR(13)+CHR(10),0+16,m.Docname)
  RETURN 
 ENDIF 
 
 IF SEEK(m.mcod, 'emails')
  m.email = ALLTRIM(emails.e_mail)
 ELSE 
  m.email = ''
 ENDIF 
 
 IF EMPTY(m.email)
  MESSAGEBOX('Для выбранного ЛПУ'+CHR(13)+CHR(10)+;
   'адрес электронной почты не определен!'+CHR(13)+CHR(10),0+16,m.Docname)
  RETURN 
 ENDIF 

 #Define MAPI_ORIG 0
 #Define MAPI_TO 1
 #Define MAPI_CC 2
 #Define MAPI_BCC 3

 #Define IMPORTANCE_LOW 0
 #Define IMPORTANCE_NORMAL 1
 #Define IMPORTANCE_HIGH 2

 #DEFINE CRCR CHR(13)

 LOCAL lcMessage, llSuccess
	
 m.lcSubject = "Уведомление о модернизации"
 m.lcMessage = "Добрый день!"+CHR(13)+CHR(10)
 m.lcMessage = m.lcMessage + "Во вложении содержится уведомление о модернизации"+CHR(13)+CHR(10)
 m.lcMessage = m.lcMessage + "за "+NameOfMonth(tMonth)+ ' '+STR(tYear,4)+' года'+CHR(13)+CHR(10)
 m.lcMessage = m.lcMessage + "от Страховой медицинской организации "+m.qname

 m.lcAddress    = m.email
 m.lcAttachment = m.docname
 
 IF EMCreateMessage(m.lcSubject, m.lcMessage, IMPORTANCE_HIGH)
  IF EMAddRecipient(m.lcAddress, MAPI_TO)
   IF EMAddAttachment(m.lcAttachment)
    IF EMSend(.T.)
     m.llSuccess = .T.
    ENDIF
   ENDIF 
  ENDIF
 ENDIF
	
 IF m.llSuccess
  MESSAGEBOX("Ваше сообщение было успешно отправлено!", 0+48, "Успешная отправка сообщения") 
 ELSE
  MESSAGEBOX("Не удалось отправить сообщение!", 64, "Ошибка при отправке")
 ENDIF 

RETURN 
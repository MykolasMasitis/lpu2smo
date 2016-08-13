FUNCTION SendZapr2LpuByMail(m.mcod, m.tip)

 m.pDir    = IIF(m.Tip=0, m.pMee, m.pMee)+'\'+m.gcperiod+IIF(m.tip=0,'\','\0000000\')+m.mcod
 m.DocName = m.pDir+'\Rq'+IIF(m.Tip=0,'',flcod)+'.xls'
 
 IF !fso.FileExists(m.DocName)
  MESSAGEBOX('По выбранному ЛПУ'+CHR(13)+CHR(10)+;
   'запрос на подбор карт не сформирован!'+CHR(13)+CHR(10),0+16,m.Docname)
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
	
 m.lcSubject = "Запрос карт на проведение МЭЭ"
 m.lcMessage = "Добрый день!"+CHR(13)+CHR(10)
 m.lcMessage = m.lcMessage + "Во вложении содержится запрос на подбор карт"+CHR(13)+CHR(10)
 m.lcMessage = m.lcMessage + "для проведения медико-экономической экспертизы"+CHR(13)+CHR(10)
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
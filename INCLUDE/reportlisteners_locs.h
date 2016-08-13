* locs for VCX super-classes

#DEFINE OUTPUTCLASS_APPNAME_LOC              "КИС ВиртуалСофт"

#DEFINE OUTPUTCLASS_CONFIGTABLECREATED_LOC     "Конфигаруционная таблица "+ lcDBF + " создана."
                                                                                                "записана на диск."
#DEFINE OUTPUTCLASS_CONFIGTABLEWRONG_LOC     "Конфигаруционная таблица имеет не корректный формат."

* this is different from XML because in the XML class they
* are used in SEEK() and require specific tagnames, whereas
* the superclasses just require certain indexes for optimizing LOCATES,

#DEFINE OUTPUTCLASS_CONFIGINDEXMISSING_LOC   "Конфигаруционная таблица отсутствует "+ CHR(13) + ;
                                                                                                   "один или более требуемых индексов."

#DEFINE OUTPUTCLASS_INITSTATUS_LOC           "Инициализация... "
#DEFINE OUTPUTCLASS_PREPSTATUS_LOC           "Процесс расчета... "
#DEFINE OUTPUTCLASS_RUNSTATUS_LOC            "Подготовка отчета... "
#DEFINE OUTPUTCLASS_TIME_SECONDS_LOC         SPACE(1) + "Секунд (с)"
#DEFINE OUTPUTCLASS_CANCEL_INSTRUCTIONS_LOC  "Нажмите Esc для отмены... "
#DEFINE OUTPUTCLASS_REPORT_CANCELQUERY_LOC   "Остановить выполнения отчета?"+CHR(13) + ;
                                             "(Если нажмете 'Нет(No)', выполнение отчета продолжиться.)"
#DEFINE OUTPUTCLASS_REPORT_INCOMPLETE_LOC    "Выполнение отчета было прервано." + CHR(13) + ;
                                             "Вы получили не полный результат."

#DEFINE OUTPUTCLASS_SUCCESS_LOC              THIS.AppName+" создать ваш отчет как" + ;
                                             CHR(13) + THIS.TargetFileName+"." + CHR(13) + ;
                                             IIF(THIS.AllowModalMessages,;
                                              "Нажмите 'Да(Yes)' для сохранения" + CHR(13) + ;
                                             "этого имени файла в буфере обмена.","")

#DEFINE OUTPUTCLASS_NOFILECREATE_LOC         "Файл " + THIS.TargetFileName+" не может быть создан."

#DEFINE OUTPUTCLASS_CREATEERRORS_LOC         THIS.AppName+" создан ваш репорт как"+ ;
                                             CHR(13)+THIS.TargetFileName+". "+CHR(13)+ ;
                                             "Однака произошла ошибка в процессе." + CHR(13) + ;
                                             OUTPUTCLASS_REPORT_INCOMPLETE_LOC

#DEFINE OUTPUTCLASS_NOCREATE_LOC             THIS.AppName +" не доступен для создания отчета."
 
#DEFINE OUTPUTCLASS_ERRNOLABEL_LOC           "Ошибка:           "
#DEFINE OUTPUTCLASS_ERRPROCLABEL_LOC         "Метод:       "
#DEFINE OUTPUTCLASS_ERRLINELABEL_LOC         "Линия:            " 

* the following loc is eval'd for updateListener's actual progress bar message.  
* In most cases,
* changing this value is overkill, as the localizable portions of
* the message are already localized as separate properties.
* All the status messagse as well as the therm caption can
* also be set at runtime without touching the locs.
#DEFINE OUTPUTCLASS_THERMCAPTION_LOC        [cMessage+ " "+ TRANSFORM(INT(THIS.PercentDone*100)) + "%" ] + ;
                                            [+ IIF(NOT THIS.IncludeSeconds, "" , " "+] + ;
                                            [TRANSFORM(IIF(THIS.IsRunning,DATETIME(), THIS.ReportStopRunDateTime)-] + ;
                                            [THIS.ReportStartRunDateTime)+" " + THIS.SecondsText)]


* locs for XML Listener class:

#DEFINE OUTPUTXML_APPNAME_LOC               "XML Listener"

#DEFINE OUTPUTXML_CONFIGTAGMISSING_LOC      "At least one required index tag is missing "+ CHR(13) + ;
                                            "from the configuration table."

#DEFINE OUTPUTXML_FRXMISSING_LOC            "Required FRX cursor is not available." 
   
   
#DEFINE OUTPUTXML_FRXCURSOR_MISSING_LOC     "FRX cursor helper object cannot be found in " + CHR(13)+ ;
                                            "_FRXCURSOR.VCX, "+CHR(13)+ ; 
                                            "_REPORTOUTPUT." + CHR(13) + CHR(13)+ ;
                                            "Some features of this class may not be active."   

* locs for XML Display Listener Class:

#DEFINE OUTPUTXMLDISPLAY_APPNAME_LOC        "XML Display Listener"

* locs for HTML:

#DEFINE OUTPUTHTML_APPNAME_LOC              "HTML Listener"
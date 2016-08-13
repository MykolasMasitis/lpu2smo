PROCEDURE m_menu
SET SYSMENU TO

DEFINE PAD mmenu_1 OF _MSYSMENU PROMPT '\<хмтнплюжхъ нр кос' COLOR SCHEME 3 KEY ALT+A, ""
DEFINE PAD mmenu_2 OF _MSYSMENU PROMPT '\<щйяоепрхгю' COLOR SCHEME 3 KEY ALT+C , ""
DEFINE PAD mmenu_3 OF _MSYSMENU PROMPT '\<яопюбнвмхйх' COLOR SCHEME 3 KEY ALT+C , ""
DEFINE PAD mmenu_4 OF _MSYSMENU PROMPT '\<хмтнплюжхъ б лцтнля' COLOR SCHEME 3 KEY ALT+D , ""
DEFINE PAD mmenu_5 OF _MSYSMENU PROMPT '\<тхмюмяш' COLOR SCHEME 3 KEY ALT+E , ""
DEFINE PAD mmenu_6 OF _MSYSMENU PROMPT '\<яепбхя' COLOR SCHEME 3 KEY ALT+F , ""
DEFINE PAD mmenu_7 OF _MSYSMENU PROMPT '\<яепбхя-2' COLOR SCHEME 3 KEY ALT+F , ""
ON PAD mmenu_1 OF _MSYSMENU ACTIVATE POPUP popInfFrLpu
ON PAD mmenu_2 OF _MSYSMENU ACTIVATE POPUP popMEE
ON PAD mmenu_3 OF _MSYSMENU ACTIVATE POPUP popMedSpr
ON PAD mmenu_4 OF _MSYSMENU ACTIVATE POPUP popInfToMGF
ON PAD mmenu_5 OF _MSYSMENU ACTIVATE POPUP popBuch
ON PAD mmenu_6 OF _MSYSMENU ACTIVATE POPUP popTuneUp
ON PAD mmenu_7 OF _MSYSMENU ACTIVATE POPUP popTuneUp2

DEFINE POPUP popInfFrLpu MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popInfFrLpu PROMPT 'опхел хмнплюжхх нр кос (юхя нля)'
DEFINE BAR 02 OF popInfFrLpu PROMPT 'тнплхпнбюмхе UD-тюикнб'
DEFINE BAR 03 OF popInfFrLpu PROMPT 'тнплхпнбюмхе UD-тюикнб (VER. 2)' SKIP 
DEFINE BAR 04 OF popInfFrLpu PROMPT 'тнплхпнбюмхе UP-тюикнб'
DEFINE BAR 05 OF popInfFrLpu PROMPT 'яопюбнвмхй днцнбнпнб я кос'
DEFINE BAR 06 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 07 OF popInfFrLpu PROMPT 'детейрмше хмтнплюжхнммше оняшкйх' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 08 OF popInfFrLpu PROMPT 'онбрнпмше хмтнплюжхнммше оняшкйх' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 09 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 10 OF popInfFrLpu PROMPT 'ябндмши явер гю оепхнд'
DEFINE BAR 11 OF popInfFrLpu PROMPT 'ябндмши явер гю цнд' SKIP FOR m.IsNotePad
DEFINE BAR 12 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 13 OF popInfFrLpu PROMPT 'нропюбйю пегскэрюрнб ткй б кос' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 14 OF popInfFrLpu PROMPT 'нропюбйю пегскэрюрнб лщй б кос' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 15 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 16 OF popInfFrLpu PROMPT 'дхяоюмяепхгюжхъ'
DEFINE BAR 17 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 18 OF popInfFrLpu PROMPT 'нропюбйю тхм-тюикю б лцтнля' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 19 OF popInfFrLpu PROMPT 'нропюбйю оепянрверю б лцтнля' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 20 OF popInfFrLpu PROMPT 'нропюбйю VZV-тюикю б лцтнля' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 21 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 22 OF popInfFrLpu PROMPT 'нропюбйю ME-тюикнб б лцтнля' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 23 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 24 OF popInfFrLpu PROMPT 'ямърэ PPA ян ябнху'
DEFINE BAR 25 OF popInfFrLpu PROMPT 'бняярюмнбхрэ тюикш ньханй он оепянрверс'
DEFINE BAR 26 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 27 OF popInfFrLpu PROMPT 'явер мю оюжхемрю'
DEFINE BAR 28 OF popInfFrLpu PROMPT 'яопюбнвмхй онкэгнбюрекеи'
DEFINE BAR 29 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 30 OF popInfFrLpu PROMPT 'оюйермюъ оевюрэ'
DEFINE BAR 31 OF popInfFrLpu PROMPT 'бшунд'

ON SELECTION BAR 01 OF popInfFrLpu DO FORM MailView
ON SELECTION BAR 02 OF popInfFrLpu do MakeUDFiles
ON SELECTION BAR 03 OF popInfFrLpu do MakeUDFL2
ON SELECTION BAR 04 OF popInfFrLpu do MakeUPFiles
ON SELECTION BAR 05 OF popInfFrLpu DO FORM ViewDogs
ON SELECTION BAR 07 OF popInfFrLpu DO FORM MailTrash
ON SELECTION BAR 08 OF popInfFrLpu DO FORM MailDouble
ON SELECTION BAR 10 OF popInfFrLpu DO FORM ViewPeriod
ON SELECTION BAR 11 OF popInfFrLpu DO FORM ViewSvYear
ON SELECTION BAR 13 OF popInfFrLpu DO Flk2Lpu
ON SELECTION BAR 14 OF popInfFrLpu DO Mek2Lpu
ON BAR 16 OF popInfFrLpu ACTIVATE POPUP DispMenu
ON SELECTION BAR 18 OF popInfFrLpu DO SendFinFile
ON SELECTION BAR 19 OF popInfFrLpu DO SendPers
ON SELECTION BAR 20 OF popInfFrLpu DO SendVzvFile
ON SELECTION BAR 22 OF popInfFrLpu DO SendMEFiles
ON SELECTION BAR 24 OF popInfFrLpu DO DelPPA
ON SELECTION BAR 25 OF popInfFrLpu DO restefls
ON SELECTION BAR 27 OF popInfFrLpu DO FindPaz
ON SELECTION BAR 28 OF popInfFrLpu DO form sprusers
ON BAR 30 OF popInfFrLpu ACTIVATE POPUP popPrn
ON SELECTION BAR 31 OF popInfFrLpu clea events 

DEFINE POPUP popPrn MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF popPrn PROMPT 'оевюрэ опнрнйнкнб опх╗лйх яверю'
DEFINE BAR 2 OF popPrn PROMPT 'оевюрэ юйрнб лщй'
DEFINE BAR 3 OF popPrn PROMPT 'оевюрэ юйрнб на нокюре он ондсьебнцн тхмюмяхпнбюмхъ'
ON SELECTION BAR 1 OF popPrn DO PackPrnPr
ON SELECTION BAR 2 OF popPrn DO PackPrnMc
ON SELECTION BAR 3 OF popPrn DO PackPrnPdf

DEFINE POPUP DispMenu MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF DispMenu PROMPT 'тнплхпнбюмхе тюикю DSP'
DEFINE BAR 2 OF DispMenu PROMPT '\-'
DEFINE BAR 3 OF DispMenu PROMPT 'тнплхпнбюмхе нрверю'
DEFINE BAR 4 OF DispMenu PROMPT '\-'
DEFINE BAR 5 OF DispMenu PROMPT 'тнплхпнбюмхе нрвернб б кос'
DEFINE BAR 6 OF DispMenu PROMPT '\-'
DEFINE BAR 7 OF DispMenu PROMPT 'йнппейрхпнбйю йнднб кос'
DEFINE BAR 8 OF DispMenu PROMPT '\-'
DEFINE BAR 9 OF DispMenu PROMPT 'йнппейрхпнбйю йнднб кос-2'
ON SELECTION BAR 1 OF DispMenu DO MakeDspFile
ON SELECTION BAR 3 OF DispMenu DO NewDspMonitorN
ON SELECTION BAR 5 OF DispMenu DO FormDDDS
ON SELECTION BAR 7 OF DispMenu DO CorrDsp
ON SELECTION BAR 9 OF DispMenu DO CorrMcod

DEFINE POPUP AccsPeriod MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF AccsPeriod PROMPT 'ябндмши явер'
ON SELECTION BAR 1 OF AccsPeriod DO FORM ViewPeriod

DEFINE POPUP AccsYear MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF AccsYear PROMPT 'пецхярп'
DEFINE BAR 2 OF AccsYear PROMPT 'сяксцх'
DEFINE BAR 3 OF AccsYear PROMPT 'йнлахмюжхъ'
ON SELECTION BAR 3 OF AccsYear DO FORM ViewSvYear

DEFINE POPUP popInfToMGF MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 01 OF popInfToMGF PROMPT 'юцпецхпнбюммше яверю' PICTURE 'GROUP3.BMP' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 02 OF popInfToMGF PROMPT 'оепянмхтхжхпнбюммше яверю' PICTURE 'GROUP3.BMP' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 03 OF popInfToMGF PROMPT '\-'
DEFINE BAR 04 OF popInfToMGF PROMPT 'тнплю ╧1'
DEFINE BAR 05 OF popInfToMGF PROMPT 'тнплю "оц" (лщй)'
DEFINE BAR 06 OF popInfToMGF PROMPT '\-' 
DEFINE BAR 07 OF popInfToMGF PROMPT 'тнплю "оц" (лщщ ОКЮМНБЮЪ)'
DEFINE BAR 08 OF popInfToMGF PROMPT 'тнплю "оц" (лщщ ЖЕКЕБЮЪ)'
DEFINE BAR 09 OF popInfToMGF PROMPT 'тнплю "оц" (лщщ РЕЛЮРХВЕЯЙЮЪ)'
DEFINE BAR 10 OF popInfToMGF PROMPT 'тнплю "оц" (лщщ ОН ФЮКНАЮЛ)'
DEFINE BAR 11 OF popInfToMGF PROMPT '\-' 
DEFINE BAR 12 OF popInfToMGF PROMPT 'тнплю "оц" (щйло ОКЮМНБЮЪ)'
DEFINE BAR 13 OF popInfToMGF PROMPT 'тнплю "оц" (щйло ЖЕКЕБЮЪ)'
DEFINE BAR 14 OF popInfToMGF PROMPT 'тнплю "оц" (щйло РЕЛЮРХВЕЯЙЮЪ)'
DEFINE BAR 15 OF popInfToMGF PROMPT 'тнплю "оц" (щйло ОН ФЮКНАЮЛ)'
DEFINE BAR 16 OF popInfToMGF PROMPT '\-' 
DEFINE BAR 17 OF popInfToMGF PROMPT 'напюыюелнярэ гюярпюунбюммшу' 
DEFINE BAR 18 OF popInfToMGF PROMPT '\-' 
DEFINE BAR 19 OF popInfToMGF PROMPT 'ябепхрэ я мнлепмхйнл' 
DEFINE BAR 20 OF popInfToMGF PROMPT '\-' 
DEFINE BAR 21 OF popInfToMGF PROMPT 'нрвер он онкс/бнгпюярс (демэцх)' 
DEFINE BAR 22 OF popInfToMGF PROMPT 'нрвер он онкс/бнгпюярс (кчдх)' 
DEFINE BAR 23 OF popInfToMGF PROMPT 'нрвер он онкс/бнгпюярс (демэцх) ярюж' 
DEFINE BAR 24 OF popInfToMGF PROMPT 'нрвер он онкс/бнгпюярс (кчдх) ярюж' 

ON SELECTION BAR 01 OF popInfToMGF goApp.doForm('frm_agreg','mgfoms')
ON SELECTION BAR 02 OF popInfToMGF DO MakeYFiles
ON SELECTION BAR 04 OF popInfToMGF DO FormN1
IF m.qcod!='S6'
 ON SELECTION BAR 05 OF popInfToMGF DO FormPGMek
ELSE 
 ON SELECTION BAR 05 OF popInfToMGF DO FormPGMekS6
ENDIF 
ON SELECTION BAR 07 OF popInfToMGF DO FormPGMee WITH 2
ON SELECTION BAR 08 OF popInfToMGF DO FormPGMee WITH 3
ON SELECTION BAR 09 OF popInfToMGF DO FormPGMee WITH 7
ON SELECTION BAR 10 OF popInfToMGF DO FormPGMee WITH 8
ON SELECTION BAR 12 OF popInfToMGF DO FormPGMee WITH 4
ON SELECTION BAR 13 OF popInfToMGF DO FormPGMee WITH 5
ON SELECTION BAR 14 OF popInfToMGF DO FormPGMee WITH 6
ON SELECTION BAR 15 OF popInfToMGF DO FormPGMee WITH 9
ON SELECTION BAR 14 OF popInfToMGF DO ObrPrikl
ON SELECTION BAR 16 OF popInfToMGF DO SvOutS
ON SELECTION BAR 21 OF popInfToMGF DO SagOpl
ON SELECTION BAR 22 OF popInfToMGF DO SagOpl2
ON SELECTION BAR 23 OF popInfToMGF DO SagOpls
ON SELECTION BAR 24 OF popInfToMGF DO SagOpl2s

DEFINE POPUP popMEE MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popMEE PROMPT 'щйяоепрхгю рейсыецн оепхндю'
DEFINE BAR 02 OF popMEE PROMPT 'щйяоепрхгю опнхгбнкэмнцн оепхндю'
DEFINE BAR 03 OF popMEE PROMPT 'йпхрепхх нранпю дкъ щйяоепрхгш'
DEFINE BAR 04 OF popMEE PROMPT 'йпхрепхх нранпю дкъ щйяоепрхгш (NEW)'
DEFINE BAR 05 OF popMEE PROMPT '\-'
DEFINE BAR 06 OF popMEE PROMPT 'тнплхпнбюмхе ME-тюикнб'
DEFINE BAR 07 OF popMEE PROMPT 'фспмюк ябндмшу юйрнб'
DEFINE BAR 08 OF popMEE PROMPT 'фспмюк юйрнб ярпюунбшу яксвюеб'
DEFINE BAR 09 OF popMEE PROMPT '\-'
DEFINE BAR 10 OF popMEE PROMPT 'ятнплхпнбюрэ тюикш хлонпрю' PICTURE 'GROUP3.BMP'
DEFINE BAR 11 OF popMEE PROMPT 'хлонпрхпнбюрэ юйрш щйяоепрхгш' PICTURE 'GROUP3.BMP'
DEFINE BAR 12 OF popMEE PROMPT '\-'
DEFINE BAR 13 OF popMEE PROMPT 'гюцпсгхрэ тюикш хлонпрю' PICTURE 'GROUP2.BMP'
DEFINE BAR 14 OF popMEE PROMPT 'щйяонпрхпнбюрэ юйрш щйяоепрхгш' PICTURE 'GROUP2.BMP'
DEFINE BAR 15 OF popMEE PROMPT '\-'
DEFINE BAR 16 OF popMEE PROMPT 'опнбепйю яннрберярбхъ яслл ямърхи'
DEFINE BAR 17 OF popMEE PROMPT '\-'
DEFINE BAR 18 OF popMEE PROMPT 'ледхйн-щйнмнлхвеяйхе тнплш S7'
DEFINE BAR 19 OF popMEE PROMPT 'ледхйн-щйнмнлхвеяйхе тнплш S2'
DEFINE BAR 20 OF popMEE PROMPT 'гюопня ттнля нр 18.02.2016'
DEFINE BAR 21 OF popMEE PROMPT 'гюопня ттнля нр 18.02.2016 (лщй)'

ON SELECTION BAR 01 OF PopMEE DO viewmee
ON SELECTION BAR 02 OF PopMEE DO FORM ViewYear
ON BAR 03 OF PopMEE ACTIVATE POPUP ExpCriteria
ON BAR 04 OF PopMEE ACTIVATE POPUP ExpCritNew
ON SELECTION BAR 06 OF popMEE DO MakeMEFiles
ON SELECTION BAR 07 OF popMEE DO FORM ViewActSV
ON SELECTION BAR 08 OF popMEE DO FORM ViewActSS
ON SELECTION BAR 10 OF popMEE DO ImpExp
ON SELECTION BAR 11 OF popMEE DO ImpActs
ON SELECTION BAR 13 OF popMEE DO ExpExp
ON SELECTION BAR 14 OF popMEE DO ExpActs
ON SELECTION BAR 16 OF popMEE DO CmpMee
ON BAR 18 OF PopMEE ACTIVATE POPUP popS7
ON BAR 19 OF PopMEE ACTIVATE POPUP popS2
ON SELECTION BAR 20 OF popMEE DO FFOMS18022016
ON SELECTION BAR 21 OF popMEE DO FFOMS18022016mek

DEFINE POPUP popS7 MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popS7 PROMPT 'тнплю ь-0 (нрвер дкъ йпс)'
DEFINE BAR 02 OF popS7 PROMPT 'тнплю ь-1'
DEFINE BAR 03 OF popS7 PROMPT 'тнплю ь-2 (нрвер он пндднлюл)'
DEFINE BAR 04 OF popS7 PROMPT 'тнплю ь-3 (нрвер он опнбедеммшл щйяоепрхгюл)'
DEFINE BAR 05 OF popS7 PROMPT 'тнплю ь-3ахя (нрвер он опнбедеммшл щйяоепрхгюл)'
DEFINE BAR 06 OF popS7 PROMPT 'тнплю ь-4 лщщ (опхкнфемхе 8)'
DEFINE BAR 07 OF popS7 PROMPT 'тнплю ь-4 щйло (опхкнфемхе 8)'
DEFINE BAR 08 OF popS7 PROMPT 'тнплю ь-5 (ярюрхярхйю ямърхи он лщй)'
DEFINE BAR 09 OF popS7 PROMPT 'ярюрхярхйю ямърхи б пюгпеге лщй,лщщ,щйло'
DEFINE BAR 10 OF popS7 PROMPT '\-'
DEFINE BAR 11 OF popS7 PROMPT 'тнплю ь-6 (пюгдек II тнплш 1)'
DEFINE BAR 12 OF popS7 PROMPT 'тнплю ь-7 (нрвер он жекебни щйяоепрхге)'
DEFINE BAR 13 OF popS7 PROMPT 'тнплю ь-8'
ON SELECTION BAR 01 OF popS7 DO FormSh0
ON SELECTION BAR 02 OF popS7 DO FormSh1
ON SELECTION BAR 03 OF popS7 DO FormSh2
ON SELECTION BAR 04 OF popS7 DO FormSh3
ON SELECTION BAR 05 OF popS7 DO FormSh3Bis
ON SELECTION BAR 06 OF popS7 FormSh4(1) && лщщ
ON SELECTION BAR 07 OF popS7 FormSh4(2) && щйло
ON SELECTION BAR 08 OF popS7 DO FormSh5
ON SELECTION BAR 09 OF popS7 DO FormSh5Bis
ON SELECTION BAR 11 OF popS7 DO FormSh6
ON SELECTION BAR 12 OF popS7 do RepExp7
ON SELECTION BAR 13 OF popS7 do FormSh8

DEFINE POPUP popS2 MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popS2 PROMPT 'ябндмши гю леяъж'
DEFINE BAR 02 OF popS2 PROMPT 'керюкэмше хяундш (онкмши)'
DEFINE BAR 03 OF popS2 PROMPT 'керюкэмше хяундш (йпюрйхи)'
ON SELECTION BAR 01 OF popS2 DO FormS20
ON SELECTION BAR 02 OF popS2 DO FormS21
ON SELECTION BAR 03 OF popS2 DO FormS22

DEFINE POPUP ExpCriteria MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF ExpCriteria PROMPT 'керюкэмше хяундш'
DEFINE BAR 2 OF ExpCriteria PROMPT 'онбрнпмше цняохрюкхгюжхх'
DEFINE BAR 3 OF ExpCriteria PROMPT 'оепеяевемхъ бхднб ледонлных'
DEFINE BAR 4 OF ExpCriteria PROMPT 'нранп он опнтхкъл'
DEFINE BAR 5 OF ExpCriteria PROMPT '"всфхе" дхяоюмяепхгюжхх'
DEFINE BAR 6 OF ExpCriteria PROMPT '"всфхе" гюярпюунбюммше'
DEFINE BAR 7 OF ExpCriteria PROMPT 'цянохрюкхгюжхх <>50%'
ON SELECTION BAR 1 OF ExpCriteria do seldeads
ON SELECTION BAR 2 OF ExpCriteria do seldblgosps
ON SELECTION BAR 3 OF ExpCriteria do selcrosss
ON SELECTION BAR 4 OF ExpCriteria do selprofus
ON SELECTION BAR 5 OF ExpCriteria do seldsps
ON SELECTION BAR 6 OF ExpCriteria do SelGuests
ON SELECTION BAR 7 OF ExpCriteria do sel50P

DEFINE POPUP ExpCritNew MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF ExpCritNew PROMPT 'керюкэмше хяундш'
DEFINE BAR 2 OF ExpCritNew PROMPT 'онбрнпмше цняохрюкхгюжхх' SKIP 
DEFINE BAR 3 OF ExpCritNew PROMPT 'оепеяевемхъ бхднб ледонлных' SKIP 
DEFINE BAR 4 OF ExpCritNew PROMPT 'нранп он опнтхкъл' SKIP 
DEFINE BAR 5 OF ExpCritNew PROMPT '"всфхе" дхяоюмяепхгюжхх' SKIP 
DEFINE BAR 6 OF ExpCritNew PROMPT '"всфхе" гюярпюунбюммше' SKIP 
DEFINE BAR 7 OF ExpCritNew PROMPT 'цянохрюкхгюжхх <>50%' SKIP 
ON SELECTION BAR 1 OF ExpCritNew do seldeadsnew
*ON SELECTION BAR 2 OF ExpCritNew do seldblgosps
*ON SELECTION BAR 3 OF ExpCritNew do selcrosss
*ON SELECTION BAR 4 OF ExpCritNew do selprofus
*ON SELECTION BAR 5 OF ExpCritNew do seldsps
*ON SELECTION BAR 6 OF ExpCritNew do SelGuests
*ON SELECTION BAR 7 OF ExpCritNew do sel50P

DEFINE POPUP popBuch MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1  OF popBuch PROMPT 'дхттепемжхпнбюммши ондсьебни мнплюрхб'
DEFINE BAR 2  OF popBuch PROMPT '\-'
DEFINE BAR 3  OF popBuch PROMPT 'опнялнрп юоят' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 4  OF popBuch PROMPT 'опнялнрп опхкнфемхъ ╧4' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 5  OF popBuch PROMPT 'оевюрэ юоят (бюпхюмр 1)'  SKIP
DEFINE BAR 6  OF popBuch PROMPT 'оевюрэ юоят (бюпхюмр 2)'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 7  OF popBuch PROMPT 'юмюкхг бшонкмемхъ окюмнбшу назелнб'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 8  OF popBuch PROMPT '\-'
DEFINE BAR 9  OF popBuch PROMPT 'ятнплхпнбюрэ опхкнфемхе ╧4' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 10 OF popBuch PROMPT 'гюцпсгйю юбюмянб' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 11 OF popBuch PROMPT 'тнплхпнбюмхе опхкнфемхи ╧1'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 12 OF popBuch PROMPT '\-'
DEFINE BAR 13 OF popBuch PROMPT 'ябндмюъ беднлнярэ опхкнфемхъ 4 (VER.1)'
DEFINE BAR 14 OF popBuch PROMPT 'ябндмюъ беднлнярэ опхкнфемхъ 4 (VER.2)'
DEFINE BAR 15 OF popBuch PROMPT '\-'
DEFINE BAR 16 OF popBuch PROMPT 'гюцпсгйю ME-тюикнб'
DEFINE BAR 17 OF popBuch PROMPT 'гюцпсгйю тюикнб хлонпрю'
DEFINE BAR 18 OF popBuch PROMPT 'опхкнфемхе й янцюгс ╧1'

ON SELECTION BAR 1 OF popBuch DO FORM ViewDifNorm
ON SELECTION BAR 3 OF popBuch DO FORM ViewAPSF
ON SELECTION BAR 4 OF popBuch DO FORM IIF(tdat1<{01.07.2014}, 'ViewPr4', 'Viewpr4n')
ON SELECTION BAR 5 OF popBuch DO MakeAPSF
ON SELECTION BAR 6 OF popBuch DO MakeAPSF2
ON SELECTION BAR 7 OF popBuch DO VolControls
ON SELECTION BAR 9 OF popBuch DO MakePr4
ON BAR 10 OF popBuch ACTIVATE POPUP popAvances
ON SELECTION BAR 11 OF popBuch DO MakePrilN1
ON SELECTION BAR 13 OF popBuch DO SvodPr4
ON SELECTION BAR 14 OF popBuch DO SvodPr4V2
ON SELECTION BAR 16 OF popBuch DO LoadMeFiles
ON SELECTION BAR 17 OF popBuch DO LoadImpFiles
ON SELECTION BAR 18 OF popBuch DO Pril1S7

DEFINE POPUP popAvances MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF popAvances PROMPT 'онксвеммши б нрвермнл оепхнде'
DEFINE BAR 2 OF popAvances PROMPT 'онксвеммши б рейсыел леяъже'
ON SELECTION BAR 1 OF popAvances DO AvansPeriod
ON SELECTION BAR 2 OF popAvances DO AvansMonth 

DEFINE POPUP popTuneUp MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 1  OF popTuneUp PROMPT 'бшанп нрвермнцн оепхндю' 
DEFINE BAR 2  OF popTuneUp PROMPT 'оюпюлерпш оевюрх' 
DEFINE BAR 3  OF popTuneUp PROMPT '\-'
DEFINE BAR 4  OF popTuneUp PROMPT 'оепехмдейяюжхъ ад мях'
DEFINE BAR 5  OF popTuneUp PROMPT 'оепехмдейяюжхъ пюанвху аюг'
DEFINE BAR 6  OF popTuneUp PROMPT '\-'
DEFINE BAR 7  OF popTuneUp PROMPT 'мюярпнийю пюанвху дхпейрнпхи'
DEFINE BAR 8  OF popTuneUp PROMPT '\-'
DEFINE BAR 9  OF popTuneUp PROMPT 'союйнбюрэ тюикш ньханй' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 10 OF popTuneUp PROMPT 'намскхрэ тюикш ньханй' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 11 OF popTuneUp PROMPT 'сдюкхрэ CTRL-ЙХ' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 12 OF popTuneUp PROMPT 'сдюкхрэ тюикш гюопнянб' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 13 OF popTuneUp PROMPT 'сдюкхрэ опнрнйнкш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 14 OF popTuneUp PROMPT 'сдюкхрэ Mc-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 15 OF popTuneUp PROMPT 'сдюкхрэ Mk-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 16 OF popTuneUp PROMPT 'сдюкхрэ Mt-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 17 OF popTuneUp PROMPT 'сдюкхрэ тюикш b_flk' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 18 OF popTuneUp PROMPT 'сдюкхрэ тюикш b_mek' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 19 OF popTuneUp PROMPT 'сдюкхрэ BAK-тюикш' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 20 OF popTuneUp PROMPT 'пюглнпнгхрэ лщй' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 21 OF popTuneUp PROMPT '\-'
DEFINE BAR 22 OF popTuneUp PROMPT '"янапюрэ" тюик рюпхтю' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 23 OF popTuneUp PROMPT 'намнбкемхе тюикю POLIS_DP' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 24 OF popTuneUp PROMPT 'намнбкемхе тюикю POLIS_H' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 25 OF popTuneUp PROMPT 'йнппейрхпнбйю ярпсйрспш пюанвху ад' SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 26 OF popTuneUp PROMPT 'союйнбюрэ ябндмше аюгш'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 27 OF popTuneUp PROMPT 'гюонкмемхе фспмюкю юйрнб'  SKIP FOR !INLIST(gcUser,'OMS','USR')
DEFINE BAR 28 OF popTuneUp PROMPT 'гюонкмемхе онкъ FIL_ID'
DEFINE BAR 29 OF popTuneUp PROMPT 'напюанрю мнлепмхйю'
IF m.IsAdmin
 DEFINE BAR 30 OF popTuneUp PROMPT '\-'
 DEFINE BAR 31 OF popTuneUp PROMPT 'яцемепхпнбюрэ тюик бепяхх'
 DEFINE BAR 32 OF popTuneUp PROMPT 'юйрсюкхгхпнбюрэ мях'
 DEFINE BAR 33 OF popTuneUp PROMPT 'хлонпр S6'
 DEFINE BAR 34 OF popTuneUp PROMPT 'хлонпр ME-Files'
 DEFINE BAR 35 OF popTuneUp PROMPT 'щйяонпр б дхюянтр'
 DEFINE BAR 36 OF popTuneUp PROMPT 'гюонкмхрэ VMP'
 DEFINE BAR 37 OF popTuneUp PROMPT '\-'
 DEFINE BAR 38 OF popTuneUp PROMPT 'хлонпр S2'
ENDIF 

ON SELECTION BAR 1  OF popTuneUp DO FORM SetPeriod
ON SELECTION BAR 2  OF popTuneUp goApp.doForm('set_print','settings')
ON SELECTION BAR 4  OF popTuneUp DO ComReind
ON SELECTION BAR 5  OF popTuneUp DO BasReind
ON SELECTION BAR 7  OF popTuneUp DO FORM TuneBase
ON SELECTION BAR 9  OF popTuneUp DO PackBD
ON SELECTION BAR 10 OF popTuneUp DO ZapEFiles
ON SELECTION BAR 11 OF popTuneUp DO DelAllCtrl
ON SELECTION BAR 12 OF popTuneUp DO DelAllZapros
ON SELECTION BAR 13 OF popTuneUp DO DelAllProtocols
ON SELECTION BAR 14 OF popTuneUp DO DelMcFiles
ON SELECTION BAR 15 OF popTuneUp DO DelMkFiles
ON SELECTION BAR 16 OF popTuneUp DO DelMtFiles
ON SELECTION BAR 17 OF popTuneUp DO DelAllBFlk
ON SELECTION BAR 18 OF popTuneUp DO DelAllBMek
ON SELECTION BAR 19 OF popTuneUp DO DelBakFiles
ON SELECTION BAR 20 OF popTuneUp DO DeFrMek
ON SELECTION BAR 22 OF popTuneUp DO MakeTarif
ON SELECTION BAR 23 OF popTuneUp DO AppPolisDP
ON SELECTION BAR 24 OF popTuneUp DO AppPolisH
ON SELECTION BAR 25 OF popTuneUp DO CorStruct
ON SELECTION BAR 26 OF popTuneUp DO PackBDSv
ON SELECTION BAR 27 OF popTuneUp DO ActsJournal
ON SELECTION BAR 28 OF popTuneUp DO FillFilId
ON SELECTION BAR 29 OF popTuneUp DO MakeOutS
IF m.IsAdmin
 ON SELECTION BAR 31 OF popTuneUp DO GenVer
 ON SELECTION BAR 32 OF popTuneUp DO ActNSI
 ON SELECTION BAR 33 OF popTuneUp DO ImpS6
 IF m.qcod='S2'
  ON SELECTION BAR 34 OF popTuneUp DO ImpS2Me
 ELSE 
  ON SELECTION BAR 34 OF popTuneUp DO ImpS6Me
 ENDIF 
 ON SELECTION BAR 35 OF popTuneUp DO ExpDsoft
 ON SELECTION BAR 36 OF popTuneUp DO FillVmp
 ON SELECTION BAR 38 OF popTuneUp DO ImpS2
ENDIF 

DEFINE POPUP popMedSpr MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 1 OF popMedSpr PROMPT 'рюпхт'
DEFINE BAR 2 OF popMedSpr PROMPT 'лйа-10'
DEFINE BAR 3 OF popMedSpr PROMPT 'мнплюрхбмше назелш сяксц '
DEFINE BAR 4 OF popMedSpr PROMPT 'опнтхкх ледхжхмяйни онлных'

IF m.IsNotePad = .F.
 ON SELECTION BAR 1 OF popMedSpr DO FORM TarifN
 ON SELECTION BAR 2 OF popMedSpr DO FORM Mkb10
 ON SELECTION BAR 3 OF popMedSpr DO FORM CodKU
ELSE 
 ON SELECTION BAR 1 OF popMedSpr DO FORM TarifN600
 ON SELECTION BAR 2 OF popMedSpr DO FORM Mkb10600
 ON SELECTION BAR 3 OF popMedSpr DO FORM CodKU600
ENDIF 
ON SELECTION BAR 4 OF popMedSpr DO FORM ViewLicAll

DEFINE POPUP popTuneUp2 MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 1  OF popTuneUp2 PROMPT 'йнппейрхпнбйю ябндмни аюгш' 
DEFINE BAR 2  OF popTuneUp2 PROMPT 'намскхрэ тюикш щйяоепрхгш' 
ON SELECTION BAR 1  OF popTuneUp2 Do CorSvBases
ON SELECTION BAR 2  OF popTuneUp2 Do KillMeFiles

SET SYSMENU AUTOMATIC
SET SYSMENU ON

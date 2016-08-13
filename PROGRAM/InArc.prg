PROCEDURE InArc
lñTitle = 'Ïåðåâîä èíôîðìàöèè â àðõèâ.'
lñSubTitle = ''
loTherm = NewObject("_thermometer","_therm",'',lñTitle)
loTherm.Show()
loTherm.Update(0,lñSubTitle)

select * from lpui1 into cursor cLpu1 order by mcod WHERE iscur
use dbf() in 0 alias cLpu again
use in cLpu1
use in lpui1
select cLpu
lnNumLpu = reccount()
go top

lcPeriod = strtran(str(oSettings.cur_month,2),' ','0') + right(str(oSettings.cur_year),1)
lcPer = strtran(str(oSettings.cur_month,2),' ','0') + right(str(oSettings.cur_year,4),1)
lcPeriodMonth = str(oSettings.cur_year,4) + strtran(str(oSettings.cur_month,2),' ','0')
lnMonth = oSettings.cur_month

lnLpu = 0
select cLpu

scan

	lcMcod = cLpu.mcod
	lcLpuName = alltrim(cLpu.name)
	lnLpu = lnLpu + 1

	lnPercent = round(lnLpu/lnNumLpu,2)*100
	loTherm.Update(lnPercent,lcLpuName)

	lcPname = oSettings.inlpu + 'P' + cLpu.mcod + '.DBF'
	lcSname = oSettings.inlpu + 'S' + cLpu.mcod + '.DBF'
	lcDname = oSettings.posmed + 'D' + cLpu.mcod + '.DBF'
	lcEname = oSettings.inlpu + 'E' + cLpu.mcod + '.DBF'

	lcPAname = oSettings.database + 'P' + cLpu.mcod + '.' + lcPer
	lcSAname = oSettings.database + 'S' + cLpu.mcod + '.' + lcPer
	lcDAname = oSettings.database + 'D' + cLpu.mcod + '.' + lcPer
	lcEAname = oSettings.database + 'E' + cLpu.mcod + '.' + lcPer

	IF NOT ( file(lcPname) AND file(lcSname) AND file(lcEname) AND file(lcDname) )
		loop
	ENDIF

	create cursor cEfile1 (f c(1),c_err c(2),recid n(6), et c(1), newcod n(6), newku n(3), newtip c(1),period c(6))
	use dbf() in 0 alias cEfile again
	use in cEfile1

	SELECT 0
	lnAlias = SELECT()
	USE (lcEname)
	SELECT * FROM (lcEname) INTO CURSOR cE
	lcdbf = dbf()
	SELECT (lnAlias)
	USE
	SELECT cEfile
	APPEND FROM (lcdbf)
	USE IN cE

	SELECT 0
	lnAlias = SELECT()
	USE (lcDname)
	SELECT *,IIF(BETWEEN(nnn,40,49),'  ','WE') as c_err FROM (lcDname) INTO CURSOR cD
	lcdbf = dbf()
	SELECT (lnAlias)
	USE
	SELECT cEfile
	APPEND FROM (lcdbf)
	USE IN cD

	SELECT cEfile
	replace et with '1',period with lcPeriodMonth all
	replace f with 'S' for empty(f)
	copy to (lcEAname) as 866 type fox2x
	USE IN cEfile

	SELECT 0
	USE (lcPname)
	copy to (lcPAname) as 866 type fox2x
	USE

	SELECT 0
	USE (lcSname)
	copy to (lcSAname) as 866 type fox2x
	USE

endscan

use in cLpu

loTherm.release
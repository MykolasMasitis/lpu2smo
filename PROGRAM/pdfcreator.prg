*PDFCREATE("C:\devvfp9\tek-tips\sample.txt","C:\devvfp9\tek-tips\sample.PDF",0,"Courier",10)

FUNCTION PDFCREATE
	PARAMETERS m.INFILENAME,m.OUTFILENAME,m.PAGELENGTH,m.NAMEFONT,m.SIZEFONT
	PRIVATE m.INFILENAME,m.OUTFILENAME
	PRIVATE m.CRLF
	PRIVATE XREF_END_CHAR
	PRIVATE PDFOBJECT_END
	PRIVATE PDFOBJECT_BEGIN
	PRIVATE PDFXREFMARKER
	PRIVATE m.STRPAGES
	PRIVATE m.NOPAGES
	PRIVATE m.STARTAT
	PRIVATE STRFONTNAME
	PRIVATE m.NOLINES
	PRIVATE m.CURRPAGE
	DECLARE ARRXREF(10)
	DECLARE	ARRDATA(29)
	DECLARE ARRTMP(1)
	IF PCOUNT() < 5
		m.SIZEFONT = 12
	ENDIF
	IF PCOUNT() < 4
		m.NAMEFONT = "Courier-Bold"
	ENDIF
	IF PCOUNT() < 3
		m.PAGELENGTH = 0  && 0 is used to force control to look for chr(12) for new page
	ENDIF

	m.STARTAT = 1
	STRFONTNAME = m.NAMEFONT
	PDFOBJECT_END = "endobj"
	PDFOBJECT_BEGIN = " 0 obj"
	PDFXREFMARKER	= "PDFXREFMARKER"
	XREF_END_CHAR = " 00000 n"

	m.OBJECTCOUNT = 6  && AT LEAST 6 OBJECT BLOCKS
	m.CRLF = CHR(13)+CHR(10)

	ARRXREF(1)  = "xref"
	ARRXREF(2)  = "0 10"
	ARRXREF(3)  = "0000000000 65535 f"
	ARRXREF(4)  = PDFXREFMARKER  && position of obj 1 - title block
	ARRXREF(5)  = PDFXREFMARKER  && position of obj 2 - catalogue
	ARRXREF(6)  = PDFXREFMARKER  && position of obj 3 - describes the number of pages
	ARRXREF(7)  = PDFXREFMARKER  && position of obj 4 - fonts
	ARRXREF(8)  = PDFXREFMARKER  && position of obj 5 - encoding
	ARRXREF(9)  = PDFXREFMARKER  && position of obj 6 - font description ?
	ARRXREF(10) = PDFXREFMARKER  && position of obj 7 - describes the container for the pages

	PDFINITIALISE()

	m.NOPAGES = PDFREADTEXTFILE(m.INFILENAME)

	m.STARTAT = 1
	m.STRPAGES = ""
	FOR m.CURRPAGE = 1 TO m.NOPAGES
		PDFCREATEPAGE()
	NEXT

	** this adds in the objects numbered 2 & 3 - at the end because the pdf does not HAVE to be in numerical order
	PDFADDCATALOGDETAILS()

	ARRXREF(5) = ARRXREF(ALEN(ARRXREF))
	ARRXREF(ALEN(ARRXREF)) = ""


	** INSERT THE UNIQUE POSITION INTO THE XREF TABLE
	ARRXREF(6) = PDFXREFMARKER
	m.OBJECTCOUNT = m.OBJECTCOUNT + 1
	** INSERT THE NUMBER OF 'OBJECTS' INTO THE XREFS TABLE
	ARRXREF(2) = "0 " + ALLTRIM(STR(m.OBJECTCOUNT))

	** ADD THE FOOTER DETAILS AND THEN APPEND TO THE DATA ARRAY
	PDFFOOTER()

	** WRITE THE DATA ARRAY OUT TO THE PDF FILE
	PDFWRITE(m.OUTFILENAME)

	RETURN(.T.)

*!******************************************************************************
*!
*! Procedure PDFREADTEXTFILE
*!
*!******************************************************************************
FUNCTION PDFREADTEXTFILE
	PARAMETERS m.INFILENAME
	PRIVATE m.INFILENAME,m.NOPAGES,I,m.LINENUMBER
	** we can read the incoming text file into a string and then fill an array with the string lines
	m.NOLINES = ALINES(ARRTMP,FILETOSTR(m.INFILENAME))
	m.NOPAGES = 1
	m.LINENUMBER = 0
	FOR I = 1 TO m.NOLINES
		m.LINENUMBER = m.LINENUMBER +1
		** TRANSFORM THE INPUT STRING TO GET RID OF HIEND ASCII CHARACTERS
		ARRTMP(I) = STRTRAN(ARRTMP(I),CHR(205),"=")
		ARRTMP(I) = STRTRAN(ARRTMP(I),CHR(186),"|")
		ARRTMP(I) = STRTRAN(ARRTMP(I),CHR(187),"+")
		ARRTMP(I) = STRTRAN(ARRTMP(I),CHR(188),"+")
		ARRTMP(I) = STRTRAN(ARRTMP(I),CHR(200),"+")
		ARRTMP(I) = STRTRAN(ARRTMP(I),CHR(201),"+")
		** Also watch out for the few characters that ean things to acrobat reader '(', '(' and '\'
		ARRTMP(I) = STRTRAN(ARRTMP(I),"\","\\")
		ARRTMP(I) = STRTRAN(ARRTMP(I),"(","\(")
		ARRTMP(I) = STRTRAN(ARRTMP(I),")","\)")
		IF CHR(12) $ ARRTMP(I) .OR. (m.LINENUMBER >= m.PAGELENGTH .AND. m.PAGELENGTH > 0)
			m.NOPAGES = m.NOPAGES +1
			m.LINENUMBER = 0
		ENDIF
	NEXT
	RETURN(m.NOPAGES)

*!******************************************************************************
*!
*! Procedure PDFCREATEPAGE
*!
*!******************************************************************************
FUNCTION PDFCREATEPAGE
	PRIVATE I,m.STARTSIZE,m.NOROWS,m.LINENUMBER,m.STREAMLENGTH
	m.OBJECTCOUNT = m.OBJECTCOUNT + 1
	m.STRPAGES = m.STRPAGES + " " + ALLTRIM(STR(m.OBJECTCOUNT)) + " 0 R"
	*INTLEN = LEN(STR(m.OBJECTCOUNT)) + LEN(STR(m.OBJECTCOUNT+1))
	** KEEP A TRACK OF WHERE THE DATA BLOCK STARTED
	m.STARTSIZE = ALEN(ARRDATA)
	** ADD 18 LINES TO THE DATA ARRAY
	DIMENSION ARRDATA(ALEN(ARRDATA)+18)
	** FILL IN THE DESCRIPTION OF THE BLOCK
	ARRDATA(m.STARTSIZE+ 1) = ALLTRIM(STR(m.OBJECTCOUNT)) + PDFOBJECT_BEGIN
	ARRDATA(m.STARTSIZE+ 2) = "<<"
	ARRDATA(m.STARTSIZE+ 3) = "/Type /Page"
	ARRDATA(m.STARTSIZE+ 4) = "/Parent 3 0 R"
	ARRDATA(m.STARTSIZE+ 5) = "/Resources 6 0 R"

	m.OBJECTCOUNT = m.OBJECTCOUNT + 1
	ARRDATA(m.STARTSIZE+ 6) = "/Contents " + ALLTRIM(STR(m.OBJECTCOUNT)) + " 0 R"
	ARRDATA(m.STARTSIZE+ 7) = ">>"
	ARRDATA(m.STARTSIZE+ 8) = PDFOBJECT_END

	DIMENSION ARRXREF(ALEN(ARRXREF)+1)
	ARRXREF(ALEN(ARRXREF)) = PDFXREFMARKER

	ARRDATA(m.STARTSIZE+ 9) = ALLTRIM(STR(m.OBJECTCOUNT)) + PDFOBJECT_BEGIN
	ARRDATA(m.STARTSIZE+ 10) = "<<"
	m.OBJECTCOUNT = m.OBJECTCOUNT + 1
	** the length if this object (distance from BT to just before endstream) is recorded in the next object
	** having just incremented objectcount
	ARRDATA(m.STARTSIZE+ 11) = "/Length " + ALLTRIM(STR(m.OBJECTCOUNT)) + " 0 R"
	ARRDATA(m.STARTSIZE+ 12) = ">>"
	ARRDATA(m.STARTSIZE+ 13) = "stream"
	ARRDATA(m.STARTSIZE+ 14) = "BT" && begin text
	** START STRACKING THE LENGTH OF THE STREAM
	m.STREAMLENGTH = LEN(ARRDATA(m.STARTSIZE+ 14))+2

	ARRDATA(m.STARTSIZE+ 15) = "/F1 "+ALLTRIM(STR(m.SIZEFONT))+" Tf" && fontsize
	m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.STARTSIZE+ 15))+2

	*ARRDATA(m.STARTSIZE+ 16) = "1 0 0 1 50 802 Tm"
	** this defines start point         lft top
	ARRDATA(m.STARTSIZE+ 16) = "1 0 0 1 50 580 Tm"
	
	m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.STARTSIZE+ 16))+2

	ARRDATA(m.STARTSIZE+ 17) = ALLTRIM(STR(m.SIZEFONT*1.2))+" TL" && linefeed spacing
	m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.STARTSIZE+ 17))+2

	** we need to scan the tmpData array for a given page
	I = m.STARTAT
	m.LINENUMBER = 0
	m.FLG = .T.
	DO WHILE I <= ALEN(ARRTMP) .AND. m.FLG
		m.LINENUMBER = m.LINENUMBER +1
		** add an entry to the data array
		DIMENSION ARRDATA(ALEN(ARRDATA)+1)
		ARRDATA(ALEN(ARRDATA)) = "T* ("+TRIM(ARRTMP(I))+") Tj"
		** add the length of this string to the stream length
		m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(ALEN(ARRDATA)))+2
		IF CHR(12) $ ARRTMP(I) .OR. (m.LINENUMBER >= m.PAGELENGTH .AND. m.PAGELENGTH > 0)
			m.STARTAT = I+1
			m.FLG = .F.
		ENDIF
		I = I + 1
	ENDDO

	m.STARTSIZE = ALEN(ARRDATA)
	** ADD ANOTHER 18 LINES TO THE DATA ARRAY
	DIMENSION ARRDATA(ALEN(ARRDATA)+18)
	ARRDATA(ALEN(ARRDATA)) = ""  && BLANK LINE?
	ARRDATA(m.STARTSIZE+ 1) = "ET"  && end text
	** add the length of this string to the stream length
	m.STREAMLENGTH = m.STREAMLENGTH + LEN(ARRDATA(m.STARTSIZE+ 1))+2
	ARRDATA(m.STARTSIZE+ 2) = "endstream"
	ARRDATA(m.STARTSIZE+ 3) = PDFOBJECT_END

	DIMENSION ARRXREF(ALEN(ARRXREF)+1)
	ARRXREF(ALEN(ARRXREF)) = PDFXREFMARKER

	ARRDATA(m.STARTSIZE+ 4) = ALLTRIM(STR(m.OBJECTCOUNT)) + PDFOBJECT_BEGIN
	ARRDATA(m.STARTSIZE+ 5) = ALLTRIM(STR(m.STREAMLENGTH))

	ARRDATA(m.STARTSIZE+ 6) = PDFOBJECT_END


	DIMENSION ARRXREF(ALEN(ARRXREF)+1)
	ARRXREF(ALEN(ARRXREF)) = PDFXREFMARKER

	RETURN(.T.)


*!******************************************************************************
*!
*! Procedure PDFWRITE
*!
*!******************************************************************************
FUNCTION PDFWRITE
	PARAMETERS m.OUTFILENAME
	PRIVATE I, m.OUTFILENAME,m.STRING,m.TEMPBIT,X,XREFINDEX,m.OFFSETPOSN
	DECLARE ARRXREFS(m.OBJECTCOUNT-1) && SET UP A NEW XREFS TABLE TO HOLD POSITIONS OF ALL BLOCKS
	FOR X = 1 TO ALEN(ARRXREFS)
		ARRXREFS(X)=0
	NEXT
	m.OFFSETPOSN = 0
	m.STRING = ""
	XREFINDEX = 0
	FOR I = 1 TO ALEN(ARRDATA)
		m.TEMPBIT = ARRDATA(I)
		IF TYPE("m.tempbit") = "C" .AND. !EMPTY(m.TEMPBIT)
			DO CASE
			CASE RIGHT(UPPER(m.TEMPBIT),LEN(PDFOBJECT_BEGIN)) = UPPER(PDFOBJECT_BEGIN)
				** WE HAVE HERE THE STARTING POSITION OF AN OBJECT - WHICH WE NEED IN THE XREFS TABLE
				** the xref we are interested in is the value of the bit before the PDFOBJECTBEGIN
				X =  VAL(LEFT(m.TEMPBIT,LEN(m.TEMPBIT)-LEN(PDFOBJECT_BEGIN)))
				ARRXREFS(X) = LEN(m.STRING)
			CASE UPPER(m.TEMPBIT) = PDFXREFMARKER && THIS IS AN XREF IN WAITING...
				XREFINDEX = XREFINDEX +1
				m.TEMPBIT = RIGHT("0000000000"+ALLTRIM(STR(ARRXREFS(XREFINDEX))),10) + XREF_END_CHAR
			CASE UPPER(m.TEMPBIT) == "XREF"
				** make a note of the position that the xref block starts
				m.OFFSETPOSN = LEN(m.STRING)
			CASE UPPER(m.TEMPBIT) == "M.OFFSETPOSN"
				** put the position of the xref block into the string
				m.TEMPBIT = ALLTRIM(STR(m.OFFSETPOSN))
			ENDCASE
			m.STRING = m.STRING + m.TEMPBIT+m.CRLF
		ENDIF
	NEXT
	STRTOFILE(m.STRING,m.OUTFILENAME)
	RETURN(.T.)

*!******************************************************************************
*!
*! Procedure PDFADDCATALOGDETAILS
*!
*!******************************************************************************
FUNCTION PDFADDCATALOGDETAILS
	PRIVATE m.STARTSIZE
	m.STARTSIZE = ALEN(ARRDATA)
	** ADD 15 LINES TO THE DATA ARRAY
	DIMENSION ARRDATA(ALEN(ARRDATA)+15)
	ARRDATA(ALEN(ARRDATA)) = ""
	ARRDATA(m.STARTSIZE + 1) = "2" + PDFOBJECT_BEGIN
	ARRDATA(m.STARTSIZE + 2) = "<<"
	ARRDATA(m.STARTSIZE + 3) = "/Type /Catalog"
	ARRDATA(m.STARTSIZE + 4) = "/Pages 3 0 R"   && the number of pages is in section/obj number 3
	ARRDATA(m.STARTSIZE + 5) = "/PageLayout /OneColumn"
	ARRDATA(m.STARTSIZE + 6) = ">>"
	ARRDATA(m.STARTSIZE + 7) = PDFOBJECT_END
	ARRDATA(m.STARTSIZE + 8) = "3" + PDFOBJECT_BEGIN
	ARRDATA(m.STARTSIZE + 9) = "<<"
	ARRDATA(m.STARTSIZE + 10) = "/Type /Pages"
	ARRDATA(m.STARTSIZE + 11) = "/Count " + ALLTRIM(STR(m.NOPAGES))
	**                                           wid hgt  multiply inches by about 72 or mm by 2.785
	ARRDATA(m.STARTSIZE + 12) = "/MediaBox [ 0 0 827 594 ]"
	ARRDATA(m.STARTSIZE + 13) = "/Kids [" + m.STRPAGES + " ]"
	ARRDATA(m.STARTSIZE + 14) = ">>"
	ARRDATA(m.STARTSIZE + 15) = PDFOBJECT_END
	RETURN(.T.)

*!******************************************************************************
*!
*! Procedure PDFINITIALISE
*!
*!******************************************************************************
FUNCTION PDFINITIALISE
	ARRDATA(1) = "%PDF-1.2 "
	ARRDATA(2) = "%MGF"

	ARRDATA(3) = "1" + PDFOBJECT_BEGIN
	ARRDATA(4) = "<<"
	ARRDATA(5) = "/Creator  (Mikhail Ryabov)"
	ARRDATA(6) = "/Producer (Independent programmer)"
	ARRDATA(7) = "/Title    (Direct PDF-Creator)"
	ARRDATA(8) = ">>"
	ARRDATA(9) = PDFOBJECT_END

	ARRDATA(10) = "4" + PDFOBJECT_BEGIN
	ARRDATA(11) = "<<"
	ARRDATA(12) = "/Type /Font"
	ARRDATA(13) = "/Subtype /Type1"
	ARRDATA(14) = "/Name /F1"
	ARRDATA(15) = "/Encoding 5 0 R"
	ARRDATA(16) = "/BaseFont /" + STRFONTNAME
	ARRDATA(17) = ">>"
	ARRDATA(18) = PDFOBJECT_END

	ARRDATA(19) = "5" + PDFOBJECT_BEGIN
	ARRDATA(20) = "<<"
	ARRDATA(21) = "/Type /Encoding"
	ARRDATA(22) = "/BaseEncoding /WinAnsiEncoding"
	ARRDATA(23) = ">>"
	ARRDATA(24) = PDFOBJECT_END

	ARRDATA(25) = "6" + PDFOBJECT_BEGIN
	ARRDATA(26) = "<<"
	ARRDATA(27) = "/Font << /F1 4 0 R   >>  /ProcSet [ /PDF  /Text ]"
	ARRDATA(28) = ">>"
	ARRDATA(29) = PDFOBJECT_END
	RETURN(.T.)

*!******************************************************************************
*!
*! Procedure PDFFOOTER
*!
*!******************************************************************************
FUNCTION PDFFOOTER
	PRIVATE I
	DIMENSION ARRXREF(ALEN(ARRXREF)+9)
	ARRXREF(ALEN(ARRXREF)-8) = "trailer"
	ARRXREF(ALEN(ARRXREF)-7) = "<<"
	ARRXREF(ALEN(ARRXREF)-6) = "/Size " + ALLTRIM(STR(m.OBJECTCOUNT))
	ARRXREF(ALEN(ARRXREF)-5) = "/Root 2 0 R"
	ARRXREF(ALEN(ARRXREF)-4) = "/Info 1 0 R"
	ARRXREF(ALEN(ARRXREF)-3) = ">>"
	ARRXREF(ALEN(ARRXREF)-2) = "startxref"
	ARRXREF(ALEN(ARRXREF)-1) = "M.OFFSETPOSN" && REPLACE THIS VALUE AS WRITTEN OUT TO TEXT FILE
	ARRXREF(ALEN(ARRXREF)) = "%%EOF"
	FOR I = 1 TO ALEN(ARRXREF)
		IF !EMPTY(ARRXREF(I))
			DIMENSION ARRDATA(ALEN(ARRDATA)+1)
			ARRDATA(ALEN(ARRDATA)) = ARRXREF(I)
		ENDIF
	NEXT
	RETURN(.T.)


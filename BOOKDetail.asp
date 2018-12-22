<!-- #INCLUDE FILE="bookLib.Asp" -->

<%
Dim strPageTitle, strTitle, strMeta
Dim sSQL, rsBOOK
Dim strYear, strCsrdNo, strLName, strFName, strName

Dim description, samplepara, review, userreview

GetFields

OpenDB

'response.write "SQL = " & sSql
'endit

OpenRS rsBOOK, sSql

rsBOOK.MoveFirst
do while NOT rsBOOK.EOF
strPageTitle = "RTS OnLine"
strTitle = "Detail information on " & rsBOOK("booktitle")
strMeta = ""

	Header strPageTitle, strTitle, strMeta
	response.write "<table border='0' width='90%'>"
	response.write "<tr class='color'><td width=150>ISBN Number:</td><td>" & rsBOOK("ISBNNumber") & "</td></tr>"
	
	response.write "<tr><td>Title:</td><td>" & rsBOOK("booktitle") & "</td></tr>"
	response.write "<tr class='color'><td>Author:</td><td>" & rsBOOK("LName") & ", " & rsBOOK("FName") & "</td></tr>"
	response.write "<tr><td>Publisher:</td><td>" & rsBOOK("Publisher") & "</td></tr>"
	response.write "<tr class='color'><td>Category:</td><td>" & rsBOOK("Category") & "</td></tr>"
	description = rsBOOK("description")
	if len(rsBOOK("description")) > 75 then
	   requirement = Left(rsBOOK("description"), 75) & "<br><a href='#' onmousedown=window.open('bookbookpopup.asp?isbn=" & rsBOOK("ISBNNumber") &  "&description=1','blank','height=350,width=500,menubar=0,resizable=0,scrollbars=1,titlebar=0')>... view full description ...</A>"
	else
		description = rsBOOK("description")
	End if

	response.write "<tr><td>Description:</td><td>" & description & "</td></tr>"
	response.write "<tr class='color'><td>Pages:</td><td>" & rsBOOK("pages") & "</td></tr>"
	if len(rsBOOK("samplepara")) > 75 then
	   samplepara = Left(rsBOOK("samplepara"), 75) & "<br><a href='#' onmousedown=window.open('bookpopup.asp?isbn=" & rsBOOK("ISBNNumber") & "&samplepara=1','blank','height=350,width=500,menubar=0,resizable=0,scrollbars=1,titlebar=0')>... view full paragraph ...</A>"
	else
		samplepara = rsBOOK("samplepara")
	End if

	response.write "<tr><td>Sample Paragraph:</td><td>" & samplepara & "</td></tr>"
	if len(rsBOOK("review")) > 75 then
	   review = Left(rsBOOK("review"), 75) & "<br><a href='#' onmousedown=window.open('bookpopup.asp?isbn=" & rsBOOK("ISBNNumber") & "&review=1','blank','height=350,width=500,menubar=0,resizable=0,scrollbars=1,titlebar=0')>... view whole review ...</A>"
	else
		review = rsBOOK("review")
	End if
	response.write "<tr class='color'><td>Editorial Review:</td><td>" & review & "</td></tr>"
	if len(rsBOOK("userreview")) > 75 then
	   userreview = Left(rsBOOK("userreview"), 75) & "<br><a href='#' onmousedown=window.open('bookpopup.asp?isbn=" & rsBOOK("ISBNNumber") & "&userreview=1','blank','height=350,width=500,menubar=0,resizable=0,scrollbars=1,titlebar=0')>... view full user review ...</A>"
	else
		userreview = rsBOOK("userreview")
	End if
	response.write "<tr ><td>User Review:</td><td>" & userreview & "</td></tr>"
	response.write "</td></tr></table>"

rsBOOK.MoveNext
Loop

Sub GetFields
	sSql = "Select * from booklist where "

	if Request("ISBN") <> "" then
'		debugwrite "CSRD = ." & Request("CSRD") & "."
'		endit
		sSql = sSql & "ISBNNumber = '" & Request("ISBN") & "'"
		exit sub
	End if
End Sub



%>
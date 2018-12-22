<!-- #INCLUDE FILE="bookLib.Asp" -->

<%
Dim strPageTitle, strTitle, strMeta
Dim sSQL, rsBOOK
Dim count, where
strPageTitle = "Book Search"
strTitle = ""
strMeta = ""

Header strPageTitle, strTitle, strMeta
%>
<!-- #INCLUDE FILE="displaybook.asp" -->
<script language=javascript>
function submitSort(crit, type)
{
	document.search.sorttype.value = type;
	document.search.sortheading.value = crit;
	document.search.submit();
}
</script>
<%
if Request("search") = "search" then

OpenDB

%>

<Table cellpadding=2 cellspacing=0 align=center Width=760>
	<tr bgcolor=#336699>
		<td class=texthead><a href="javascript: submitSort('ISBNNumber', '')"><img src=images/asc.gif border=0 alt="sort in ascending order"></a>&nbsp;<b>ISBN </b>&nbsp;<a href="javascript: submitSort('ISBNNumber', 'DESC')"><img src=images/desc.gif border=0  alt="sort in descending order"></a></td>
		<td class=texthead><a href="javascript: submitSort('booktitle', '')"><img src=images/asc.gif border=0 alt="sort in ascending order"></a>&nbsp;<b>Title</b>&nbsp;<a href="javascript: submitSort('booktitle', 'DESC')"><img src=images/desc.gif border=0  alt="sort in descending order"></a></td>
		<td class=texthead><a href="javascript: submitSort('Lname, Fname', '')"><img src=images/asc.gif border=0  alt="sort in ascending order"></a>&nbsp;<b>Author Name</b>&nbsp;<a href="javascript: submitSort('Lname DESC, Fname', 'DESC')"><img src=images/desc.gif  alt="sort in descending order" border=0></a></td>
		<td class=texthead><a href="javascript: submitSort('Publisher', '')"><img src=images/asc.gif border=0  alt="sort in ascending order"></a>&nbsp;<b>Publisher</b>&nbsp;<a href="javascript: submitSort('Publisher', 'DESC')"><img src=images/desc.gif border=0 alt="sort in descending order"></a></td>
		<td class=texthead><a href="javascript: submitSort('PubYear', '')"><img src=images/asc.gif border=0  alt="sort in ascending order"></a>&nbsp;<b>Year </b>&nbsp;<a href="javascript: submitSort('PubYear', 'DESC')"><img src=images/desc.gif border=0 alt="sort in descending order"></a></td>
		<td class=texthead><a href="javascript: submitSort('category', '')"><img src=images/asc.gif border=0  alt="sort in ascending order"></a>&nbsp;<b>Category</b>&nbsp;<a href="javascript: submitSort('category', 'DESC')"><img src=images/desc.gif border=0 alt="sort in descending order"></a></td>
	</tr>

<%
	buildSql
	OpenRS rsBOOK, sSql
	If rsBOOK.Supports(adBookmark) Then 
    	Response.Write "Your Query has returned: " & rsBOOK.RecordCount & " records<br><br>"
	end if
	Do While NOT rsBOOK.EOF
	
	if count Mod 2 <> 1 then
		Response.Write "<tr >"
	else
		Response.Write "<tr>"
	end if
%>


			<td class="list"><a href="BOOKDetail.asp?ISBN=<%=rsBOOK("ISBNNumber")%>" title='click for details'>
				<%=rsBOOK("ISBNNumber")%></a>
			</td>
			<td class="list" width=200><%=rsBOOK("booktitle")%></td>
			<td class="list"><a href="bookList.asp?search=search&lname=<%=rsBOOK("LName")%>"><%=rsBOOK("LName")%></a>, <a href="bookList.asp?search=search&fname=<%=rsBOOK("FName")%>"><%=rsBOOK("FName")%></a></td>
			<td class="list" width=75><%=rsBOOK("Publisher")%>&nbsp;</td>
			<td class="list" align=center><%=rsBOOK("PubYear")%>&nbsp;</td>
			<td class="list"><%=rsBOOK("category")%>&nbsp;</td>
		</tr>
<%		count = count + 1
		rsBOOK.MoveNext
	Loop
	response.write "<tr><td colspan=5>" & rsBOOK.RecordCount & " Records</td></tr></table><BR>" 

	rsBOOK.close
set rsBOOK = Nothing
CloseDatabase GetSess("dbc")
end if

footer


sub buildSql
	sSql = "Select * from booklist "

	if Request("booktitle") <> "" then
		if where <> 1 then
			sSql = sSql & " WHERE "
		else
			sSql = sSql & " AND "
		end if
		where = 1
		sSql = sSql & " booktitle LIKE '%" & Request("booktitle") & "%'"
	end if
	if Request("publisher") <> "" then
		if where <> 1 then
			sSql = sSql & " WHERE "
		else
			sSql = sSql & " AND "
		end if
		where = 1
		sSql = sSql & " Publisher LIKE '%" & Request("publisher") & "%'"
	end if
	if Request("lname") <> "" then
		if where <> 1 then
			sSql = sSql & " WHERE "
		else
			sSql = sSql & " AND "
		end if
		where = 1
		sSql = sSql & " LName LIKE '%" & Request("lname") & "%'"
	end if
	if Request("fname") <> "" then
		if where <> 1 then
			sSql = sSql & " WHERE "
		else
			sSql = sSql & " AND "
		end if
		where = 1
		sSql = sSql & " FName LIKE '%" & Request("fname") & "%'"
	end if
	if Request("year") <> "" then
		if where <> 1 then
			sSql = sSql & " WHERE "
		else
			sSql = sSql & " AND "
		end if
		where = 1
		sSql = sSql & "PubYear = " & CInt(Request("year"))
	end if
	if Request("sortheading") <> "" then
		sSql = sSql & " ORDER BY " & Request("sortheading") & " " & Request("sorttype")
	end if
	if Request("category") <> "" then
		if where <> 1 then
			sSql = sSql & " WHERE "
		else
			sSql = sSql & " AND "
		end if
		where = 1
		sSql = sSql & "category = '" & Request("category") & "'"
	end if
end Sub
%>

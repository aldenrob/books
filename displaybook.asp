<script language=javascript>
function clearForm(form)
{
	form.year.value = '';
	form.title.value = '';
	form.publisher.value = '';
	form.fname.value = '';
	form.lname.value = '';
	form.category.value = '';
}
</script>
<%
Dim stSQL, rsStatus


stSql = "SELECT DISTINCT category FROM booklist"

OpenRS rsStatus, stSql

%>
<form name=search method=post action=bookList.asp>
<input type=hidden value=search name=search>
<input type=hidden value='' name=sorttype>
<input type=hidden value='' name=sortheading>
<div class="big">Sample Database Query Tool</div><p>
<table ALIGN=CENTER border="0" cellspacing="1" cellpadding="2" valign="top" bgcolor="#336699">
					<tr bgcolor="#336699">
						<td bgcolor="#336699"><font color=white>
						Search Criteria:</font></td></tr>

											<tr bgcolor="#ffffff">
						<td align="center" bgcolor="#FFFFFF">
<table align=center class='searchform'>
<tr><td>Title:</td><td><input type=text name=booktitle></td>
<td>Publisher:</td><td><input type=text name=publisher value=<%=Request("Publisher")%>></td></tr>
<tr><td>First Name:</td><td><input type=text name=fname value=<%=Request("fname")%>></td>
<td>Last Name:</td><td><input type=text name=lname value=<%=Request("lname")%>></td></tr>
<tr><td>Category:</td><td>
<select name=category class='paul'>
<option value=''></option>
<%
Do While NOT rsStatus.EOF
Response.Write "<option " 
if rsStatus("category") = Request("category") then
	Response.Write "SELECTED>" & rsStatus("category") & "</option>"
else
	Response.Write ">" & rsStatus("category") & "</option>"
end if
rsStatus.MoveNext
Loop
%>
</select>
</td>
<td>Year:</td><td><input type=text name=year value=<%=Request("year")%>></td></tr>
</table>
</td></tr></table>
<br><center>
<input type=submit value=Search class='button'>&nbsp;<input type=button value=Clear class='button' onClick="javascript: clearForm(search.form)">
</center>
</form>
<%
rsStatus.close
set rsStatus = Nothing


%>

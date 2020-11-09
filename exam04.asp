<% 
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache" 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>總複習總表題型</title>
<link rel="stylesheet" href="ie4.css" type="text/css">
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>
<!--#include file= "connect.asp" -->
<body>
<h2>Please print out and fill in the blanks</h2>
<%
	cell_width  = "250"
	'cell_width  = "150.15pt"
	'cell_height = "22.4pt"
	cell_height = "35"
%>
<table border="1" cellspacing=0 cellpadding=0 >
  <tr height="<%=cell_height%>">
    <td width="<%=cell_width%>"  bgcolor="#CCCCCC" scope="col">Generic Name </td>
    <td width="<%=cell_width%>"  bgcolor="#CCCCCC" scope="col">Brand Name </td>
    <td width="<%=cell_width%>"  bgcolor="#CCCCCC" scope="col">Therapeutic Class </td>
    <td width="<%=cell_width%>"  bgcolor="#CCCCCC" scope="col">Schedule</td>
    <td width="<%=cell_width%>"  bgcolor="#CCCCCC" scope="col">Indication(s)</td>
  </tr>
	<%
	' Int ((200 - 150 + 1) * Rnd + 150) 	would return a random number between 150 and 200

	qt1 = 20
	total_s1 = 0 '填充
	'儲存已取出題目的容器
	redim ary_qt1(qt1) 
	For p = 0 to qt1 - 1
		ary_qt1(p) = 0
	Next 

	'Connect to Database
	Const adCmdText = &H0001
	strsql = "SELECT * FROM tblRxData" 
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open strsql , conn , 3 , , adCmdText
	if not rs.eof then
	Do while (not rs.eof and total_s1 < qt1)
	i=i+1
	'亂數產生，並取不重複出現的值
	rndCheck = 0
	Do While rndCheck = 0
		Randomize Timer
		rndNumber = Int(rs.RecordCount * Rnd)
		match_count = 0 '下列查詢有配合到的總數
		For p = 0 to total_s1 
			if ary_qt1(p) = rndNumber then		
				match_count = match_count + 1
			End if
		Next
		if match_count = 0 then
			ary_qt1(total_s1) = rndNumber
			rndCheck = 1
		End if
	Loop
	'End of 亂數產生
	rs.MoveFirst
	rs.Move rndNumber
	' 挑選顯示模式
	Randomize Timer
  randNum = (CInt(5000 * Rnd) + 1) ' range from 1~5000
	mode = (randNum mod 4)
	Select Case rs("Schedule")
		case 0
			schedule = "&nbsp;"
		case 1
			schedule = "C I"
		case 2
			schedule = "C II"
		case 3
			schedule = "C III"
		case 4
			schedule = "C IV"
		case 5
			schedule = "C V"
		case else
	end select

	select case mode
		Case 0
		%>
		<tr height="<%=cell_height%>">
			<td width="150"><%=rs("GName")%></td>
			<td width="150">&nbsp;</td>
			<%
				if request("Adv") = "" then
			%>
			<td width="150"><%=rs("TClass")%></td>
			<%
				else
			%>
			<td width="150">&nbsp;</td>
			<%
				end if
			
			%>
			<td width="50" >&nbsp;</td>
			<td width="200">&nbsp;</td>
		</tr>
		<%
		Case 1
		%>
		<tr height="<%=cell_height%>">
			<td width="150"><%=rs("GName")%></td>
			<td width="150">&nbsp;</td>
			<td width="150">&nbsp;</td>
			<td width="50">&nbsp;</td>
			<%
				if request("Adv") = "" then
			%>
			<td width="200"><%=rs("Indication")%></td>
			<%
				else
			%>
			<td width="200">&nbsp;</td>
			<%
				end if
			
			%>
		</tr>
		<%
		Case 2
		%>
		<tr height="<%=cell_height%>">
			<td width="150">&nbsp;</td>
			<td width="150"><%=rs("BName")%></td>
			<%
				if request("Adv") = "" then
			%>
			<td width="150"><%=rs("TClass")%></td>
			<%
				else
			%>
			<td width="150">&nbsp;</td>
			<%
				end if
			%>
			<td width="50">&nbsp;</td>
			<td width="200">&nbsp;</td>
		</tr>
		<%
		Case 3
		%>
		<tr height="<%=cell_height%>">
			<td width="150">&nbsp;</td>
			<td width="150"><%=rs("BName")%></td>
			<td width="150">&nbsp;</td>
			<td width="50">&nbsp;</td>
			<%
				if request("Adv") = "" then
			%>
			<td width="200"><%=rs("Indication")%></td>
			<%
				else
			%>
			<td width="200">&nbsp;</td>
			<%
				end if
			
			%>
		</tr>
		<%
		Case Else
	End select
		total_s1 = total_s1 + 1
	Loop
	End if
	rs.close
	set rs = nothing
	%>
</table>
<br clear=all style='page-break-after:always'>
<input type="button" onClick="MM_goToURL('parent','exam04.asp');return document.MM_returnValue" value="More Practice">
<input type="button" onClick="MM_goToURL('parent','exam04.asp?Adv=1');return document.MM_returnValue" value="Advance Practice">
<input type="button" onClick="MM_goToURL('parent','testdisp.asp');return document.MM_returnValue" value="Back Main Page">
<input tYPE="button" value="print" onClick="window.print()">
</body>
</html>

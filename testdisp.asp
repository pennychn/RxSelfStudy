<% 
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache" 
' 一進測驗平台，要請先選題數
test_type = trim(Request("t_type"))
if test_type = "" then
	test_type = 0
End if
%>
<html>
<head>
<title>RX On-line Learing</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.p15 {  font-size: 15px; line-height: 20px; border-color: black black #000000; border-style: solid; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; font-family: "Arial"}
.p13 {  font-size: 13px; line-height: 20px; border-color: black black #CCCCCC; border-style: solid; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px; height: 0px; padding-top: 6px; padding-bottom: 6px; font-family: "Arial"}
.p1 {  font-size: 10px; line-height: 10px; border-color: black black #666666; border-style: dotted; border-top-width: 0px; border-right-width: 0px; border-bottom-width: 1px; border-left-width: 0px}
-->
</style>

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
<% 
'依不同按鈕(test_type)給予不同題型之題目數
'{
Select Case test_type
	Case 1
		qt1 = 20
	Case 2
		qt2 = 10
	Case 3
		qt3 = 10 
	Case 4
	Case ELSE
End Select
'}
%>
<form action="testans.asp" method="post" name="testdisp">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr> <td> 
<div align="center">
<table width="760" border="0" cellspacing="0" cellpadding="0">
	<tr> <td width="104" valign="top"> </td>
	<td valign="top"> 
	<table width="600" border="0" cellspacing="0" cellpadding="0">
	<tr> 
<% 
if test_type = "" OR test_type=0 then
%>
	<br>
	<table width="500" cellspacing="0" cellpadding="0" border="1">
	<% '{ %>
	<tr>
	<td width="70">Categories</td>
	<td width="200">Description</td>
	</tr>
	<tr>
	<td width="100" align="center"><input type="button" onClick="MM_goToURL('parent','exam01.asp');return document.MM_returnValue" value="Cate-1"></td>
	<td width="200">Generic and Brand Names </td>
	</tr>
	<tr>
	<td width="100" align="center"><input type="button" onClick="MM_goToURL('parent','exam02.asp');return document.MM_returnValue" value="Cate-2"></td>
	<td width="200">Therapeutic Class and Indication</td>
	</tr>
	<tr>
	<td width="100" align="center"><input type="button" onClick="MM_goToURL('parent','exam03.asp');return document.MM_returnValue" value="Cate-3"></td>
	<td width="200">Schedule </td>
	</tr>
	<tr>
	<td width="100" align="center"><input type="button" onClick="MM_goToURL('parent','exam04.asp');return document.MM_returnValue" value="Cate-4"></td>
	<td width="200">Review</td>
	</tr>
	<% '} %>
	</table>
<% 
Response.End()
End if 
%>

<%
	'Response.write(hint)
	'bsec=time()
	'bsec=test_sec+60
%>
<!--
<div align=left>
<iframe src="timesec.asp?test_type=<%=test_type%>"
align=left frameborder=0 height=30  width="140" marginheight=0 marginwidth=0 name=framename 
scrolling=no>
</iframe>
</div>
-->
</tr>
<tr>
<td> 
<table width="610" border="0">
<% '{ %>
<tr> 
<td width="31">&nbsp;</td>
<td width="333">&nbsp; </td>
</tr>
<tr><td>&nbsp;</td><td>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%
'---------------- 試題顯示主程式：Cate-1 ----------------
total_s1 = 0 '填充
'儲存已取出題目的容器
redim ary_qt1(qt1) 
For p = 0 to qt1 - 1
	ary_qt1(p) = 0
Next 

Const adCmdText = &H0001
strsql = "SELECT * FROM qryExamType1" 
' strsql = "SELECT top " & qt1 & " Rank, GName, BName FROM tblRxData ORDER BY RND([RID])" 
set rs=server.CreateObject("ADODB.Recordset")
rs.open strsql , conn , 3 , , adCmdText
if not rs.eof then
%>
<tr>
<td colspan="2" class="p15"><b>Please write its generic name or brand name.</b></td>
</tr>
<%
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
%>
<tr>
<td class="p13" colspan="2">
<% 
'Response.write(rs("Rank") & " ")
Response.write(i & ".")
if (i mod 2 = 1) then
%>
	<font color="#0000FF">
	<%=rs("GName")%> &nbsp; → &nbsp;<input type="text" name=myans<%=i%> value=""%>
  </font>
	<% 
	' Response.write (rs("BName"))
	%>
	<input type="hidden" name=qans<%=i%> value="<%=rs("BName")%>">
<%
else
%>
	<font color="#0000FF">
	<input type="text" name=myans<%=i%> value=""> &nbsp; ← &nbsp; <%=rs("BName")%>
	</font>
	<input type="hidden" name=qans<%=i%> value="<%=rs("GName")%>">
<%
end if
%> 
</td></tr>
<%
	total_s1 = total_s1 + 1
Loop
end if ' if not rs.eof then
%>                      
<% '} %>
</table>
</td>
</tr>
<%
rs.close
'---------------- End of 試題顯示主程式：Cate-1 ----------------
%>
<tr><td>&nbsp;</td></tr>
</table>
</td>
</tr><input type="hidden" name=total value=<%=total%>>

<tr> 
<td width="31">&nbsp;</td>
<td colspan="2" class="p1">&nbsp;</td>
</tr>
<tr> 
<td width="20%"></td><input type="hidden" name=recno value=<%=i%>>
<td width="80%" class="p13"><div align=left><input type=submit name="action" value="計分" class=c12border></div>

</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
<input type="hidden" name=test_type value=<%=test_type%>>
</form>
</body>
</html>

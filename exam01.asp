<% 
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache" 
%>
<html>
<head>
<title>�Ī��ӫ~�W/�ǦW����</title>
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

<form action="exam01_.asp" method="post" name="exam01">
<table width="610" border="0">
<% '{ %>
<tr> 
<td width="31">&nbsp;</td>
<td width="333">&nbsp; </td>
</tr>
<tr><td>&nbsp;</td><td>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%
qt1 = 20
total_s1 = 0 '��R
'�x�s�w���X�D�ت��e��
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
	'�üƲ��͡A�è������ƥX�{����
	rndCheck = 0
	Do While rndCheck = 0
		Randomize Timer
		rndNumber = Int(rs.RecordCount * Rnd)
		match_count = 0 '�U�C�d�ߦ��t�X�쪺�`��
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
	'End of �üƲ���
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
	<%=rs("GName")%> &nbsp; �� &nbsp;<input type="text" name=myans<%=i%> value=""%>
  </font>
	<% 
	' Response.write (rs("BName"))
	%>
	<input type="hidden" name=qans<%=i%> value="<%=rs("BName")%>">
<%
else
%>
	<font color="#0000FF">
	<input type="text" name=myans<%=i%> value=""> &nbsp; �� &nbsp; <%=rs("BName")%>
	</font>
	<input type="hidden" name=qans<%=i%> value="<%=rs("GName")%>">
<%
end if
%> 
</td>
</tr>
<%
	total_s1 = total_s1 + 1
Loop
		else 
	Response.Write ("Database is empty!!")
end if ' if not rs.eof then
rs.Close
%>                      
<% '} %>
<tr><td align="right">
<input type=submit name="action" value="Check answer">
<input type="button" onClick="MM_goToURL('parent','testdisp.asp');return document.MM_returnValue" value="Back Main Page">
</td></tr>
<input type="hidden" name=total value="<%=total%>">
<input type="hidden" name=recno value="<%=i%>">
</table>
</td>
</tr>
</table>
</form>
</body>
</html>

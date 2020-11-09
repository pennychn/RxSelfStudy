<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>藥理治療類別測驗</title>
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>
<!--#include file=connect.asp -->

<body>
<h1 align="left">管制藥分類測驗</h1>
<form action="exam03_.asp" method=post name="exam03">
<table width="700" border="1">
  <% 
		'Connect to Database
		Const adCmdText = &H0001
		strsql = "SELECT GName, BName, Schedule FROM tblRxData where Schedule <> '0'" 
		set rs=server.CreateObject("ADODB.Recordset")
		rs.open strsql , conn , 3 , , adCmdText
		' =============================================================
		' 顯示所有列管的藥品
		' =============================================================
  	totalQuestion = rs.RecordCount
		qcnt = 0 
		' Declare an array to store questions with size "totalQuestion"(user's assignment)
		redim questAry(totalQuestion) 
		For p = 0 to totalQuestion - 1
			questAry(p) = -1
		Next 
		' =============================================================
		if not rs.eof then
			Do while (not rs.eof and qcnt < totalQuestion)
				'亂數產生題目，並取不重複出現的值
				'rndCheck = 0
				'Do While rndCheck = 0
				'	Randomize Timer
				'	rndNumber = Int(rs.RecordCount * Rnd) + 1
				'	match_count = 0 '下列查詢有配合到的總數
				'	For p = 0 to qcnt
				'		if questAry(p) = rndNumber then		
				'			match_count = match_count + 1
				'		End if
				'	Next
				'	if match_count = 0 then
				'		questAry(qcnt) = rndNumber
				'		rndCheck = 1
				'	End if
				'Loop 
				' Move to assigned record
				'rs.MoveFirst
				'rs.Move rndNumber
				' 顯示選項
				%>
				<tr>
					<td>
				  <% ' Response.Write ("The answer is " & rs(2) & " ") %>
					<b><% Response.Write (qcnt + 1 & ". " & rs(0) & "(" & rs(1) & ") is Schedule _____?" ) %></b> <br>
					&nbsp;a)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=1>C I<br>
					&nbsp;b)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=2>C II<br>
					&nbsp;c)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=3>C III<br>
					&nbsp;d)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=4>C IV<br>
					&nbsp;e)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=5>C V<br>
					<!-- &nbsp;f)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=0>None<br> -->
					<input type="hidden" name=qans<%=qcnt+1%> value="<%=rs(2)%>">
					</td>
				</tr>
				<%
				qcnt = qcnt + 1
				rs.MoveNext
			Loop
		End if
		rs.close
		set rs = nothing 
  %>
</table>
<input type="hidden" name="recno" value=<%=totalQuestion%>>
<div align=left><input type=submit name="action" value="Check answer" class=c12border>
<input type="button" onClick="MM_goToURL('parent','testdisp.asp');return document.MM_returnValue" value="Back Main Page">
</div>

</form>
</body>
</html>

<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<!--#include file="connect.asp"-->
<%

cnt=cint(trim(request("recno"))) '�`���D��
'total_s=request("total") 
'Response.Write("cnt :" & cnt &" total_s :" & total_s & "<br>")
test_type=trim(request("test_type")) '�Ҹ��D��

idx=0
total=0
rtotal=0 '������� 
ttotal =0 '�`��
'ref_num = 1 '�ѦҫY�� 

'�U�������D��
rqt1 = 0 


errmsg=""
do while(idx<cnt)
	idx=idx+1
	qans = trim(request("qans"&cstr(idx)))
	myans= trim(request("myans"&cstr(idx)))
	if (myans <> "") then
		myans = ucase(myans)
		qans = ucase(qans)
	end if
	Response.write (idx & ". ")
	if (myans = qans) then
	%>
	Correct <br>
  <%
	else
	%>
	Not Correct! The answer is <u><b><%=qans%></b></u><br>
  <%
	end if
loop
%>
<center>
<input type="button" onClick="MM_goToURL('parent','testdisp.asp?t_type=1');return document.MM_returnValue" value="�A�դ@��">
<input type="button" onClick="MM_goToURL('parent','testdisp.asp');return document.MM_returnValue" value="�^�D����">
</center>

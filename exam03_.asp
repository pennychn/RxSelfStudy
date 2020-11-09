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

cnt=cint(trim(request("recno"))) '總問題數
idx=0

do while(idx<cnt)
	idx=idx+1
	myans = trim(request("qtid"&cstr(idx)))
	qans  = trim(request("qans"&cstr(idx)))
	'Response.write (qans & "<--->" & myans)
	Response.write (idx & ". ")
	if (myans = qans) then
		Response.Write("Correct! <br>")
	else
		Select Case qans
			case 0
				qans_desc = "None"
			case 1
				qans_desc = "C I"
			case 2
				qans_desc = "C II"
			case 3
				qans_desc = "C III"
			case 4
				qans_desc = "C IV"
			case 5
				qans_desc = "C V"
			case else
		end select
		Response.Write("Not Correct! The answer is <u><b>" & qans_desc & " </b></u><br>")
	end if
loop
%>
<center>
<input type="button" onClick="MM_goToURL('parent','exam03.asp');return document.MM_returnValue" value="Re-try">
<input type="button" onClick="MM_goToURL('parent','testdisp.asp');return document.MM_returnValue" value="Back Main Page">
</center>

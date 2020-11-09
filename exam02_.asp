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

dim match_count 
cnt=cint(trim(request("recno"))) '總問題數
idx=0

do while(idx<cnt)
	idx=idx+1
	myans = trim(request("qtid"&cstr(idx)))
	qans  = trim(request("qans"&cstr(idx)))
		
	'Response.write (qans & "<--->" & myans)
	Response.write (idx & ". Part 1: ")
	if (myans = qans) then
		Response.Write("Correct! <br>")
	else
		Response.Write("Not Correct! The answer is <u><b>Option " & qans & " </b></u><br>")
	end if

	chkgrp = request("chkgrp"&cstr(idx)) 
	checkAry = split(chkgrp, ",")
	'for each x in checkAry
		'Response.Write( x & "<br>")
	'next
	chkgrp_ans = request("chkgrp_ans"&cstr(idx))
	chkgrp_ans_num = request("GrpId"&cstr(idx)&"_AnsNum")
	if (chkgrp_ans_num = 1) then
	'{
		miss_match_count = 0
		for each x in checkAry
			if (CINT(x) <> CInt(chkgrp_ans)) then 
				miss_match_count = miss_match_count + 1
			end if
		next
		'Response.Write("*")
		if (miss_match_count <> 0 OR UBound(checkAry) = -1) then
			Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;Part 2: <font color=red>Not Correct!</font> The answer is <b>" & request("chkgrp_ansDesc"&cstr(idx)&"1") & "</b><br>")
		else
			Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;Part 2: Correct! <br>")
		end if
	'}
	elseif (chkgrp_ans_num = 2) then
	'{
		miss_match_count = 0
		for each x in checkAry
			if (CInt(x) > 2) then 
				'Response.Write(x & " ")
				miss_match_count = miss_match_count + 1
			end if
		next
		'Response.Write("**")
		'Response.Write(chkgrp_ans)
		if (miss_match_count <> 0 OR UBound(checkAry) = -1) then
			Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;Part 2: <font color=red> Not Correct! </font> The answers are <b>" & request("chkgrp_ansDesc"&cstr(idx)&"1") & " and " & request("chkgrp_ansDesc"&cstr(idx)&"2") & "</b><br>")
		else
			Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;Part 2: <b>Correct!</b> <br>")
		end if
	'}
	else
	'{
	' skip
	'}
	end if 
	Response.Write("<br>")
loop
%>
<center>
<input type="button" onClick="MM_goToURL('parent','exam02.asp');return document.MM_returnValue" value="Re-try">
<input type="button" onClick="MM_goToURL('parent','testdisp.asp');return document.MM_returnValue" value="Back Main Page">
</center>

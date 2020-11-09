<% Response.Expires = 0 %>
<% 
test_type = trim(Request("test_type"))
'依不同題數，給予不同參數
Select Case test_type
	Case 1
		%>
		<script language="javascript" src="../liveclock_rm1.js"></script>
		<%
	Case 2
		%>
		<script language="javascript" src="../liveclock_rm2.js"></script>
		<%
	Case 3
		%>
		<script language="javascript" src="../liveclock_rm3.js"></script>
		<%
End Select
%>
<HTML>
<BODY topmargin=0 leftmargin=0 bgcolor=#eeeeee onload="javascript:if(add_sec()) {javascript:top.location.href='self.asp';}">
</BODY>
</HTML>

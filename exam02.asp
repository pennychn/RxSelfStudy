<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>Therapeutic Class and Indication</title>
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<%
	Function generatePassword(passwordLength, sDefaultChars)
		'Declare variables
		Dim iCounter
		Dim sMyPassword
		Dim iPickedChar
		Dim iPasswordLength
		Dim LPart
		Dim RPart
		'Initialize variables
		iPasswordLength=passwordLength
		Randomize'initialize the random number generator
		'Loop for the number of characters password is to have
		For iCounter = 1 To iPasswordLength
			'Next pick a number from 1 to length of character set
			iPickedChar = Int((Len(sDefaultChars) * Rnd) + 1)
			sMyPassword = sMyPassword & Mid(sDefaultChars,iPickedChar,1)
		'(left(first_day,4),mid(first_day,5,2),right(first_day,2))
			LPart = Left(sDefaultChars, iPickedChar-1)
			RPart = Right(sDefaultChars, Len(sDefaultChars) - iPickedChar)
			sDefaultChars = LPart & RPart
		Next
		generatePassword = sMyPassword
	End Function
%>
</head>
<!--#include file=connect.asp -->

<body>
<h1 align="left">Therapeutic Class and Indication</h1>
<form action="exam02_.asp" method=post name="exam02">
<table width="900" border="1">
  <caption>&nbsp;
  </caption>
  <tr>
    <th scope="col">Please choose correct "therapeutic class."</th>
  </tr>
  <% 
  	const totalQuestion = 15 
		redim aryMainQuest(totalQuestion) ' store randomly generated Generic Name or Brand Name
		
		dim aryLevelOneOpts(3) ' Options for "Therapeutic Class"
		dim aryLevelTwoOpts(6) ' Options for "Indication"
		dim ansLevelTwo
		dim ansLevelTwoNum
		' =======================================================
		' Inital value assignment
		' =======================================================
		for p = 0 to totalQuestion-1
			aryMainQuest(p) = "&nbsp;"
		Next 
		For p = 0 to UBound(aryLevelOneOpts) - 1
			aryLevelOneOpts(p) = "&nbsp;"
		Next 
		For p = 0 to UBound(aryLevelTwoOpts) - 1
			aryLevelTwoOpts(p) = "&nbsp;"
		Next 

		qcnt = 0 
		' Declare an array to store questions with size "totalQuestion"(user's assignment)
		redim questAry(totalQuestion) 
		For p = 0 to totalQuestion - 1
			questAry(p) = -1 
		Next 
		dim questAry2(2)
		' =======================================================
		' Connect to Database
		' =======================================================
		Const adCmdText = &H0001
		strsql = "SELECT b.RID, a.Title, b.TClass, b.IName FROM qryUnionBGNames as a, qryIndicationByRank as b where a.Rank = b.Rank" 
		set rs = server.CreateObject("ADODB.Recordset")
		rs.open strsql , conn , 3 , , adCmdText
		if not rs.eof then
		'{
			'Response.Write("Total RecordCnt = " & rs.RecordCount & "<br>")
			Do while (not rs.eof and qcnt < totalQuestion)
				'亂數產生題目，並取不重複出現的值
				rndCheck = 0
				Do While rndCheck = 0
					Randomize Timer
					rndNumber = Int(rs.RecordCount * Rnd) ' Never modify this value
					match_count = 0 '下列查詢有配合到的總數
					' Move to assigned record
					rs.MoveFirst
					rs.Move rndNumber
					'Response.Write("RandNum = " & rndNumber)
					For p = 0 to qcnt
						if (strcomp(aryMainQuest(p), rs(1)) = 0) then		
							match_count = match_count + 1
						End if
					Next
					if (match_count = 0) then
						aryMainQuest(qcnt) = rs(1)
						rndCheck = 1
					End if
				Loop 
				' =========================================================
				' Generate "Therapeutic Class" options (answer is included)
				' #Options = 3
				' =========================================================
				'{
				subCurCnt = 0
				aryLevelOneOpts(2) = rs(2) ' Assign answer in the last options
				strsql2 = "SELECT a.TName FROM tblTherapeuticClass as a, tblDrugTherapeuticClass as b WHERE a.TherapeuticClassID = b.TherapeuticClassID and b.RID <> 27 and a.TName <> '" & rs(2) & "'"
				set rsx=server.CreateObject("ADODB.Recordset")
				rsx.open strsql2 , conn , 3 , , adCmdText
				if not rsx.eof then
					'Response.Write("Total RecordCnt = " & rsx.RecordCount & "<br>")
					Do while (not rsx.eof and subCurCnt < 2)
						'亂數產生題目，並取不重複出現的值
						rndCheck = 0
						Do While rndCheck = 0
							Randomize Timer
							rndNumber2 = Int(rsx.RecordCount * Rnd)
							rsx.MoveFirst
							rsx.Move rndNumber2
							match_count = 0 '下列查詢有配合到的總數
							'Response.Write("RandNum = " & rndNumber2)
							For p = 0 to subCurCnt
								if (strcomp(aryLevelOneOpts(p), rsx(0)) = 0) then		
									match_count = match_count + 1
								End if
							Next
							if (match_count = 0) then
								'Response.Write("  ==> add!!<br>")
								aryLevelOneOpts(subCurCnt) = rsx(0)
								rndCheck = 1
							End if
						Loop 
						subCurCnt = subCurCnt + 1
					Loop
				End if
				rsx.close
				set rsx = nothing 
				'}
				' ==================================================================================
				' Generate "Indication" options (answer is included)
				' #Options = 5(1 for hidden)
				' More than two answers is allowed!! 
				' ==================================================================================
				' {
				' 1. Fill the answer firstly
				icnt = 0
				ansLevelTwo = ""
				ansLevelTwoNum = 0
				strsql3 = "SELECT a.IName FROM tblIndication as a, tblDrugIndications as b WHERE a.IndicationID = b.IndicationID and b.RID = " & rs(0)
				set rsn=server.CreateObject("ADODB.Recordset")
				rsn.open strsql3 , conn , 3 , , adCmdText
				if not rsn.eof then
					Do while (not rsn.eof)
						aryLevelTwoOpts(icnt) = rsn(0)
						'aryLevelTwoOpts(icnt) = rsn(0)&"*" 'debug use
						ansLevelTwo = ansLevelTwo & (icnt+1)
						icnt = icnt + 1
						rsn.MoveNext
					LOOP
				else
					Response.Write("Attention: RID = " & rs(0) & ", not in tblIndication! <br>")
				end if
				rsn.close
				set rsn = nothing
				ansLevelTwoNum = icnt
				' 2. Fill the others options
				strsql3 = "SELECT IName FROM tblIndication WHERE IndicationID <> 18"
				set rsn=server.CreateObject("ADODB.Recordset")
				rsn.open strsql3 , conn , 3 , , adCmdText
				if not rsn.eof then
					Do while (not rsn.eof and icnt < UBound(aryLevelTwoOpts)-1)
						'亂數產生題目，並取不重複出現的值
						rndCheck = 0
						Do While rndCheck = 0
							Randomize Timer
							rndNumber3 = Int(rsn.RecordCount * Rnd)
							rsn.MoveFirst
							rsn.Move rndNumber3
							match_count = 0 
							For p = 0 to icnt
								if (strcomp(aryLevelTwoOpts(p), rsn(0)) = 0) then		
									match_count = match_count + 1
								End if
							Next
							if (match_count = 0) then
								aryLevelTwoOpts(icnt) = rsn(0)
								rndCheck = 1
							End if
						Loop 
						icnt = icnt + 1
					Loop
				End if
				rsn.close
				set rsn = nothing 
				aryLevelTwoOpts(icnt) = "None above"
				' }
				' ==================================================================================
				' Question Table Generation
				' ==================================================================================
				' 挑選答案排列方式
				' {
				Randomize Timer
				randNum = (CInt(5000 * Rnd) + 1) ' range from 1~400
				mode = (randNum mod 3)
				' Test: 檢查是否有出現重複的選項
				if (strcomp(aryLevelOneOpts(2), aryLevelOneOpts(0)) = 0 or strcomp(aryLevelOneOpts(2), aryLevelOneOpts(1)) = 0) then
					Response.Write(aryLevelOneOpts(2) & " --Some Bad!!")
				end if
				if (strcomp(aryLevelOneOpts(0), aryLevelOneOpts(1)) = 0) then
					Response.Write(aryLevelOneOpts(0) & " " & aryLevelOneOpts(1)& " --Some Bad2!!<br>")
				end if
				' Six Question Format
				' 1. GName w/ answer in A, B or C
				' 2. BName w/ answer in A, B or C
				%>
				<tr>
					<td>
					<b><% Response.Write (qcnt + 1 & ". ") %><%=aryMainQuest(qcnt)%></b> <br>
				<%
				select case mode
					Case 0 '@A1
						%>
							&nbsp;a)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=1><%=rs(2)%><br>
							&nbsp;b)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=2><%=aryLevelOneOpts(0)%><br>
							&nbsp;c)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=3><%=aryLevelOneOpts(1)%><br>
							<input type="hidden" name=qans<%=qcnt+1%> value="1"%>
						<%
					Case 1 '@B
						%>
							&nbsp;a)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=1><%=aryLevelOneOpts(0)%><br>
							&nbsp;b)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=2><%=rs(2)%><br>
							&nbsp;c)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=3><%=aryLevelOneOpts(1)%><br>
							<input type="hidden" name=qans<%=qcnt+1%> value="2"%>
						<%
					Case 2 '@C
						%>
							&nbsp;a)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=1><%=aryLevelOneOpts(0)%><br>
							&nbsp;b)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=2><%=aryLevelOneOpts(1)%><br>
							&nbsp;c)&nbsp;<input type=radio name=qtid<%=qcnt+1%> value=3><%=rs(2)%><br>
							<input type="hidden" name=qans<%=qcnt+1%> value="3"%>
						<%
					Case Else
				End select
				%>
				<br><b> What is this for?  </b> <br>
				<%
				' Show "What is this for?"
				if (ansLevelTwoNum = 1) then
				'{
					optOrder = generatePassword(5, "01234")
					for p = 1 to len(optOrder)-1 ' Print 4 value in five
						select case mid(optOrder, p, 1)
							case 0
							%>
								&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=1>&nbsp;<%=aryLevelTwoOpts(0)%>
							<%
							case 1
							%>
								&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=2>&nbsp;<%=aryLevelTwoOpts(1)%>
							<%
							case 2
							%>
								&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=3>&nbsp;<%=aryLevelTwoOpts(2)%>
							<%
							case 3
							%>
								&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=4>&nbsp;<%=aryLevelTwoOpts(3)%>
							<%
							case 4
							%>
								&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=5>&nbsp;<%=aryLevelTwoOpts(4)%>
							<%
							case else
						end select
					Next
					if strcomp(right(optOrder,1), "0") = 0 then
						ansLevelTwo = "6"
						aryLevelTwoOpts(0) = "None Above"
						'Response.Write("<b>Answer is none above</b>")
					end if 
				'}
				else
				'{
					optOrder = generatePassword(4, "0123")
					for p = 1 to len(optOrder) ' Print 4 value in five
						select case mid(optOrder, p, 1)
							case 0
							%>
								&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=1>&nbsp;<%=aryLevelTwoOpts(0)%>
							<%
							case 1
							%>
								&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=2>&nbsp;<%=aryLevelTwoOpts(1)%>
							<%
							case 2
							%>
								&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=3>&nbsp;<%=aryLevelTwoOpts(2)%>
							<%
							case 3
							%>
								&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=4>&nbsp;<%=aryLevelTwoOpts(3)%>
							<%
							case else
						end select
					Next
				'}
				end if
				%>
				&nbsp;<input type="checkbox" name=chkgrp<%=qcnt+1%> value=6>&nbsp;<%=aryLevelTwoOpts(5)%>
							<input type="hidden" name=chkgrp_ans<%=qcnt+1%> value="<%=ansLevelTwo%>"%>
							<input type="hidden" name="GrpId<%=qcnt+1%>_AnsNum" value="<%=ansLevelTwoNum%>"%>
							<input type="hidden" name="chkgrp_ansDesc<%=qcnt+1%>1" value="<%=aryLevelTwoOpts(0)%>"%>
							<input type="hidden" name="chkgrp_ansDesc<%=qcnt+1%>2" value="<%=aryLevelTwoOpts(1)%>"%>
					</td>
				</tr>
				<%
				'}
				qcnt = qcnt + 1
				' Clear all data in aryLevelOneOpts and aryLevelTwoOpts
				For p = 0 to UBound(aryLevelOneOpts)-1
					aryLevelOneOpts(p) = "&nbsp;"
				Next 
				For p = 0 to UBound(aryLevelTwoOpts)-1
					aryLevelTwoOpts(p) = "&nbsp;"
				Next 
			Loop
		'}
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

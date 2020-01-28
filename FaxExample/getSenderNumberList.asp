<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 팩스 발신번호 목록을 확인합니다.
	' - https://docs.popbill.com/fax/asp/api#GetSenderNumberList
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		 
	
	On Error Resume Next

	Set Presponse = m_FaxService.GetSenderNumberList(testCorpNum)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>팩스 발신번호 목록 확인</legend>
				<%
					For i=0 To Presponse.length -1
				%>
				<fieldset class="fieldset2">
				<ul>
					<li>발신번호 (number) : <%=Presponse.Get(i).number%> </li>
					<li>대표번호 지정여부 (representYN) : <%=Presponse.Get(i).representYN%> </li>
					<li>등록상태 (state) : <%=Presponse.Get(i).state%> </li>
				</ul>
				</fieldset>
				<%
					Next
				%>

		 </div>
	</body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 팩스 전송단가를 확인합니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 
	
	On Error Resume Next

	unitCost = m_FaxService.GetUnitCost(testCorpNum)
	
	If Err.Number <> 0 then
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
				<legend>팩스 전송 단가 확인 </legend>
				<ul>
					<% If code = 0 Then %>
						<li>전송 단가 : <%=unitCost%> </li>
					<% Else %>
						<li> Response.code : <%=code%></li>
						<li> Response.message : <%=message%></li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
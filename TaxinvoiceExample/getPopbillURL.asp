<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	 '회원 사업자번호, "-" 제외
	userID = "testkorea"  ' 회원 아이디
	TOGO = "LOGIN"   'LOGIN = 팝빌로그인 / CHRG : 포인트충전 팝업 / CERT : 공인인증서 등록 팝업 / SEAL : 인감 및 첨부문서 등록 URL

	On Error Resume Next

	url = m_TaxinvoiceService.GetPopbillURL(testCorpNum, userID, TOGO)

	If Err.Number <> 0 then
		code = Err.Number
		message =  Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>팝빌 SSO URL 요청</legend>
				<% If code = 0 Then %>
					<ul>
						<li>URL : <%=CStr(url)%> </li>
					</ul>
				<%	Else  %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>
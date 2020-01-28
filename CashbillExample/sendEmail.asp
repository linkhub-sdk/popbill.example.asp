<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 발행 안내메일을 재전송합니다.
	' - https://docs.popbill.com/cashbill/asp/api#SendEmail
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	'팝빌 회원 아이디
	userID = "testkorea"		 

	'문서번호	
	mgtKey = "20190103-001"		 

	'수신 메일주소
	'팝빌 개발환경에서 테스트하는 경우에도 안내 메일이 전송되므로,
	'실제 거래처의 메일주소가 기재되지 않도록 주의
	receiver = "test@test.com"		

	On Error Resume Next
		
	Set Presponse = m_CashbillService.SendEmail(testCorpNum, mgtKey, receiver, userID)

	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = Presponse.code
		message = Presponse.message
	End If

	On Error GoTo 0 

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>알림메일 재전송</legend>
				<ul>
					<li>Response.code : <%=code%></li>
					<li>Response.message : <%=message%></li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
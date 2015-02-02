<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"			'팝빌 회원 사업자번호, "-"제외 10자리
	userID = "testkorea"				'팝빌 회원 아이디
	itemCode = "121"					'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	mgtKey = "20150130-01"				'연동관리번호
	sender = "07075103710"				'발신자번호
	receiver = "010111222"			'수신자번호
	contents = "전자명세서 알림문자전송 테스트입니다."		'메시지내용

	On Error Resume Next

	Set Presponse = m_StatementService.SendSMS(testCorpNum, itemCode, mgtKey, sender, receiver, contents, userID)

	If Err.Number <> 0 Then
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
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
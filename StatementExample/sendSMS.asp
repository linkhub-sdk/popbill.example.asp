<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 알림문자를 전송합니다. (단문/SMS- 한글 최대 45자)
	' - 알림문자 전송시 포인트가 차감됩니다. (전송실패시 환불처리)
	' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [전송내역] 탭에서
	'   전송결과를 확인할 수 있습니다.
	'**************************************************************
	
	'팝빌 회원 사업자번호, "-"제외 10자리
	testCorpNum = "1234567890"
	
	'팝빌 회원 아이디
	userID = "testkorea"

	'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
	itemCode = "121"

	'문서번호
	mgtKey = "20190103-001"

	'발신번호
	sender = "07043042991"

	'수신번호
	receiver = "010111222"

	'메시지 내용, 90byte초과시 길이가 조정되어 전송됨
	contents = "전자명세서 알림문자전송 테스트입니다."

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
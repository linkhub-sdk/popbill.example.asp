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
	' - 전송내역 확인은 "팝빌 로그인" > [문자 팩스] > [문자] > [전송내역] 탭에서
	'   전송결과를 확인할 수 있습니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	' 팝빌회원 아이디
	testUserID = "testkorea"
	
	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType = "SELL"

	' 문서번호 
	MgtKey = "20190103-001"

	'발신번호
	Sender = "07043042991"

	'수신번호
	Receiver = "070111222"
	
	'메시지 내용, 90byte초과시 길이가 조정되어 전송됨
	Contents = "문자 테스트입니다 90Bytes를 초과한내용은 전송되지 않습니다" 
	
	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.SendSMS(testCorpNum, KeyType, MgtKey, Sender, Receiver, Contents, testUserID)
	
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
				<legend>알림문자 전송</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
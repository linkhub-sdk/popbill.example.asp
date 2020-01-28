<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 전자세금계산서를 팩스전송합니다.
	' - 팩스 전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
	' - https://docs.popbill.com/taxinvoice/asp/api#SendFAX
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외 10자리
	testCorpNum = "1234567890"

	' 팝빌회원 아이디
	testUserID = "testkorea"
	
	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"

	' 문서번호 
	MgtKey = "20190103-001"

	' 발신번호
	Sender = "07043042991"

	' 수신팩스번호
	Receiver = "070111222"

	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.SendFAX(testCorpNum, KeyType, MgtKey, Sender, Receiver, testUserID)
	
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
				<legend>팩스 재전송</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
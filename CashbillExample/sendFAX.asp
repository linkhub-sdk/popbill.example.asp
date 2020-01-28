<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 현금영수증을 팩스전송합니다.
	' - 팩스 전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
	' - https://docs.popbill.com/cashbill/asp/api#SendFAX
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	'팝빌 회원 아이디
	userID = "testkorea"		 

	'문서번호	
	mgtKey = "20190103-001"		 

	'발신번호
	sender = "07043042991"		 

	'수신팩스번호
	receiver = "070111222"		 

	On Error Resume Next 

	Set Presponse = m_CashbillService.SendFAX(testCorpNum, mgtKey, Sender, Receiver, UserID)

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
				<legend>현금영수증 팩스전송 </legend>
				<ul>
					<li>Response.code : <%=code%></li>
					<li>Response.message : <%=message%></li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	'팝빌 회원 아이디
	userID = "testkorea"

	'광고문자 전송여부
	adsYN = False

	'예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
	reserveDT = ""
	
	Set msgList = CreateObject("Scripting.Dictionary")

	'문자메시지 전송정보
	Set message = New Messages

	'발신번호
	message.sender = "07075103710"

	'발신자명
	message.senderName = "발신자명"

	'수신번호
	message.receiver = "000111222"

	'수신자명
	message.receivername = "수신자이름"

	'메시지내용, 90byte 기준으로 단/장문이 자동으로 인식되어 전송
	message.content = "단/장문 메시지 자동인식전송 테스트입니다. 전송하는 메시지의 길이가 90byte이상인 경우 장문(LMS)타입으로 메시지가 전송됩니다. 문자전송 테스트입니다."
	
	msgList.Add 0, message

	On Error Resume Next

	receiptNum = m_MessageService.SendXMS(testCorpNum, "", "", "", msgList, reserveDT, adsYN, userID)

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
				<legend>단/장문 자동인식전송 1건 전송 </legend>
				<% If code = 0 Then %>
					<ul>
						<li>ReceiptNum(접수번호) : <%=receiptNum%> </li>
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
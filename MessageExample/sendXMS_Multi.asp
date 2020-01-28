<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
    ' [대량전송] XMS(단문/장문 자동인식)를 전송합니다.
    ' - 메시지 내용의 길이(90byte)에 따라 SMS/LMS(단문/장문)를 자동인식하여 전송합니다.
    ' - 90byte 초과시 LMS(장문)으로 인식 합니다.
    ' - https://docs.popbill.com/message/asp/api#SendXMS
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	'팝빌 회원 아이디
	userID = "testkorea"

	'광고문자 전송여부
	adsYN = False

	'예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
	reserveDT = ""


	'문자전송정보 배열, 최대 1000건
	Set msgList = CreateObject("Scripting.Dictionary")
	
	For i = 0 To 49
		Set message = New Messages

		'발신번호
		message.sender = "07043042991"

		'발신자명
		message.senderName = "발신자명"

		'수신번호
		message.receiver = "000111222"

		'수신자명
		message.receivername = " 수신자이름"+CStr(i)

		'메시지내용, 90byte기준으로 단/장문 자동인식 전송
		message.content = "문자내용이 90byte 이하인경우 단문(sms)로 전송됩니다."

		msgList.Add i, message
	Next

	For i = 50 To 99
		Set message = New Messages

		'발신번호
		message.sender = "07043042991"

		'발신자명
		message.senderName = "발신자명"

		'수신번호
		message.receiver = "000111222"

		'수신자명
		message.receivername = " 수신자이름"+CStr(i)

		'메시지내용, 90byte기준으로 단/장문 자동인식 전송
		message.content = "단/장문 자동인식 메시지 테스트입니다. 문자내용의 길이가 90byte 이상인경우 장문(LMS)로 전송됩니다 단/장문 자동인식 메시지 테스트입니다."

		'메시지제목
		message.subject = "장문 제목입니다"

		msgList.Add i, message
	Next

	'전송요청번호 (팝빌 회원별 비중복 번호 할당)
	'영문,숫자,'-','_' 조합, 최대 36자
	requestNum = ""	

	On Error Resume Next


	receiptNum = m_MessageService.SendXMS(testCorpNum, "", "", "", msgList, reserveDT, adsYN, requestNum, userID)

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
				<legend>단/장문 자동인식 문자메시지 100건 전송 </legend>
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
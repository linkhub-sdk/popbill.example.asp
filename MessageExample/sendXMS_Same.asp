<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
    ' [동보전송] XMS(단문/장문 자동인식)를 전송합니다.
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

	'발신번호
	senderNum = "07043042991"

	'메시지 제목
	subject = "문자전송 제목"

	'메시지 내용, 90byte 기준으로 단/장문 메시지가 자동인식되어 전송
	content = "문자전송 내용, 90Byte초과시 LMS(장문) 메시지로 전송됨" 


	'수신정보 배열, 최대 1000건	
	Set msgList = CreateObject("Scripting.Dictionary")
	
	For i =0 To 49
		Set message = New Messages

		'수신번호
		message.receiver = "000111222"

		'수신자명
		message.receivername = " 수신자이름"+CStr(i)

		msgList.Add i, message
	Next

	For i =50 To 99
		Set message = New Messages

		'수신번호
		message.receiver = "000111222"

		'수신자명
		message.receivername = " 수신자이름"+CStr(i)

		msgList.Add i, message
	Next
	
	'전송요청번호 (팝빌 회원별 비중복 번호 할당)
	'영문,숫자,'-','_' 조합, 최대 36자
	requestNum = ""	

	On Error Resume Next

	receiptNum = m_MessageService.SendXMS(testCorpNum, senderNum, subject,content, msgList, reserveDT, adsYN, requestNum, userID)

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
				<legend>단/장문 자동인식 문자메시지 동보전송 </legend>
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
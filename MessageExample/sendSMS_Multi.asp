<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'팝빌 회원 사업자번호, "-" 제외
	userID = "testkorea"					'팝빌 회원 아이디
	adsYN = False							'광고문자 전송여부
'	reserveDT = "20150128200000"    '예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
	
	Set msgList = CreateObject("Scripting.Dictionary")

	For i=0 To 99
		Set message = New Messages
		message.sender = "07075103710"
		message.senderName = "발신자명"
		message.receiver = "000111222"
		message.receivername = " 수신자이름"+CStr(i)
		message.content = "This is Message 메시지 테스트중"
		msgList.Add i, message
	Next
	

	On Error Resume Next

	receiptNum = m_MessageService.SendSMS(testCorpNum, "", "", "", msgList, reserveDT, adsYN, userID)

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
				<legend>단문 문자메시지 100건 전송 </legend>
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
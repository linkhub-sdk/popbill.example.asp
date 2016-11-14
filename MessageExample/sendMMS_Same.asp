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
	
	'발신번호
	senderNum = "07043042991"        

	'메시지 제목
	subject = "동보전송 메시지 제목"

	'메시지 네ㅐ용 
	content = "동보전송 메시지 내용"
	
	Set msgList = CreateObject("Scripting.Dictionary")
	
	'문자전송정보 배열, 최대 1000건
	For i =0 To 99

		Set message = New Messages
		'수신번호
		message.receiver = "000111222"

		'수신자명
		message.receivername = " 수신자이름"+CStr(i)
		msgList.Add i, message
	Next
		
	'포토메시지 이미지파일, 300Kbyte JPEG 포맷 전송가능
	FilePaths = Array("C:\popbill.example.asp\test.jpg")

	On Error Resume Next

	receiptNum = m_MessageService.SendMMS(testCorpNum, senderNum, subject, content, msgList, FilePaths, reserveDT, adsYN, userID)

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
				<legend>MMS 문자메시지 전송</legend>
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
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
    ' [동보전송] LNS(장문)를 전송합니다.
    '  - 메시지 내용이 2,000Byte 초과시 메시지 내용은 자동으로 제거됩니다.
    '  - 단건/대량 전송에 대한 설명은 "[문자 API 연동매뉴얼] > 3.2.2 SendLMS(장문전송)"을 참조하시기 바랍니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		

	'팝빌 회원 아이디
	userID = "testkorea"					

	'발신번호
	senderNum = "07043042991"		

	'메시지 제목
	subject = "동보전송 메시지 제목"

	'메시지 내용, 최대 2000byte 초과시 길이가 조정되어 전송됨
	content = "동보전송 메시지 내용"

	'예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
	reserveDT = ""

	'광고문자 전송여부
	adsYN = False

	'수신정보배열, 최대 1000건
	Set msgList = CreateObject("Scripting.Dictionary")
	
	For i =0 To 15
		Set message = New Messages

		'수신번호
		message.receiver = "000111222"

		'수신자명
		message.receivername = " 수신자이름"+CStr(i)
	
		msgList.Add i, message
	Next
	
	On Error Resume Next

	'전송요청번호 (팝빌 회원별 비중복 번호 할당)
	'영문,숫자,'-','_' 조합, 최대 36자
	requestNum = ""	

	receiptNum = m_MessageService.SendLMS(testCorpNum, senderNum, subject, content, msgList, reserveDT, adsYN, requestNum, userID)

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
				<legend>장문 문자메시지 동보전송 </legend>
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
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
	testUserID = "testkorea"					

	'알림톡 템플릿 코드 - 템플릿 목록 조회 (ListATSTemplate API)의 반환항목 확인
	templateCode = "018020000001"

	'팝빌에 사전 등록된 발신번호
	senderNum = "07043042993"

	'알림톡 내용, 최대 400자
	content = "[테스트] 테스트 템플릿입니다."

	'대체문자 내용
	altContent = "대체문자 메시지 내용"

	'대체문자 전송유형 공백-미전송, A-대체문자내용 전송, C-알림톡내용 전송
	altSendType = "C"

	'예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
	reserveDT = ""

	Set receiverList = CreateObject("Scripting.Dictionary")

	'수신정보 배열, 최대 1000건
	For i =0 To 9
		Set rcvInfo = New KakaoReceiver

		'수신자번호
		rcvInfo.rcv = "01011222"+ CStr(i)			

		'수신자명
		rcvInfo.rcvnm = " 수신자이름"

		receiverList.Add i, rcvInfo
	Next 
	
	On Error Resume Next
	
	receiptNum = m_KakaoService.SendATS(testCorpNum, templateCode, senderNum, content, altContent, altSendType, reserveDT, receiverList, testUserID)

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
				<legend>알림톡 동일내용 대량전송</legend>
				<% If code = 0 Then %>
					<ul>
						<li>ReceiptNum(접수번호) : <%=receiptNum%> </li>
					</ul>
				<% Else %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<% End If %>
			</fieldset>
		 </div>
	</body>
</html>
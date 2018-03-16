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

	'팝빌에 등록된 플러스친구 아이디
	plusFriendID = "@팝빌"

	'팝빌에 사전 등록된 발신번호
	senderNum = "07043042993"

	'친구톡 내용, 최대 1000자
	content = "친구톡 메시지 내용입니다"

	'대체문자 내용
	altContent = "대체문자 메시지 내용"

	'대체문자 전송유형 공백-미전송, A-대체문자내용 전송, C-알림톡내용 전송
	altSendType = "C"

	'예약전송시간 yyyyMMddHHmmss, reserveDT값이 없는 경우 즉시전송
	reserveDT = ""

	'광고전송 여부 
	adsYN = False

	'이미지 파일 경로
	filePaths = Array("C:\popbill.example.asp\test03.jpg")

	'이미지 링크 URL
	imageURL = "http://popbill.com"

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


	'친구톡 버튼정보 구성
	Set btnList = CreateObject("Scripting.Dictionary")
	Set btnInfo = New KakaoButton
	btnInfo.n = "버튼이름"			
	btnInfo.t = "WL"		
	btnInfo.u1 = "http://www.popbil.com"
	btnInfo.u2 = "http://www.llinkhub.co.kr"
	btnList.Add 0, btnInfo

	Set btnInfo = New KakaoButton
	btnInfo.n = "메시지 전달"			
	btnInfo.t = "MD"		
	btnList.Add 1, btnInfo
	
	On Error Resume Next

	receiptNum = m_KakaoService.SendFMS(testCorpNum, plusFriendID, senderNum, content, _
		altContent, altSendType, reserveDT, adsYN, receiverList, btnList, filePaths, imageURL, testUserID)

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
				<legend>친구톡 동일내용 대량 전송</legend>
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
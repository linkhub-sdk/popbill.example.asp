<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
    ' [동보전송] 알림톡 전송을 요청합니다.
    ' - 사전에 승인된 템플릿의 내용과 알림톡 전송내용(content)이 다를 경우 전송실패 처리됩니다
	' - https://docs.popbill.com/kakao/asp/api#SendATS
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		

	'팝빌 회원 아이디
	testUserID = "testkorea"					

	'알림톡 템플릿 코드 - 템플릿 목록 조회 (ListATSTemplate API)의 반환항목 확인
	templateCode = "019020000163"

	'팝빌에 사전 등록된 발신번호
	senderNum = "01043245117"

	'알림톡 내용, 최대 1000자
	content = "[ 팝빌 ]" & vbCrLf
	content = content + "신청하신 #{템플릿코드}에 대한 심사가 완료되어 승인 처리되었습니다." & vbCrLf
	content = content + "해당 템플릿으로 전송 가능합니다." & vbCrLf & vbCrLf
	content = content + "문의사항 있으시면 파트너센터로 편하게 연락주시기 바랍니다. " & vbCrLf & vbCrLf
	content = content + "팝빌 파트너센터 : 1600-8536" & vbCrLf
	content = content + "support@linkhub.co.kr"

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

		'파트너 지정키, 수신자 구별용 메모, 미사용시 공백처리
		rcvInfo.interOPRefKey = "20211201-" +CStr(i)

		receiverList.Add i, rcvInfo
	Next 
	
	
	
	'전송요청번호 (팝빌 회원별 비중복 번호 할당)
	'영문,숫자,'-','_' 조합, 최대 36자
	requestNum = ""	

	' 알림톡 버튼정보를 템플릿 신청시 기재한 버튼정보와 동일하게 전송하는 경우 btnList를 선언만 하고 함수호출.
	Set btnList = CreateObject("Scripting.Dictionary")
	
	'알림톡 버튼 URL에 #{템플릿변수}를 기재한경우 템플릿변수 영역을 변경하여 버튼정보 구성
	'Set btnInfo = New KakaoButton
	'btnInfo.n = "템플릿 안내"			
	'btnInfo.t = "WL"		
	'btnInfo.u1 = "https://www.popbil.com"
	'btnInfo.u2 = "http://www.llinkhub.co.kr"
	'btnList.Add 0, btnInfo

	On Error Resume Next

	receiptNum = m_KakaoService.SendATS(testCorpNum, templateCode, senderNum, content, altContent, altSendType, reserveDT, receiverList, requestNum, testUserID, btnList)

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
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 전송요청번호(requestNum)를 할당한 알림톡/친구톡 전송내역 및 전송상태를 확인한다.
	' - https://docs.popbill.com/kakao/asp/api#GetMessagesRN
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	'팝빌 회원 아이디
	userID = "testkorea"

	'전송 요청시 할당한 전송요청번호(requestNum)
	requestNum = "20180928111311"
	
	On Error Resume Next

	Set result = m_KakaoService.GetMessagesRN(testCorpNum, requestNum, UserID)

	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>카카오톡 전송결과 확인</legend>
					<%
						If code = 0 Then
					%>
					<ul>
						<li>contentType (카카오톡 유형) : <%=result.contentType%></li>
						<li>templateCode (알림톡 템플릿 코드) : <%=result.templateCode%></li>
						<li>plusFriendID (플러스친구 아이디) : <%=result.plusFriendID%></li>
						<li>sendNum (발신번호) : <%=result.sendNum%></li>
						<li>altContent (대체문자 내용) : <%=result.altContent%></li>
						<li>altSendType (대체문자 유형) : <%=result.altSendType%></li>
						<li>reserveDT (예약일시) : <%=result.reserveDT%></li>
						<li>adsYN (광고전송 여부) : <%=result.adsYN%></li>
						<li>imageURL (친구톡 이미지 URL) : <%=result.imageURL%></li>
						<li>sendCnt (전송건수) : <%=result.sendCnt%></li>
						<li>successCnt (성공건수) : <%=result.successCnt%></li>
						<li>failCnt (실패건수) : <%=result.failCnt%></li>
						<li>altCnt (대체문자 건수) : <%=result.altCnt%></li>
						<li>cancelCnt (취소건수) : <%=result.cancelCnt%></li>
					</ul>
					<%
						For i=0 To Ubound(result.btns)-1
					%>
						<fieldset class="fieldset2">
							<legend>버튼정보 [<%=i+1%>]</legend>
							<ul>
								<li>n (버튼명) : <%=result.btns(i).n%> </li>
								<li>t (버튼유형) : <%=result.btns(i).t%> </li>
								<li>u1 (버튼링크1) : <%=result.btns(i).u1%> </li>
								<li>u2 (버튼링크2) : <%=result.btns(i).u2%> </li>
							</ul>
						</fieldset>						
					<%
						Next
					%>
					<fieldset class="fieldset2">
						<legend>전송결과 정보 목록</legend>
					<%
						For i=0 To UBound(result.msgs) -1
					%>
						<fieldset class="fieldset3">
							<legend>전송결과 정보 [<%=i+1%>]</legend>
							<ul>
								<li>state (전송상태 코드) : <%=result.msgs(i).state%> </li>
								<li>sendDT (전송일시) : <%=result.msgs(i).sendDT%> </li>
								<li>receiveNum (수신번호) : <%=result.msgs(i).receiveNum%> </li>
								<li>receiveName (수신자명) : <%=result.msgs(i).receiveName%> </li>
								<li>content (알림톡/친구톡 내용) : <%=result.msgs(i).content%> </li>
								<li>result (알림톡/친구톡 전송결과 코드) : <%=result.msgs(i).result%> </li>
								<li>resultDT (알림톡/친구톡 전송결과 수신일시) : <%=result.msgs(i).resultDT%> </li>
								<li>altContent (대체문자 내용) : <%=result.msgs(i).altContent%> </li>
								<li>altContentType (대체문자 전송유형) : <%=result.msgs(i).altContentType%> </li>
								<li>altSendDT (대체문자 전송일시) : <%=result.msgs(i).altSendDT%> </li>
								<li>altResult (대체문자 전송결과 코드) : <%=result.msgs(i).altResult%> </li>
								<li>altResultDT (대체문자 전송결과 수신일시) : <%=result.msgs(i).altResultDT%> </li>
								<li>receiptNum (접수번호) : <%=result.msgs(i).receiptNum%> </li>
								<li>requestNum (요청번호) : <%=result.msgs(i).requestNum%> </li>
								<li>interOPRefKey (파트너 지정키) : <%=result.msgs(i).interOPRefKey%> </li>
							</ul>
						</fieldset>
					<% 
						Next
					%>
						</fieldset>
					<%
						Else
					%>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>
			</fieldset>
		 </div>
	</body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 문자전송요청에 대한 전송결과를 확인합니다.
	' - https://docs.popbill.com/message/asp/api#GetMessages
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	'팝빌 회원 아이디
	userID = "testkorea"

	'문자 전송시 발급받은 접수번호(ReceiptNum)
	ReceiptNum = "017071711000000002"
	
	On Error Resume Next

	Set result = m_MessageService.GetMessages(testCorpNum, ReceiptNum, UserID)
	
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
				<legend>문자메시지 전송결과 확인</legend>
				<ul>
					<% If code = 0 Then
						For i=0 To result.Count-1 
					%>
						<fieldset class="fieldset2">
							<legend>문자메시지 전송결과 [<%=i+1%>]</legend>
							<ul>
								<li>state (전송상태 코드) : <%=result.Item(i).state%> </li>
								<li>result (전송결과 코드) : <%=result.Item(i).result%> </li>
								<li>subject (메시지 제목) : <%=result.Item(i).subject%> </li>
								<li>content (메시지 내용) : <%=result.Item(i).content%> </li>
								<li>type (메시지 유형) : <%=result.Item(i).msgType%> </li>
								<li>sendnum (발신번호) : <%=result.Item(i).sendnum%> </li>
								<li>senderName (발신자명) : <%=result.Item(i).senderName%> </li>
								<li>receiveNum (수신번호) : <%=result.Item(i).receiveNum%> </li>
								<li>receiveName (수신자명) : <%=result.Item(i).receiveName%> </li>
								<li>receiptDT (접수일시) : <%=result.Item(i).receiptDT%> </li>
								<li>sendDT (전송일시) : <%=result.Item(i).sendDT%> </li>
								<li>resultDT (전송결과 수신일시) : <%=result.Item(i).resultDT%> </li>
								<li>reserveDT (예약일시) : <%=result.Item(i).reserveDT%> </li>
								<li>tranNet (전송처리 이동통신사명) : <%=result.Item(i).tranNet%> </li>
								<li>receiptNum (접수번호) : <%=result.Item(i).receiptNum%> </li>
								<li>requestNum (요청번호) : <%=result.Item(i).requestNum%> </li>
							</ul>
						</fieldset>
					<% 
						Next
						Else
					%>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
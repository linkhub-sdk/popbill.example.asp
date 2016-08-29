<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"			'팝빌 회원 사업자번호, "-" 제외
	userID = "testkorea"				'팝빌 회원 아이디
	ReceiptNum = "016082909000000003"   '문자 전송시 발급받은 접수번호(ReceiptNum)
	
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
							<legend> 문자메시지 전송결과 [<%=i+1%>] </legend>
							<ul>
								<li>state : <%=result.Item(i).state%> </li>
								<li>sendResult : <%=result.Item(i).sendResult%> </li>
								<li>sendDT : <%=result.Item(i).sendDT%> </li>
								<li>resultDT : <%=result.Item(i).resultDT%> </li>
								<li>subject : <%=result.Item(i).subject%> </li>
								<li>content : <%=result.Item(i).content%> </li>
								<li>type : <%=result.Item(i).msgType%> </li>
								<li>sendnum: <%=result.Item(i).sendnum%> </li>
								<li>senderName: <%=result.Item(i).senderName%> </li>
								<li>receiveNum : <%=result.Item(i).receiveNum%> </li>
								<li>receiveName : <%=result.Item(i).receiveName%> </li>
								<li>reserveDT : <%=result.Item(i).reserveDT%> </li>
								<li>tranNet : <%=result.Item(i).tranNet%> </li>
								<li>receiptDT : <%=result.Item(i).receiptDT%> </li>
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
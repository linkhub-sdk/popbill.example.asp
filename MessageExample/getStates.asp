<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 문자 전송내역 요약정보를 확인합니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	'팝빌 회원 아이디
	userID = "testkorea"

	'문자 전송시 발급받은 접수번호 배열(ReceiptNumList)
	Dim ReceiptNumList(2)
	ReceiptNumList(0) = "018041717000000018"
	ReceiptNumList(1) = "018041717000000019"
	
	On Error Resume Next

	Set result = m_MessageService.GetStates(testCorpNum, ReceiptNumList, UserID)
	
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
				<legend>문자메시지 요약정보 확인</legend>
				<ul>
					<% If code = 0 Then
							For i=0 To result.Count-1 
					%>
						<fieldset class="fieldset2">
							<legend>문자메시지 전송결과 [<%=i+1%>]</legend>
							<ul>
								<li>rNum (접수번호) : <%=result.Item(i).rNum%> </li>
								<li>sn (일련번호) : <%=result.Item(i).sn%> </li>
								<li>stat (전송 상태코드) : <%=result.Item(i).stat%> </li>
								<li>rlt (전송 결과코드) : <%=result.Item(i).rlt%> </li>
								<li>sDT (전송일시) : <%=result.Item(i).sDT%> </li>
								<li>rDT (결과코드 수신일시) : <%=result.Item(i).rDT%> </li>
								<li>net (전송 이동통신사명) : <%=result.Item(i).net%> </li>
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
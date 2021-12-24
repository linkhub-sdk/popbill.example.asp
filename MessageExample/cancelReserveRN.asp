<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
    '문자전송요청시 할당한 전송요청번호(requestNum)로 예약문자 전송을 취소합니다.
    ' - 예약취소는 예약전송시간 10분전까지만 가능합니다.
	' - https://docs.popbill.com/message/asp/api#CancelReserveRN
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"
	
	'팝빌 회원 아이디
	userID = "testkorea"			   

	'예약문자전송 요청시 할당한 전송요청번호
	requestNum = "20211201-001"  

	On Error Resume Next

	Set result = m_MessageService.CancelReserveRN(testCorpNum, requestNum, userID)

	If Err.Number <> 0 then
		code = Err.Number
		message =  Err.Description
		Err.Clears
	Else
		code = result.code
		message = result.message
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>문자예약전송 취소</legend>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
			</fieldset>
		 </div>
	</body>
</html>
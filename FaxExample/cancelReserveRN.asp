<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 팩스전송요청시 할당한 전송요청번호(requestNum)로 팩스 예약전송건을 취소합니다.
	' - 예약전송 취소는 예약전송시간 10분전까지 가능하며, 팩스변환 이후 가능합니다.
	'**************************************************************

	' 팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	
	
	' 팝빌 회원 아이디
	userID = "testkorea"				

	' 예약팩스전송 요청시 할당한 전송요청번호
	requestNum = "20180905-fax"   

	On Error Resume Next

	Set Presponse = m_FaxService.CancelReserveRN(testCorpNum, requestNum, userID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
	End If

	On Error GoTo 0

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>팩스예약전송 취소</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
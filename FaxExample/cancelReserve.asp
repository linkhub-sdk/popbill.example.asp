<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 예약전송 팩스요청건을 취소합니다.
	' - 예약전송 취소는 예약전송시간 10분전까지 가능합니다.
	'**************************************************************

	' 팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	
	
	' 팝빌 회원 아이디
	userID = "testkorea"				

	' 팩스 전송시 발급받은 접수번호
	receiptNum = "015012713201000001"   

	On Error Resume Next

	Set Presponse = m_FaxService.CancelReserve(testCorpNum , receiptNum, userID )

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
				<legend>팩스예약전송취소</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
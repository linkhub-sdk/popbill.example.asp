<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		   '팝빌 회원 사업자번호, "-" 제외
	userID = "testkorea"			   '팝빌 회원 아이디
	receiptNum = "015012713000000010"  '문자 전송시 발급받은 접수번호(ReceiptNum)

	On Error Resume Next

	Set result = m_MessageService.CancelReserve(testCorpNum, receiptNum, userID)

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
				<legend>예약문자 메시지 취소</legend>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
			</fieldset>
		 </div>
	</body>
</html>
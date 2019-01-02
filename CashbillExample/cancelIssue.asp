<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' [발행완료] 상태의 현금영수증을 발행취소합니다.
	' - 발행취소는 국세청 전송전에만 가능합니다.
	' - 발행취소된 형금영수증은 국세청에 전송되지 않습니다.
	'**************************************************************

	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	
	
	'팝빌회원 아이디
	userID = "testkorea"

	'문서관리번호
	mgtKey = "20190103-001"				

	'메모
	memo = "현금영수증 발행취소메모"	

	On Error Resume Next

	Set Presponse = m_CashbillService.CancelIssue(testCorpNum, mgtKey, memo, UserID)
	
	If Err.Number <> 0 then
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
				<legend>현금영수증 발행취소</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
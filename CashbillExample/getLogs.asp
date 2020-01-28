<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 현금영수증 상태 변경이력을 확인합니다.
	' - https://docs.popbill.com/cashbill/asp/api#GetLogs
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	'팝빌 회원 아이디
	userID = "testkorea"		 

	'문서번호
	mgtKey = "20190103-001"		 

	On Error Resume Next
	
	Set Presponse = m_CashbillService.GetLogs(testCorpNum, mgtKey, userID)

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
				<legend>현금영수증 이력 확인</legend>
				<ul>
					<% If code = 0 Then
						For i=0 To Presponse.Count-1
					%>
						<fieldset class="fieldset2">
							 <ul>
								<li>docLogType(로그타입) : <%=Presponse.Item(i).docLogType%></li>
								<li>log(이력정보) : <%=Presponse.Item(i).log%></li>
								<li>procType(처리형태) : <%=Presponse.Item(i).procType%></li>
								<li>procMemo(처리메모) : <%=Presponse.Item(i).procMemo%></li>
								<li>regDT(등록일시) : <%=Presponse.Item(i).regDT%></li>
								<li>ip(아이피) : <%=Presponse.Item(i).ip%></li>
							</ul>
						</fieldset>
					<%	
						Next
						Else
					%>
						<li>Response.code : <%=code%></li>
						<li>Response.message : <%=message%><li>
					<%
						End If
					%>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
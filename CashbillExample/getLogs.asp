<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	 '팝빌 회원 사업자번호, "-" 제외
	userID = "testkorea"		 '팝빌 회원 아이디
	mgtKey = "20150201-01"		 '연동관리번호

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
								<li>docLogType : <%=Presponse.Item(i).docLogType%></li>
								<li>log : <%=Presponse.Item(i).log%></li>
								<li>procType : <%=Presponse.Item(i).procType%></li>
								<li>procMemo : <%=Presponse.Item(i).procMemo%></li>
								<li>regDT : <%=Presponse.Item(i).regDT%></li>
								<li>ip : <%=Presponse.Item(i).ip%></li>
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
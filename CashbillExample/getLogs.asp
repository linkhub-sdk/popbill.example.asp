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
	' - 상태 변경이력 확인(GetLogs API) 응답항목에 대한 자세한 정보는
	'   "[현금영수증 API 연동매뉴얼] > 3.4.4 상태 변경이력 확인"
	'   을 참조하시기 바랍니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	'팝빌 회원 아이디
	userID = "testkorea"		 

	'문서관리번호
	mgtKey = "20161114-01"		 

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
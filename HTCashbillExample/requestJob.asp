<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 현금영수증 매출/매입 내역 수집을 요청합니다
	' - 수집 요청후 반환받은 작업아이디(JobID)의 유효시간은 1시간 입니다.
	' - https://docs.popbill.com/htcashbill/asp/api#RequestJob
	'**************************************************************

	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		

	'발행유형 SELL(매출), BUY(매입)
	KeyType= "BUY"						

	'시작일자, 표시형식(yyyyMMdd)
	SDate = "20181201"

	'종료일자, 표시형식(yyyyMMdd)
	EDate =	"20190101"					

	'팝빌회원 아이디
	testUserID = "testkorea"			
	
	On Error Resume Next

	jobID = m_HTCashbillService.requestJob(testCorpNum, KeyType, SDate, EDate, testUserID)

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
				<legend>수집 요청</legend>
				<% If code = 0 Then %>
					<ul>
						<li>jobID(작업아이디) : <%=jobID%> </li>
					</ul>
				<%	Else  %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>
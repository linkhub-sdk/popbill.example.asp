<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 수집 요청 상태를 확인합니다.
	' - https://docs.popbill.com/htcashbill/asp/api#GetJobState
	'**************************************************************

	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	'팝빌회원 아이디
	UserID = "testkorea"

		'수집요청시 반환받은작업아이디(jobID)
	JobID = "016111417000000002"	
	
	On Error Resume Next
	
	Set result = m_HTCashbillService.GetJobState(testCorpNum, JobID, UserID)

	If Err.Number <> 0 Then
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
				<legend>수집 상태 확인</legend>
				<%
					If code = 0 Then
				%>
						<ul>
							<li> jobID (작업아이디) : <%=result.jobID%></li>
							<li> jobState (수집상태) : <%=result.jobState%></li>
							<li> queryType (수집유형) : <%=result.queryType%></li>
							<li> queryDateType (일자유형) : <%=result.queryDateType%></li>
							<li> queryStDate (시작일자) : <%=result.queryStDate%></li>
							<li> queryEnDate (종료일자) : <%=result.queryEnDate%></li>
							<li> errorCode (오류코드) : <%=result.errorCode%></li>
							<li> errorReason (오류메시지) : <%=result.errorReason%></li>
							<li> jobStartDT (작업 시작일시) : <%=result.jobStartDT%></li>
							<li> jobEndDT (작업 종료일시) : <%=result.jobEndDT%></li>
							<li> collectCount (수집개수) : <%=result.collectCount%></li>
							<li> regDT (수집 요청일시) : <%=result.regDT%></li>
						</ul>
				<%
					Else
				%>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	
					End If
				%>
			</fieldset>
		 </div>
	</body>
</html>

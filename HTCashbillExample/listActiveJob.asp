<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 수집 요청건들에 대한 상태 목록을 확인합니다.
	' - 수집 요청 작업아이디(JobID)의 유효시간은 1시간 입니다.
	' - 응답항목에 관한 정보는 "[홈택스 현금영수증 연계 API 연동매뉴얼]
	'   > 3.2.3. ListActiveJob (수집 상태 목록 확인)" 을 참고하시기 바랍니다.
	'**************************************************************

	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "123456780"

	'팝빌회원 아이디
	UserID = "testkorea"					
	
	On Error Resume Next

	Set result = m_HTCashbillService.ListActiveJob(testCorpNum, UserID)
	
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
				<legend>수집 목록 조회</legend>
				<%
					If code = 0 Then
						For i=0 To result.Count-1
				%>
							<fieldset class="fieldset2">					
								<legend>ListActiveJob [ <%=i+1%> / <%=result.Count%> ] </legend>
									<ul>
										<li> jobID (작업아이디) : <%=result.Item(i).jobID%></li>
										<li> jobState (수집상태) : <%=result.Item(i).jobState%></li>
										<li> queryType (수집유형) : <%=result.Item(i).queryType%></li>
										<li> queryDateType (일자유형) : <%=result.Item(i).queryDateType%></li>
										<li> queryStDate (시작일자) : <%=result.Item(i).queryStDate%></li>
										<li> queryEnDate (종료일자) : <%=result.Item(i).queryEnDate%></li>
										<li> errorCode (오류코드) : <%=result.Item(i).errorCode%></li>
										<li> errorReason (오류메시지) : <%=result.Item(i).errorReason%></li>
										<li> jobStartDT (작업 시작일시) : <%=result.Item(i).jobStartDT%></li>
										<li> jobEndDT (작업 종료일시) : <%=result.Item(i).jobEndDT%></li>
										<li> collectCount (수집개수) : <%=result.Item(i).collectCount%></li>
										<li> regDT (수집 요청일시) : <%=result.Item(i).regDT%></li>
									</ul>
								</fieldset>
				<%
						Next
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

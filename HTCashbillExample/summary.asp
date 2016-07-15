<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'팝빌회원 사업자번호, "-" 제외
	UserID = "testkorea"				'팝빌회원 아이디
	
	'수집 요청(requestJob) 시 반환받은 작업아이디(jobID)
	JobID = "016071511000000009"

	'현금영수증 배열 N-일반현금영수증, C-취소현금영수증
	Dim TradeType(2) 
	TradeType(0) = "N"
	TradeType(1) = ""

	'거래용도 배열, P-소득공제용, C-지출증빙용
	Dim TradeUsage(2)
	TradeUsage(0) = "P"
	TradeUsage(1) = ""
	
	Set result = m_HTCashbillService.Summary(testCorpNum, JobID, TradeType, TradeUsage, UserID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If


%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>수집 결과 조회</legend>
				<%
					If code = 0 Then
				%>
					<ul>
						<li> count (수집 결과 건수) : <%=result.count%> </li>
						<li> supplyCostTotal (공급가액 합계) : <%=result.supplyCostTotal%> </li>
						<li> taxTotal (세액 합계) : <%=result.taxTotal%> </li>
						<li> serviceFeeTotal (봉사료 합계) : <%=result.serviceFeeTotal%> </li>
						<li> amountTotal (합계 금액) : <%=result.amountTotal%> </li>
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
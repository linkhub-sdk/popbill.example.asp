<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'회원 사업자번호, "-" 제외
	KeyType= SELL						'발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	DType = "W"							'검색 일자유형, W-작성일자, I-발행일자, S-전송일자
	SDate = "20160601"					'시작일자, 표시형식(yyyyMMdd)
	EDate =	"20160831"					'종료일자, 표시형식(yyyyMMdd)
	testUserID = "testkorea"		'회원 아이디
	
	On Error Resume Next

	'수집요청시 반환되는 jobID의 유효시간은 1시간 입니다.
	jobID = m_HTTaxinvoiceService.requestJob(testCorpNum, KeyType, DType, SDate, EDate, testUserID)

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
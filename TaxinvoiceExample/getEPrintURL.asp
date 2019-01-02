<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 세금계산서 인쇄(공급받는자) URL을 반환합니다.
	' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	' 팝빌회원 아이디
	userID = "testkorea"

	' 세금계산서 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType = "SELL"

	' 문서관리번호 
	MgtKey = "20190103-001"

	On Error Resume Next
	
	url = m_TaxinvoiceService.GetEPrintURL(testCorpNum, KeyType, MgtKey, userID)

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
				<legend>세금계산서 인쇄 팝업 URL - 공급받는자용 </legend>
					<ul>
					<% If code = 0 Then%>
						<li>URL : <%=url%> </li>
					<% Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If%>		
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
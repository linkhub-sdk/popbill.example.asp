<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 대량건의 전자세금계산서 인쇄팝업 URL을 반환합니다. (최대 100건)
	' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	' 팝빌회원 아이디
	userID = "testkorea"

	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"
	
	' 인쇄할 세금계산서 문서관리번호 배열, 최대 100건
	Dim MgtKeyList(3) 
	MgtKeyList(0) = "20161114-02"
	MgtKeyList(1) = "20150121-02"
	MgtKeyList(2) = "20150121-03"
	
	On Error Resume Next

	url = m_TaxinvoiceService.GetMassPrintURL(testCorpNum, KeyType, MgtKeyList, userID)

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
				<legend>세금계산서 인쇄 URL - 대량</legend>
				<ul>
					<% If code = 0 Then %>
						<li>URL : <%=url%> </li>
					<% Else %>
						<li> Response.code : <%=code%> </li>
						<li> Response.message : <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
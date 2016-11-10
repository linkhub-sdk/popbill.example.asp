<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 1건의 전자명세서 보기 팝업 URL을 반환합니다.
	' - 보안정책으로 인해 반환된 URL의 유효시간은 30초입니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	'팝빌 회원 아이디
	userID = "testkorea"

	'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	itemCode = "121"

	'문서관리번호
	mgtKey = "20150201-01"

	On Error Resume Next

	url = m_StatementService.GetPopUpURL(testCorpNum, itemCode, mgtKey, userID)

	If Err.Number <> 0 then
		code = Err.Number
		message =  Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend> 전자명세서 팝업 URL 요청</legend>
				<% If code = 0 Then %>
					<ul>
						<li>URL : <%=CStr(url)%> </li>
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
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'팝빌 회원 사업자번호
	testCorpNum = "1234567890"		

	'첨부할 명세서 종류코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	itemCode = 121						

	'첨부할 문서관리번호
	mgtKey = "20160126-54"			

	'첨부해제할 명세서 종류코드- 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	subItemCode = 121					

	'첨부해제할 명세서 관리번호 
	subMgtKey = "20160126-31"		

	On Error Resume Next

	Set result = m_StatementService.DetachStatement(testCorpNum, itemCode, mgtKey, subItemCode, subMgtKey)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = result.code
		message = result.message
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>다른 전자명세서 첨부해제</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
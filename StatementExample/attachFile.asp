<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 전자명세서에 첨부파일을 등록합니다.
	' - 첨부파일 등록은 전자명세서가 [임시저장] 상태인 경우에만 가능합니다.
	' - 첨부파일은 최대 5개까지 등록할 수 있습니다.
	' - https://docs.popbill.com/statement/asp/api#AttachFile
	'**************************************************************

	'팝빌 회원 사업자번호, "-"제외 10자리
	testCorpNum = "1234567890"
	
	'팝빌 회원 아이디
	userID = "testkorea"

	'명세서 종류코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
	itemCode = "121"

	'문서관리번호
	mgtKey = "20211201-001"

	'첨부 파일 경로
	filePath = "C:\popbill.example.asp\Popbill\로고.gif"

	On Error Resume Next

	Set result = m_StatementService.AttachFile(testCorpNum, itemCode, mgtKey, filePath, userID)

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
				<legend>전자명세서 파일첨부</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
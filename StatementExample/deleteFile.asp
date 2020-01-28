<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 전자명세서에 첨부된 파일을 삭제합니다.
	' - 파일을 식별하는 파일아이디는 첨부파일 목록(GetFileList API) 의 응답항목
	'   중 파일아이디(AttachedFile) 값을 통해 확인할 수 있습니다.
	' - https://docs.popbill.com/statement/asp/api#DeleteFile
	'**************************************************************

	'팝빌 회원 사업자번호, "-"제외 10자리
	testCorpNum = "1234567890"
	
	'팝빌 회원 아이디
	userID = "testkorea"

	'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
	itemCode = "121"

	'문서번호
	mgtKey = "20190103-001"

	'파일아이디, 첨부파일목록(getFiles) API의 AttachedFile값
	FileID = "2556D18D-9380-4843-B748-5B8120C31BA5.PBF"

	On Error Resume Next

	Set result = m_StatementService.DeleteFile(testCorpNum, itemCode, mgtKey, FileID, userID)

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
				<legend>첨부파일 삭제</legend>
					<ul>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>
			</fieldset>
		 </div>
	</body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 세금계산서에 첨부파일을 등록합니다.
	' - [임시저장] 상태의 세금계산서만 파일을 첨부할수 있습니다.
	' - 첨부파일은 최대 5개까지 등록할 수 있습니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	' 팝빌회원 아이디
	testUserID = "testkorea"
	
	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"

	' 문서번호 
	MgtKey = "20190103-001"

	' 첨부할 파일경로
	filePath = "C:\popbill.example.asp\Popbill\로고.gif"

	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.AttachFile(testCorpNum, KeyType ,MgtKey, filePath, testUserID)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
	End If
	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>세금계산서 첨부파일 추가 </legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
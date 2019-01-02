<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 발행예정 세금계산서를 [취소] 처리 합니다.
	' - [취소]된 세금계산서를 삭제(Delete API)하면 등록된 문서관리번호를
	'   재사용할 수 있습니다.
	'**************************************************************

	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	'팝빌회원 아이디
	testUserID = "testkorea"
	
	'발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"             

	'문서관리번호 
	MgtKey = "20190103-001"     
	
	'메모
	Memo = "발행예정 메모"    
	
	'발행예정 안내메일 제목, 공백으로 처리시 기본문구로 전송
	EmailSubject = ""

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.Send(testCorpNum, KeyType ,MgtKey, Memo, EmailSubject, testUserID)
	
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
				<legend>세금계산서 발행예정</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
				
			</fieldset>
		 </div>
	</body>
</html>
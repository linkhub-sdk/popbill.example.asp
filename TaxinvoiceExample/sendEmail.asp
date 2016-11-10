<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 발행 안내메일을 재전송합니다.
	' - 메일내용중 전자세금계산서 [보기] 버튼이 동작하지 않는 경우,
	'   키보드 왼쪽 Shift 키를 누르고 버튼을 클릭해보시기 바랍니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 팝빌회원 아이디
	testUserID = "testkorea"   
	
	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"             

	' 문서관리번호 
	MgtKey = "20150121-17"      

	'이메일주소
	Receiver = "test@test.com"    

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.SendEmail(testCorpNum, KeyType, MgtKey, Receiver, testUserID)
	
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
				<legend>세금계산서 메일 재전송</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
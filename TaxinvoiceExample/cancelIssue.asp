<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	'[발행완료] 상태의 세금계산서를 [발행취소] 처리합니다.
	' - [발행취소]는 국세청 전송전에만 가능합니다.
	' - 발행취소된 세금계산서는 국세청에 전송되지 않습니다.
	' - 발행취소 세금계산서에 기재된 문서관리번호를 재사용 하기 위해서는
	'   삭제(Delete API)를 호출하여 [삭제] 처리 하셔야 합니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 팝빌회원 아이디
	testUserID = "testkorea"   
	
	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "SELL"             

	' 문서관리번호 
	MgtKey = "20161114-02"
	
	' 메모
	Memo = "발행취소 메모"      

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.CancelIssue(testCorpNum, KeyType ,MgtKey, Memo, testUserID)
	
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
				<legend>세금계산서 발행취소</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
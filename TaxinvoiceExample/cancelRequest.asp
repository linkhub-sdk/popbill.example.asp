<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 역발행 세금계산서를 [취소] 처리합니다.
	' - [취소]한 세금계산서의 문서관리번호를 재사용하기 위해서는 삭제 (Delete API)
	'   를 호출해야 합니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	' 팝빌회원 아이디
	testUserID = "testkorea"

	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType = "BUY"

	' 문서관리번호 
	MgtKey = "20190103-001"

	'메모
	Memo = "역발행 요청 취소 메모"     

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.CancelRequest(testCorpNum, KeyType ,MgtKey, Memo, testUserID)
	
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
				<legend>세금계산서 역발행요청 취소</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
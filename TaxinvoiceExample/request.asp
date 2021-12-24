<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 공급받는자가 1건의 역발행 세금계산서를 발행요청합니다.
	' - 역발행 세금계산서 프로세스를 구현하기 위해서는 공급자/공급받는자가 모두
	'   팝빌에 회원이여야 합니다.
	' - 역발행 요청후 공급자가 [발행] 처리시 포인트가 차감되며 역발행
	'   세금계산서 항목중 과금방향(ChargeDirection) 에 기재한 값에 따라
	'   정과금(공급자과금) 또는 역과금(공급받는자과금) 처리됩니다.
	' - https://docs.popbill.com/taxinvoice/asp/api#Request
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"

	' 팝빌회원 아이디
	testUserID = "testkorea"

	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType = "BUY"

	' 문서번호 
	MgtKey = "20211201-001"

	' 메모
	Memo = "역발행 요청 메모"

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.Request(testCorpNum, KeyType ,MgtKey, Memo, testUserID)
	
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
				<legend>세금계산서 역발행요청</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
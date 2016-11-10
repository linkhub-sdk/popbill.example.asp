<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 발행예정 세금계산서를 [거부]처리 합니다.
	' - [거부]처리된 세금계산서를 삭제(Delete API)하면 등록된 문서관리번호를
	'   재사용할 수 있습니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외 10자리
	testCorpNum = "1231212312"	  

	' 팝빌회원 아이디
	testUserID = "userid"		  

	' 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType= "BUY"				  

	' 문서관리번호 
	MgtKey = "20150122-23"        

	' 메모
	Memo = "발행예정거부 메모"    

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.Deny(testCorpNum, KeyType ,MgtKey, Memo ,testUserID)
	
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
				<legend>세금계산서 발행예정 거부</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
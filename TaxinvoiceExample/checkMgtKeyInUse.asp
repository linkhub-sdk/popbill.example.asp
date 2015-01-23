<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	 '회원 사업자번호, "-" 제외
	testUserID = "testkorea"	 '회원 아이디
	mgtKey = "20150122-00"		 '연동관리번호
	keyType = "SELL"			 '발행형태, (SELL-매출) (BUY-매입) (TRUSTEE-위수탁)

	On Error Resume Next

	checkMgtKeyInUse = m_TaxinvoiceService.CheckMgtKeyInUse(testCorpNum, keyType, mgtKey)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
	Else	
		If checkMgtKeyInUse = True Then
			code = 1
			message = "사용중"
		Else
			code = 0 
			message = "미사용중"
		End If
	End If 

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>연동관리번호 사용여부 확인</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
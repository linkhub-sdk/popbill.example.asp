<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'연동회원 사업자번호
	testCorpNum = "1234567890"

	'국세청승인번호 
	NTSConfirmNum = "2016071441000029000007fa"
	
	' 연동회원 아이디 
	UserID = "innoposttest"				 
	
	On Error Resume Next

	Set result = m_HTTaxinvoiceService.GetXML ( testCorpNum, NTSConfirmNum, UserID )

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend> 상세정보 조회 - XML</legend>
				<%
					If code = 0 Then
				%>
					<ul>
						<li> ResultCode (상태코드) : <%=result.ResultCode%></li>
						<li> Message (국세청승인번호) : <%=result.Message%></li>
						<li> retObject (전자세금계산서 XML 문서) : <%=Replace(result.retObject, "<" ,"&lt")%></li>
					</ul>
				<%
					Else
				%>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	
					End If
				%>
			</fieldset>
		 </div>
	</body>
</html>

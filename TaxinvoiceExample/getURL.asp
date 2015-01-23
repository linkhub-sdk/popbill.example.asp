<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	 '회원 사업자번호, "-" 제외
	userID = "testkorea"  ' 회원 아이디
	TOGO = "PBOX"   'TBOX(임시문서함), SBOX(매출문서함), PBOX(매입문서함)

	On Error Resume Next

	url = m_TaxinvoiceService.GetURL(testCorpNum, userID, TOGO)

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
		Response.End
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>팝빌 전자세금계산서 문서함 URL</legend>
				<ul>
					<li>URL : <%=url%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
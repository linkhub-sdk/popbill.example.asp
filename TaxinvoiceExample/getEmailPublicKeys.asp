<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		 ' 회원 사업자번호, "-" 제외
	
	On Error Resume Next

	Set Presponse = m_TaxinvoiceService.GetEmailPublicKeys(testCorpNum)

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
		Response.end
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>대용량 연계사업자 이메일 목록 확인 </legend>
				<fieldset class="filedset2">
				<ul>
				<%
					For i=0 To Presponse.length -1
				%>
						<li> <%=Presponse.Get(i).email%></li>
				<%
					Next
				%>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
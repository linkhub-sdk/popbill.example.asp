<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	 'ȸ�� ����ڹ�ȣ, "-" ����
	userID = "testkorea"  ' ȸ�� ���̵�
	TOGO = "PBOX"   'TBOX(�ӽù�����), SBOX(���⹮����), PBOX(���Թ�����)

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
				<legend>�˺� ���ڼ��ݰ�꼭 ������ URL</legend>
				<ul>
					<li>URL : <%=url%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���ڸ����� ÷�ε� ������ �����մϴ�.
	' - ������ �ĺ��ϴ� ���Ͼ��̵�� ÷������ ���(GetFileList API) �� �����׸�
	'   �� ���Ͼ��̵�(AttachedFile) ���� ���� Ȯ���� �� �ֽ��ϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
	testCorpNum = "1234567890"	
	
	'�˺� ȸ�� ���̵�
	userID = "testkorea"				

	'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
	itemCode = "121"					

	'����������ȣ
	mgtKey = "20150201-01"				

	'���Ͼ��̵�, ÷�����ϸ��(getFiles) API�� AttachedFile��
	FileID = "96D6F12A-0192-469F-9306-D56B2A9DB939.PBF"		

	On Error Resume Next

	Set result = m_StatementService.DeleteFile(testCorpNum, itemCode, mgtKey, FileID, userID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = result.code
		message = result.message
	End If

	On Error GoTo 0
	
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>÷������ ����</legend>
					<ul>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>
			</fieldset>
		 </div>
	</body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' [�ӽ�����] ������ ���ݰ�꼭�� [����]ó�� �մϴ�.
	' - ����(Issue API)�� ȣ���ϴ� �������� ����Ʈ�� �����˴ϴ�.
	' - [����Ϸ�] ���ݰ�꼭�� ����ȸ���� ����û ���ۼ����� ����
	'   ����/������� ó���˴ϴ�. �⺻����(��������)
	' - ����û ���ۼ����� "�˺� �α���" > [���ڼ��ݰ�꼭] > [ȯ�漳��] >
	'   [���ڼ��ݰ�꼭 ����] > [����û ���� �� �������� ����] �ǿ���
	'   Ȯ���� �� �ֽ��ϴ�.
	' - ����û ������å�� ���� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] >
	'   1.4. ����û ���� ��å" �� �����Ͻñ� �ٶ��ϴ�
	'**************************************************************

	' �˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"

	' �˺�ȸ�� ���̵�
	testUserID = "testkorea"

	' ���ݰ�꼭 �������� SELL(����), BUY(����), TRUSTEE(����Ź)
	KeyType= "SELL"

	' ����������ȣ 
	MgtKey = "20190227-023"
	
	' �޸�
	Memo = "���� �޸�"

	' ���� �ȳ����� ����, �̱���� �⺻������� ����
	EmailSubject = ""
	
	' �������� ��������, �⺻�� - False
    ' ���ึ������ ���� ���ݰ�꼭�� �����ϴ� ���, ���꼼�� �ΰ��� �� �ֽ��ϴ�.
    ' �������� ���ݰ�꼭�� �Ű��ؾ� �ϴ� ��� forceIssue ���� True�� 
	' �����Ͽ� ����(Issue API)�� ȣ���� �� �ֽ��ϴ�.
	ForceIssue = False

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.Issue(testCorpNum, KeyType ,MgtKey, Memo ,EmailSubject, ForceIssue, testUserID)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		ntsConfirmNum = ""
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
		ntsConfirmNum = Presponse.ntsConfirmNum
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>���ݰ�꼭 ����</legend>
				<ul>
					<li>�����ڵ� (Response.code) : <%=code%> </li>
					<li>����޽��� (Response.message) : <%=message%> </li>
					<% If ntsConfirmNum <> "" Then %>
					<li>����û���ι�ȣ (Response.ntsConfirmNum) : <%=ntsConfirmNum%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
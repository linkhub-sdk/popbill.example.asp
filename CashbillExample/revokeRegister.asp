<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
    ' 1���� ������ݿ������� �ӽ����� �մϴ�.
    ' - https://docs.popbill.com/cashbill/asp/api#RevokeRegister
	'**************************************************************

	' �˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	

	' �˺� ȸ�� ���̵�
	userID = "testkorea"				 

	' ������ȣ, ������ ����ڴ��� ������ȣ �Ҵ�, 1~24�ڸ� ����,������������ �ߺ����� ����.
	mgtKey = "20190103-001"

	' ���� ���ݿ����� ����û���ι�ȣ
	orgConfirmNum = "820116333"

	' ���� ���ݿ����� �ŷ�����
	orgTradeDate = "20181231"

	' ����ȳ� ���� ���ۿ���
	smssendYN = False

	On Error Resume Next

	Set Presponse = m_CashbillService.RevokeRegister(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, userID)

	If Err.Number <> 0 then
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
				<legend>������ݿ����� �ӽ�����</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 1���� (�κ�) ������ݿ������� ��ù����մϴ�.
	' - ���ݿ����� ����û ���� ��å : https://docs.popbill.com/cashbill/ntsSendPolicy?lang=asp
	' - https://docs.popbill.com/cashbill/asp/api#RevokeRegistIssue_Part
	'**************************************************************

	' �˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	

	' �˺� ȸ�� ���̵�
	userID = "testkorea"				 

	' ������ȣ, ������ ����ڹ�ȣ ���� ������ȣ �Ҵ�, 1~24�ڸ� ����,������������ �ߺ����� ����.
	mgtKey = "20171115-01"

	' ���� ���ݿ����� ����û���ι�ȣ
	orgConfirmNum = "820116333"

	' ���� ���ݿ����� �ŷ�����
	orgTradeDate = "20170711"

	' ����ȳ� ���� ���ۿ���
	smssendYN = False

	' �޸�
	memo = "��ù��� �޸�"
	
	'�κ���ҿ���, True-�κ����, False-��ü���
	isPartCancel = True

	'��һ���, 1-�ŷ����, 2-�����߱� ���, 3-��Ÿ
	cancelType = 1

	'[���] ���ް���
	supplyCost = "5000"

	'[���] ����
	tax = "500"

	'[���] �����
	serviceFee = "0"
	
	'[���] �հ�ݾ�
	totalAmount = "5500"

	On Error Resume Next

	Set Presponse = m_CashbillService.RevokeRegistIssue_Part(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID, _
		isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount)

	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
		confirmNum = Presponse.confirmNum
		tradeDate = Presponse.tradeDate		
	End If

	On Error GoTo 0 

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>(�κ�) ������ݿ����� ��ù���</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
					<li> Response.confirmNum : <%=confirmNum%> </li>
					<li> Response.tradeDate : <%=tradeDate%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
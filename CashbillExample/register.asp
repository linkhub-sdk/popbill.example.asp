<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	 '�˺� ȸ�� ����ڹ�ȣ, "-" ����
	userID = "testkorea"		 '�˺� ȸ�� ���̵�
	mgtKey = "20150201-01"		 '����������ȣ, �����ں� ������ȣ �Ҵ�, 1~24�ڸ� ����,������������ �ߺ����� ����.

	Set CashbillObj = New CashBill

    CashbillObj.mgtKey = mgtKey				   
    CashbillObj.tradeType = "���ΰŷ�"				'���ΰŷ� or ��Ұŷ�
    CashbillObj.franchiseCorpNum = testCorpNum		'������ ����ڹ�ȣ
    CashbillObj.franchiseCorpName = "������ ��ȣ"
    CashbillObj.franchiseCEOName = "������ ��ǥ��"
    CashbillObj.franchiseAddr = "������ �ּ�"
    CashbillObj.franchiseTEL = "070-1234-1234"
    CashbillObj.identityNum = "01041680206"
    CashbillObj.customerName = "����"
    CashbillObj.itemName = "��ǰ��"
    CashbillObj.orderNumber = "�ֹ���ȣ"
    CashbillObj.email = "test@test.com"
    CashbillObj.hp = "111-1234-1234"
    CashbillObj.fax = "777-444-3333"			
    CashbillObj.serviceFee = "0"				'�����
    CashbillObj.supplyCost = "10000"			'���ް���
    CashbillObj.tax = "1000"					'�ΰ���
    CashbillObj.totalAmount = "11000"			'�ŷ��ݾ�
    CashbillObj.tradeUsage = "�ҵ������"       '�ҵ������ or ����������
    CashbillObj.taxationType = "����"           '���� or �����
    
	CashbillObj.smssendYN = False				'����� �ȳ����� �ڵ����ۿ���

	On Error Resume Next

	Set Presponse = m_CashbillService.Register(testCorpNum, CashbillObj, UserID)

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
				<legend>�˺� ���ݿ����� ���</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
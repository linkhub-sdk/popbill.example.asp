<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	' �˺� ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	

	' �˺� ȸ�� ���̵�
	userID = "testkorea"				 


	' ����������ȣ, �����ں� ������ȣ �Ҵ�, 1~24�ڸ� ����,������������ �ߺ����� ����.
	mgtKey = "20170718-04"

	' �޸�
	memo = "��ù��� �޸�"

	' ���ݿ����� ��ü ����
	Set CashbillObj = New CashBill
	
    CashbillObj.mgtKey = mgtKey				   

    '���ݿ����� ����, [���ΰŷ�, ��Ұŷ�] �� ����
    CashbillObj.tradeType = "��Ұŷ�"

	'[��Ұŷ��� �ʼ�] ���� ���ݿ����� ����û���ι�ȣ
	CashbillObj.orgConfirmNum = "820116333"

	'[��Ұŷ��� �ʼ�] ���� ���ݿ����� �ŷ�����
	CashbillObj.orgTradeDate = ""

	'������ ����ڹ�ȣ, "-" ���� 10�ڸ�
	CashbillObj.franchiseCorpNum = testCorpNum	

    '������ ��ȣ��
    CashbillObj.franchiseCorpName = "������ ��ȣ"

    '������ ��ǥ�� ����
	CashbillObj.franchiseCEOName = "������ ��ǥ��"
    
    '������ �ּ�
	CashbillObj.franchiseAddr = "������ �ּ�"
    
    '������ ����ó
	CashbillObj.franchiseTEL = "070-1234-1234"
    
    '�ŷ�ó �ĺ���ȣ, �ŷ������� ���� �ۼ�
    '�ҵ������ - �ֹε��/�޴���/ī���ȣ ���簡��
    '���������� - ����ڹ�ȣ/�ֹε��/�޴���/ī���ȣ ���簡��	
	CashbillObj.identityNum = "0101112222"

    '�ŷ�����, [�ҵ������, ����������] �� ����
	CashbillObj.tradeUsage = "�ҵ������"
    
    '��������, [����, �����] �� ����
	CashbillObj.taxationType = "����" 
	
	'���ް���
	CashbillObj.supplyCost = "10000"			

	'����
	CashbillObj.tax = "1000"					

	'�����
	CashbillObj.serviceFee = "0"				
    
	'�հ�ݾ�, ���ް��� + ����� + ����
	CashbillObj.totalAmount = "11000"			
    
	
    '�ֹ�����
	CashbillObj.customerName = "����"
    
	'��ǰ��	
	CashbillObj.itemName = "��ǰ��"
    
	'�ֹ���ȣ
	CashbillObj.orderNumber = "�ֹ���ȣ"
    
	'�� �����ּ�
	CashbillObj.email = "test@test.com"
    
	'�� �޴�����ȣ
	CashbillObj.hp = "111-1234-1234"
    
	'�� �ѽ���ȣ
	CashbillObj.fax = "777-444-3333"			
    
	'����ȳ����� ���ۿ���
	'�ȳ����� ���۽� ����Ʈ�� �����Ǹ�, ���۽��н� ȯ��ó���˴ϴ�.
	CashbillObj.smssendYN = False

	On Error Resume Next

	Set Presponse = m_CashbillService.RegistIssue(testCorpNum, CashbillObj, memo, userID)

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
				<legend>���ݿ����� ��ù���</legend>
				<ul>
					<li> Response.code : <%=code%> </li>
					<li> Response.message : <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
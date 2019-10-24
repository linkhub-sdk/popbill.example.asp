<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �˺��� ������� �ʰ� ���ڸ����� �ѽ������մϴ�.
	' - �ѽ� ���� ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
	' - ���۳��� Ȯ���� "�˺� �α���" > [���� �ѽ�] > [�ѽ�] > [���۳���]
	'   �޴����� ���۰���� Ȯ���� �� �ֽ��ϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
	testCorpNum = "1234567890"

	'�˺� ȸ�� ���̵�	
	userID = "testkorea"

	'�߽Ź�ȣ
	sendNum = "07043042991"

	'�����ѽ���ȣ
	receiveNum = "010111222"

	'������ȣ 
	mgtKey = "20190103-002"
	

	'���ڸ��� ��ü ����
	Set newStatement = New Statement

    '[�ʼ�] ����� �ۼ�����, ��¥����(yyyyMMdd)
    newStatement.writeDate = "20190103"  

	'[�ʼ�] {����, û��} �� ����
    newStatement.purposeType = "����"

    '[�ʼ�] ��������, {����, ����, �鼼} �� ����
    newStatement.taxType = "����"                   

    '�������ڵ�, ����ó���� �⺻������� �ۼ�
    newStatement.formCode = ""						
	
	'[�ʼ�] ���� �����ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
    newStatement.itemCode = "121"				

    '[�ʼ�] ������ȣ, ����, ����, '-', '_' ���� (�ִ�24�ڸ�)���� ����ں��� �ߺ����� �ʵ��� ����   
    newStatement.mgtKey = mgtKey
    


	'**************************************************************
    '				                              �߽��� ����
	'**************************************************************

    '�߽��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    newStatement.senderCorpNum = testCorpNum

    '�߽��� ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    newStatement.senderTaxRegID = ""

	'�߽��� ��ȣ
    newStatement.senderCorpName = "�߽��� ��ȣ"

    '�߽��� ��ǥ�ڼ���
    newStatement.senderCEOName = "�߽���"" ��ǥ�� ����"

	'�߽��� �ּ�
    newStatement.senderAddr = "�߽��� �ּ�"

	'�߽��� ����
    newStatement.senderBizClass = "�߽��� ����"

	'�߽��� ����
    newStatement.senderBizType = "�߽��� ����,����2"

	'�߽��� ����� ����
    newStatement.senderContactName = "�߽��� ����ڸ�"

	'�߽��� �����ּ�
    newStatement.senderEmail = "test@test.com"

	'�߽��� ����ó
    newStatement.senderTEL = "070-7070-0707"

	'�߽��� �޴�����ȣ
    newStatement.senderHP = "010-000-2222"



	'**************************************************************
    '				                      ������ ����
	'**************************************************************
    
    '������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    newStatement.receiverCorpNum = "8888888888"

    '������ ��ȣ
    newStatement.receiverCorpName = "������ ��ȣ"

    '������ ��ǥ�� ����
    newStatement.receiverCEOName = "������ ��ǥ�� ����"

    '������ �ּ�
    newStatement.receiverAddr = "������ �ּ�"

    '������ ����
    newStatement.receiverBizClass = "������ ����"

    '������ ����
    newStatement.receiverBizType = "������ ����"

    '������ ����ڸ�
    newStatement.receiverContactName = "������ ����ڸ�"

    '������ �����ּ�
    newStatement.receiverEmail = "test@receiver.com"

	'������ ����ó
	newStatement.receiverTEL = "070-4304-2991"

	'������ �޴�����ȣ
	newStatement.receiverHP = "010-111-222"



	'**************************************************************
    '				                      ���ڸ��� �������
	'**************************************************************	

    '[�ʼ�] ���ް��� �հ�
	newStatement.supplyCostTotal = "100000"

	'[�ʼ�] ���� �հ�
    newStatement.taxTotal = "10000"

    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    newStatement.totalAmount = "110000"
    
    '���� �� �Ϸù�ȣ �׸�
    newStatement.serialNum = "123"

    '���� �� ��� �׸�
    newStatement.remark1 = "���1"
    newStatement.remark2 = "���2"
    newStatement.remark3 = "���3"
    
			
	'����ڵ���� �̹��� ÷�ο���
    newStatement.businessLicenseYN = False 

	'����纻 �̹��� ÷�ο���
    newStatement.bankBookYN = False        
	
	'����� �˸����� ���ۿ���
    newStatement.smssendYN = True 
	




	'**************************************************************
    '				                      ���ڸ��� ��(ǰ��)
	'**************************************************************	

	Set newDetail = New StatementDetail

    newDetail.serialNum = "1"             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20190103"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"
    newDetail.spec = "�԰�"
    newDetail.unit = "����"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "100000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "���"
    newDetail.spare1 = "spare1"
    newDetail.spare2 = "spare2"
    newDetail.spare3 = "spare3"
    newDetail.spare4 = "spare4"
    newDetail.spare5 = "spare5"

	newStatement.AddDetail newDetail
	
	Set newDetail = New StatementDetail

    newDetail.serialNum = "2"             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20190103"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"
    newDetail.spec = "�԰�"
    newDetail.unit = "����"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "100000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "���"
    newDetail.spare1 = "spare1"
    newDetail.spare2 = "spare2"
    newDetail.spare3 = "spare3"
    newDetail.spare4 = "spare4"
    newDetail.spare5 = "spare5"

	newStatement.AddDetail newDetail
	

	'**************************************************************
	'										���ڸ��� �߰��Ӽ�
    ' - �߰��Ӽ��� ���� �ڼ��� ������ "[���ڸ��� API �����Ŵ���] >
    '   5.2. �⺻��� �߰��Ӽ� ���̺�"�� �����Ͻñ� �ٶ��ϴ�.
	'**************************************************************

	newStatement.propertyBag.Set "Balance", "150000"
	newStatement.propertyBag.Set "CBalance", "100000"


	On Error Resume Next

	Presponse = m_StatementService.FAXSend(testCorpNum, newStatement, sendNum, receiveNum, userID)

	If Err.Number <> 0 Then
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
				<legend>���ѽ� ����</legend>
				<ul>
					<% If code = 0 Then %>
						<li>recepitNum : <%=Presponse%> </li>
					<% Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>

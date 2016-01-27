<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%

	testCorpNum = "1234567890"		 ' �˺� ȸ�� ����ڹ�ȣ
	userID = "testkorea"					 ' �˺� ȸ�� ���̵�
	memo = "��ù��� �޸�"			 ' �޸� 
	mgtKey = "20160126-10"			 ' ������ȣ 

	Set newStatement = New Statement

    newStatement.writeDate = "20160126"             '�ʼ�, ����� �ۼ�����
    newStatement.purposeType = "����"               '�ʼ�, {����, û��}
    newStatement.taxType = "����"                   '�ʼ�, {����, ����, �鼼}
    newStatement.formCode = ""						'�������ڵ�(�⺻�� "")
    
    newStatement.itemCode = "121"					'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
    
    newStatement.mgtKey = mgtKey
    
    newStatement.senderCorpNum = testCorpNum
    newStatement.senderTaxRegID = "" '������� �ĺ���ȣ. �ʿ�� ����. ������ ���� 4�ڸ�.
    newStatement.senderCorpName = "������ ��ȣ"
    newStatement.senderCEOName = "������"" ��ǥ�� ����"
    newStatement.senderAddr = "������ �ּ�"
    newStatement.senderBizClass = "������ ����"
    newStatement.senderBizType = "������ ����,����2"
    newStatement.senderContactName = "������ ����ڸ�"
    newStatement.senderEmail = "test@test.com"
    newStatement.senderTEL = "070-7070-0707"
    newStatement.senderHP = "010-000-2222"
    
    newStatement.receiverCorpNum = "8888888888"
    newStatement.receiverCorpName = "���޹޴��� ��ȣ"
    newStatement.receiverCEOName = "���޹޴��� ��ǥ�� ����"
    newStatement.receiverAddr = "���޹޴��� �ּ�"
    newStatement.receiverBizClass = "���޹޴��� ����"
    newStatement.receiverBizType = "���޹޴��� ����"
    newStatement.receiverContactName = "���޹޴��� ����ڸ�"
    newStatement.receiverEmail = "test@receiver.com"
    
    newStatement.supplyCostTotal = "100000"      '�ʼ� ���ް��� �հ�
    newStatement.taxTotal = "10000"                  '�ʼ� ���� �հ�
    newStatement.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
    
    newStatement.serialNum = "123"
    newStatement.remark1 = "���1"
    newStatement.remark2 = "���2"
    newStatement.remark3 = "���3"
    
    newStatement.businessLicenseYN = False		'����ڵ���� �̹��� ÷�ν� ����.
    newStatement.bankBookYN = False				'����纻 �̹��� ÷�ν� ����.
    newStatement.faxsendYN = False				'����� Fax�߼۽� ����.
    newStatement.smssendYN = True				'����� ���ڹ߼۱�� ���� Ȱ��
	

	Set newDetail = New StatementDetail

    newDetail.serialNum = "1"             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20150110"   '�ŷ�����  yyyyMMdd
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
    newDetail.purchaseDT = "20150112"   '�ŷ�����  yyyyMMdd
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
	

	'�߰��Ӽ�, �ڼ��ѻ����� ���ڸ��� API �����Ŵ��� [5.�η� > 5.2 �⺻��� �߰��Ӽ� ���̺�] ����.
	newStatement.propertyBag.Set "Balance", "150000"
	newStatement.propertyBag.Set "CBalance", "100000"

	On Error Resume Next

	Set result = m_StatementService.RegistIssue(testCorpNum, newStatement, memo, userID)

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
				<legend>���ڸ��� ��ù���</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
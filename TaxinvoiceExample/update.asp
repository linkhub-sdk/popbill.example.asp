<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' [�ӽ�����] ������ ���ݰ�꼭�� �׸��� �����մϴ�.
	' - ���ݰ�꼭 �׸� ������ "[���ڼ��ݰ�꼭 API �����Ŵ���] > 4.1. (����)��꼭
	'   ����"�� �����Ͻñ� �ٶ��ϴ�.
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ
	testCorpNum = "1234567890"		

	'�˺�ȸ�� ���̵�
	userID = "testkorea"			

	'��������
	keyType = "SELL"				

	'������ȣ
	mgtKey = "20190103-001"		
	
	'�ŷ����� �����ۼ�����
	writeSpecificationYN = False	


	' ���ݰ�꼭 ���� ��ü ����
	Set newTaxinvoice = New Taxinvoice

	' [�ʼ�] �ۼ�����, ��¥����(yyyyMMdd)
	newTaxinvoice.writeDate = "20161114"

	' [�ʼ�] {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    newTaxinvoice.chargeDirection = "������"        
	
	' [�ʼ�] ��������, {������, ������, ����Ź} �� ����
    newTaxinvoice.issueType = "������"               

	' [�ʼ�] {����, û��} �� ���� 
    newTaxinvoice.purposeType = "����"               

	' [�ʼ�] �������, {��������, ���ν��ڵ�����}
	' ���ν��ڵ������� ��� ���࿹�� ���μ��������� �̿밡��
    newTaxinvoice.issueTiming = "��������"         
	
	' [�ʼ�] ��������,  {����, ����, �鼼} �� ���� 
    newTaxinvoice.taxType = "����"                   
    
    

	'**************************************************************
    '						                       ������ ����
	'**************************************************************

    '[�ʼ�] ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    newTaxinvoice.invoicerCorpNum = "1234567890"

	'[�ʼ�] ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    newTaxinvoice.invoicerTaxRegID = ""					

    '[�ʼ�] ������ ��ȣ
	newTaxinvoice.invoicerCorpName = "������ ��ȣ_����"

    '[�ʼ�] ������ ������ȣ, 1~24�ڸ� (����, ����, '-', '_') ��������
    '����� ���� �ߺ����� �ʵ��� ����
    newTaxinvoice.invoicerMgtKey = mgtKey		

	'[�ʼ�] ������ ��ǥ�� ����
    newTaxinvoice.invoicerCEOName = "������ ��ǥ�� ����"
    
	' ������ �ּ�
	newTaxinvoice.invoicerAddr = "������ �ּ�"
    
	' ������ ����
	newTaxinvoice.invoicerBizClass = "������ ����"
    
	' ������ ����
	newTaxinvoice.invoicerBizType = "������ ����,����2"
    
	' ������ ����ڸ�
	newTaxinvoice.invoicerContactName = "������ ����ڸ�"
    
	' ������ ����� �����ּ� 
	newTaxinvoice.invoicerEmail = "test@test.com"
    
	' ������ ����� ����ó 
	newTaxinvoice.invoicerTEL = "070-7070-0707"
    
	' ������ ����� �޴�����ȣ
	newTaxinvoice.invoicerHP = "010-000-2222"

    '������� ���޹޴��ڿ��� ����ȳ����� ���ۿ���
    '- �ȳ����� ���۱�� �̿�� ����Ʈ�� �����˴ϴ�.	
	newTaxinvoice.invoicerSMSSendYN = False
    


	'**************************************************************
    '				                            ���޹޴��� ����
	'**************************************************************

	'[�ʼ�] ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
    newTaxinvoice.invoiceeType = "�����"

    '[�ʼ�] ���޹޴��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    newTaxinvoice.invoiceeCorpNum = "1231212312"

    '[�ʼ�] ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����	
	newTaxinvoice.invoiceeTaxRegID = ""
    
	'[�ʼ�] �����ڹ޴��� ��ȣ
	newTaxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"

    '[������� �ʼ�] ���޹޴��� ������ȣ(������� �ʼ�)
    newTaxinvoice.invoiceeMgtKey = ""

	'[�ʼ�] ���޹޴��� ��ǥ�� ����
	newTaxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    
	'���޹޴��� �ּ�
	newTaxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    
	'���޹޴��� ����
	newTaxinvoice.invoiceeBizClass = "���޹޴��� ����"
    
	'���޹޴��� ����
	newTaxinvoice.invoiceeBizType = "���޹޴��� ����"
    
	'���޹޴��� ����ڸ�
	newTaxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    
	'���޹޴��� ����� �����ּ�
	newTaxinvoice.invoiceeEmail1 = "test@invoicee.com"
	
	'���޹޴��� ����ó
	newTaxinvoice.invoiceeTEL1 = "070-111-222"
	
	'���޹޴��� �޴�����ȣ
	newTaxinvoice.invoiceeHP1 = "010-111-222"

    '������� �����ڿ��� ����ȳ����� ���ۿ���
    newTaxinvoice.invoiceeSMSSendYN = False



	'**************************************************************
    '				                            ���ݰ�꼭 ����
	'**************************************************************

    '[�ʼ�] ���ް��� �հ�
    newTaxinvoice.supplyCostTotal = "100000"

    '[�ʼ�] ���� �հ�
    newTaxinvoice.taxTotal = "10000"

    '[�ʼ�] �հ�ݾ�, ���ް��� �հ� + �����հ�
	newTaxinvoice.totalAmount = "110000"             
    
    '���� �� '�Ϸù�ȣ' �׸�
    newTaxinvoice.serialNum = "123"

	'���� �� '��' �׸�, �ִ밪 32767
    newTaxinvoice.kwon = "1"

	'���� �� 'ȣ' �׸�, �ִ밪 32767
    newTaxinvoice.ho = "1"

	'���� �� '����' �׸�
    newTaxinvoice.cash = ""
    
	'���� �� '��ǥ' �׸�
    newTaxinvoice.chkBill = ""

	'���� �� '����' �׸�
    newTaxinvoice.note = ""
	
	'���� �� '�ܻ�̼���' �׸�
    newTaxinvoice.credit = ""

	'���� �� '���'�׸�
    newTaxinvoice.remark1 = "���1"
    newTaxinvoice.remark2 = "���2"
    newTaxinvoice.remark3 = "���3"

	'����ڵ���� �̹��� ÷�ο���
    newTaxinvoice.businessLicenseYN = False 

	' ����纻 �̹��� ÷�ο���
    newTaxinvoice.bankBookYN = False         
  
	
	
	'**************************************************************
    '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����
    ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - http://blog.linkhub.co.kr/650
	'**************************************************************

	' ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    newTaxinvoice.modifyCode = ""

	' �������ݰ�꼭�� ItemKey, ����Ȯ�� (GetInfo API)�� ������(ItemKey �׸�) Ȯ��
    newTaxinvoice.originalTaxinvoiceKey = ""


	'**************************************************************
	'										���׸�(ǰ��) ����
	'**************************************************************
    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20161110"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��1��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "50000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "50000"
    newDetail.tax = "5000"
    newDetail.remark = "���"

    newTaxinvoice.AddDetail newDetail

    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 2             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20161110"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��2��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "50000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "50000"
    newDetail.tax = "5000"
    newDetail.remark = "���"
    
    newTaxinvoice.AddDetail newDetail
 


	'**************************************************************
    '			                            �߰������ ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    '   ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
	'**************************************************************

    set newContact = New Contact
	newContact.serialNum = 1
    newContact.contactName = "�����1 ����"
    newContact.email = "test1@test.com"   
    newTaxinvoice.AddContact newContact

    set newContact = New Contact
	newContact.serialNum = 2
    newContact.contactName = "�����2 ����"
    newContact.email = "test2@test.com"
    newTaxinvoice.AddContact newContact
    
	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.Update(testCorpNum, keyType, mgtKey, newTaxinvoice, writeSpecificationYN, userID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = Presponse.code
		message =Presponse.message
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
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	Set newTaxinvoice = New Taxinvoice

	newTaxinvoice.writeDate = "20140122"             '�ʼ�, ����� �ۼ�����
    newTaxinvoice.chargeDirection = "������"         '�ʼ�, {������, ������}
    newTaxinvoice.issueType = "������"               '�ʼ�, {������, ������, ����Ź}
    newTaxinvoice.purposeType = "����"               '�ʼ�, {����, û��}
    newTaxinvoice.issueTiming = "��������"           '�ʼ�, {��������, ���ν��ڵ�����}
    newTaxinvoice.taxType = "����"                   '�ʼ�, {����, ����, �鼼}
    
    
    newTaxinvoice.invoicerCorpNum = "1234567890"
    newTaxinvoice.invoicerTaxRegID = ""					'������� �ĺ���ȣ. �ʿ�� ����. ������ ���� 4�ڸ�.
    newTaxinvoice.invoicerCorpName = "������ ��ȣ"
    newTaxinvoice.invoicerMgtKey = "20150122-29"		'������ ��Ʈ�� ������ȣ
    newTaxinvoice.invoicerCEOName = "������ ��ǥ�� ����"
    newTaxinvoice.invoicerAddr = "������ �ּ�"
    newTaxinvoice.invoicerBizClass = "������ ����"
    newTaxinvoice.invoicerBizType = "������ ����,����2"
    newTaxinvoice.invoicerContactName = "������ ����ڸ�"
    newTaxinvoice.invoicerEmail = "test@test.com"
    newTaxinvoice.invoicerTEL = "070-7070-0707"
    newTaxinvoice.invoicerHP = "010-000-2222"
    newTaxinvoice.invoicerSMSSendYN = False			'����� ���ڹ߼۱�� ���� Ȱ��
    
    newTaxinvoice.invoiceeType = "�����"
    newTaxinvoice.invoiceeCorpNum = "1231212312"
    newTaxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"
    newTaxinvoice.invoiceeMgtKey = ""
    newTaxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    newTaxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    newTaxinvoice.invoiceeBizClass = "���޹޴��� ����"
    newTaxinvoice.invoiceeBizType = "���޹޴��� ����"
    newTaxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    newTaxinvoice.invoiceeEmail1 = "test@invoicee.com"
    
    newTaxinvoice.supplyCostTotal = "100000"         '�ʼ� ���ް��� �հ�
    newTaxinvoice.taxTotal = "10000"                 '�ʼ� ���� �հ�
    newTaxinvoice.totalAmount = "110000"             '�ʼ� �հ�ݾ�.  ���ް��� + ����
    
    newTaxinvoice.modifyCode = ""				'�������ݰ�꼭 �ۼ��� 1~6���� ���ñ���.
    newTaxinvoice.originalTaxinvoiceKey = ""	'�������ݰ�꼭 �ۼ��� �������ݰ�꼭�� ItemKey����. ItemKey�� ����Ȯ��(getInfo.asp) API ���� Ȯ��.
    newTaxinvoice.serialNum = "123"
    newTaxinvoice.cash = ""          '����
    newTaxinvoice.chkBill = ""       '��ǥ
    newTaxinvoice.note = ""          '����
    newTaxinvoice.credit = ""        '�ܻ�̼���
    newTaxinvoice.remark1 = "���1"
    newTaxinvoice.remark2 = "���2"
    newTaxinvoice.remark3 = "���3"
    newTaxinvoice.kwon = "1"
    newTaxinvoice.ho = "1"
    
    newTaxinvoice.businessLicenseYN = False '����ڵ���� �̹��� ÷�ν� ����.
    newTaxinvoice.bankBookYN = False         '����纻 �̹��� ÷�ν� ����.
  

	'���׸� �߰�.
    
    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20140410"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "100000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "���"

    newTaxinvoice.AddDetail newDetail

    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 2
    newDetail.itemName = "ǰ��2"
    
    newTaxinvoice.AddDetail newDetail
 

	'�߰������ �߰�. �ɼ�.
    set newContact = New Contact
    newContact.contactName = "����� ����"
    newContact.email = "test2@test.com"
    
    newTaxinvoice.AddContact newContact
    
	On Error Resume Next

	testCorpNum = "1234567890"		'�˺�ȸ�� ����ڹ�ȣ
	writeSpecificationYN = False	'�ŷ����� �����ۼ�����
	userID = "testkorea"			'ȸ�� ���̵�

	Set Presponse = m_TaxinvoiceService.Register(testCorpNum, newTaxinvoice, writeSpecificationYN, userID)

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
				<legend>���ݰ�꼭 �ӽ�����</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>
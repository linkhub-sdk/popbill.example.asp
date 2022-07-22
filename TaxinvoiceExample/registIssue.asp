<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 1���� ���ݰ�꼭�� ��ù��� ó���մϴ�.
    ' - https://docs.popbill.com/taxinvoice/asp/api#RegistIssue
    '**************************************************************
    
    ' �˺�ȸ�� ����ڹ�ȣ
    testCorpNum = "1234567890"

    ' �ŷ����� �����ۼ�����
    writeSpecificationYN = False

    ' �ŷ����� �����ۼ���, �ŷ����� ������ȣ
    dealInvoiceMgtKey = ""

    ' �������� ��������  (true / false �� �� 1)
    ' �� true = ���� , false = �Ұ���
    ' - ���ึ������ ���� ���ݰ�꼭�� �����ϴ� ���, ���꼼�� �ΰ��� �� �ֽ��ϴ�.
    ' - ���꼼�� �ΰ��Ǵ��� ������ �ؾ��ϴ� ��쿡�� forceIssue�� ����
    '   true�� �����Ͽ� ����(Issue API)�� ȣ���Ͻø� �˴ϴ�.
    forceIssue = False

    ' ��ù��� �޸�
    memo = "��ù��� �޸�"

    ' ���� �ȳ����� ����
    emailSubject = ""

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    
    ' ���ݰ�꼭 ���� ��ü ����
    Set newTaxinvoice = New Taxinvoice

    ' �ۼ�����, ��¥����(yyyyMMdd)
    newTaxinvoice.writeDate = "20220721"

    ' {������} ����
    newTaxinvoice.chargeDirection = "������"
    
    ' ��������, {������, ����Ź} �� ����
    newTaxinvoice.issueType = "������"

    ' {����, û��, ����} �� ���� 
    newTaxinvoice.purposeType = "����"
    
    ' ��������, {����, ����, �鼼} �� ���� 
    newTaxinvoice.taxType = "����"
    
    

    '**************************************************************
    '                       ������ ����
    '**************************************************************

    ' ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    newTaxinvoice.invoicerCorpNum = "1234567890"

    ' ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    newTaxinvoice.invoicerTaxRegID = "0001"

    ' ������ ��ȣ
    newTaxinvoice.invoicerCorpName = "������ ��ȣ"

    ' ������ ������ȣ, 1~24�ڸ� (����, ����, '-', '_') ��������
    ' ����� ���� �ߺ����� �ʵ��� ����
    newTaxinvoice.invoicerMgtKey = "20220721-0005"

    ' ������ ��ǥ�� ����
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
    newTaxinvoice.invoicerEmail = ""
    
    ' ������ ����� ����ó 
    newTaxinvoice.invoicerTEL = ""
    
    ' ������ ����� �޴�����ȣ
    newTaxinvoice.invoicerHP = ""

    ' ���� �ȳ� ���� ���ۿ��� (true / false �� �� 1)
    ' �� true = ���� , false = ������
    ' �� ���޹޴��� (��)����� �޴�����ȣ {invoiceeHP1} ������ ���� ����
    ' - ���� �� ����Ʈ �����Ǹ�, ���۽��н� ȯ��ó��
    newTaxinvoice.invoicerSMSSendYN = False
    


    '**************************************************************
    '                      ���޹޴��� ����
    '**************************************************************

    ' ���޹޴��� ����, [�����, ����, �ܱ���] �� ����
    newTaxinvoice.invoiceeType = "�����"

    ' ���޹޴��� ����ڹ�ȣ
    ' - {invoiceeType}�� "�����" �� ���, ����ڹ�ȣ (������ ('-') ���� 10�ڸ�)
    ' - {invoiceeType}�� "����" �� ���, �ֹε�Ϲ�ȣ (������ ('-') ���� 13�ڸ�)
    ' - {invoiceeType}�� "�ܱ���" �� ���, "9999999999999" (������ ('-') ���� 13�ڸ�)
    newTaxinvoice.invoiceeCorpNum = "1234567890"

    ' ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����	
    newTaxinvoice.invoiceeTaxRegID = "0002"
    
    ' �����ڹ޴��� ��ȣ
    newTaxinvoice.invoiceeCorpName = "�˾���ư �׽�Ʈ"

    ' ���޹޴��� ��ǥ�� ����
    newTaxinvoice.invoiceeCEOName = "���޹޴��� ��ǥ�� ����"
    
    ' ���޹޴��� �ּ�
    newTaxinvoice.invoiceeAddr = "���޹޴��� �ּ�"
    
    ' ���޹޴��� ����
    newTaxinvoice.invoiceeBizClass = "���޹޴��� ����"
    
    ' ���޹޴��� ����
    newTaxinvoice.invoiceeBizType = "���޹޴��� ����"
    
    ' ���޹޴��� ����ڸ�
    newTaxinvoice.invoiceeContactName1 = "���޹޴��� ����ڸ�"
    
    ' ���޹޴��� ����� �����ּ�
    ' �˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
    ' ���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
    newTaxinvoice.invoiceeEmail1 = ""
    
    ' ���޹޴��� ����ó
    newTaxinvoice.invoiceeTEL1 = ""
    
    ' ���޹޴��� �޴�����ȣ
    newTaxinvoice.invoiceeHP1 = ""



    '**************************************************************
    '                      ���ݰ�꼭 ����
    '**************************************************************

    ' ���ް��� �հ�
    newTaxinvoice.supplyCostTotal = "100000"

    ' ���� �հ�
    newTaxinvoice.taxTotal = "10000"

    ' �հ�ݾ�, ���ް��� �հ� + �����հ�
    newTaxinvoice.totalAmount = "110000"
    
    ' ���� �� '�Ϸù�ȣ' �׸�
    newTaxinvoice.serialNum = "123"

    ' ���� �� '��' �׸�, �ִ밪 32767
    newTaxinvoice.kwon = "1"

    ' ���� �� 'ȣ' �׸�, �ִ밪 32767
    newTaxinvoice.ho = "1"

    ' ���� �� '����' �׸�
    newTaxinvoice.cash = ""
    
    ' ���� �� '��ǥ' �׸�
    newTaxinvoice.chkBill = ""

    ' ���� �� '����' �׸�
    newTaxinvoice.note = ""
    
    ' ���� �� '�ܻ�̼���' �׸�
    newTaxinvoice.credit = ""

    ' ���
    ' {invoiceeType}�� "�ܱ���" �̸� remark1 �ʼ�
    ' - �ܱ��� ��Ϲ�ȣ �Ǵ� ���ǹ�ȣ �Է�
    newTaxinvoice.remark1 = "���1"
    newTaxinvoice.remark2 = "���2"
    newTaxinvoice.remark3 = "���3"

    ' ����ڵ���� �̹��� ÷�ο���  (true / false �� �� 1)
    ' �� true = ÷�� , false = ��÷��(�⺻��)
    ' - �˺� ����Ʈ �Ǵ� �ΰ� �� ÷�ι��� ��� �˾� URL (GetSealURL API) �Լ��� �̿��Ͽ� ���
    newTaxinvoice.businessLicenseYN = False 

    ' ����纻 �̹��� ÷�ο���  (true / false �� �� 1)
    ' �� true = ÷�� , false = ��÷��(�⺻��)
    ' - �˺� ����Ʈ �Ǵ� �ΰ� �� ÷�ι��� ��� �˾� URL (GetSealURL API) �Լ��� �̿��Ͽ� ���
    newTaxinvoice.bankBookYN = False    

    
    
    '**************************************************************
    '         �������ݰ�꼭 ���� (�������ݰ�꼭 �ۼ��ÿ��� ����
    ' - �������ݰ�꼭 ���� ������ �����Ŵ��� �Ǵ� ���߰��̵� ��ũ ����
    ' - [����] �������ݰ�꼭 �ۼ���� �ȳ� - https://docs.popbill.com/taxinvoice/modify?lang=asp
    '**************************************************************

    ' [�������ݰ�꼭 ����� �ʼ�] ���������ڵ�, ���������� ���� 1~6�� ���ñ���
    newTaxinvoice.modifyCode = ""

    ' [�������ݰ�꼭 ����� �ʼ�] �������ݰ�꼭�� ����û ���ι�ȣ ����
    newTaxinvoice.orgNTSConfirmNum = ""


    '**************************************************************
    '                   ���׸�(ǰ��) ����
    '**************************************************************
    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20211201"   '�ŷ�����  yyyyMMdd
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
    newDetail.purchaseDT = "20211201"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��2��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "50000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "50000"
    newDetail.tax = "5000"
    newDetail.remark = "���"
    
    newTaxinvoice.AddDetail newDetail
 


    '**************************************************************
    '                      �߰������ ����
    ' - ���ݰ�꼭 ����ȳ� ������ ���Ź��� ���޹޴��� ����ڰ� �ټ��� ���
    '   ����� ������ �߰��Ͽ� ����ȳ������� �ټ����� ������ �� �ֽ��ϴ�.
    '**************************************************************

    set newContact = New Contact
    newContact.serialNum = 1
    newContact.contactName = "�����1 ����"
    newContact.email = "" 
    newTaxinvoice.AddContact newContact

    set newContact = New Contact
    newContact.serialNum = 2
    newContact.contactName = "�����2 ����"
    newContact.email = ""
    newTaxinvoice.AddContact newContact

    On Error Resume Next
    
    Set Presponse = m_TaxinvoiceService.RegistIssue(testCorpNum, newTaxinvoice, writeSpecificationYN, _
                                                    dealInvoiceMgtKey, forceIssue, memo, emailSubject, userID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        ntsConfirmNum = ""
        Err.Clears
    Else
        code = Presponse.code
        message =Presponse.message
        ntsConfirmNum = Presponse.ntsConfirmNum
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>���ݰ�꼭 ��ù���</legend>
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
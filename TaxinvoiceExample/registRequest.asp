<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���޹޴��ڰ� �ۼ��� ���ݰ�꼭 �����͸� �˺��� �����ϰ� �����ڿ��� �ۺ��Ͽ� ������ ��û�մϴ�.
    ' - ������ ���ݰ�꼭 ���μ����� �����ϱ� ���ؼ��� ������/���޹޴��ڰ� ��� �˺��� ȸ���̿��� �մϴ�.
    ' - ���� ��û�� ���ݰ�꼭�� "(��)������" �����̸�, �����ڰ� �˺� ����Ʈ �Ǵ� �Լ��� ȣ���Ͽ� ������ ��쿡�� ����û���� ���۵˴ϴ�.
    ' - �����ڴ� �˺� ����Ʈ�� "���� ���� �����"���� ������ ������ ������ ���ݰ�꼭�� Ȯ���� �� �ֽ��ϴ�.
    ' - �ӽ�����(Register API) �Լ��� ������ ��û(Request API) �Լ��� �� ���� ���μ����� ó���մϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/issue#RegistRequest
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"


    ' ���ݰ�꼭 ���� ��ü ����
    Set newTaxinvoice = New Taxinvoice

    ' �ۼ�����, ��¥����(yyyyMMdd)
    newTaxinvoice.writeDate = "20220720"

    ' {������, ������} �� ����, '������'�� ������ ���μ��������� �̿밡��
    newTaxinvoice.chargeDirection = "������"

    ' ��������, {������} ����
    newTaxinvoice.issueType = "������"

    ' {����, û��, ����} �� ����
    newTaxinvoice.purposeType = "����"

    ' ��������, {����, ����, �鼼} �� ����
    newTaxinvoice.taxType = "����"


    '**************************************************************
    '                       ������ ����
    '**************************************************************

    ' ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    newTaxinvoice.invoicerCorpNum = "8888888888"

    ' ������ ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    newTaxinvoice.invoicerTaxRegID = ""

    ' ������ ��ȣ
    newTaxinvoice.invoicerCorpName = "������ ��ȣ"

    ' ������ ������ȣ, 1~24�ڸ� (����, ����, '-', '_') ��������
    ' ����� ���� �ߺ����� �ʵ��� ����
    newTaxinvoice.invoicerMgtKey = ""

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
    newTaxinvoice.invoiceeCorpNum = testCorpNum

    ' ���޹޴��� ������� �ĺ���ȣ. �ʿ�� ���� 4�ڸ� ����
    newTaxinvoice.invoiceeTaxRegID = ""

    ' �����ڹ޴��� ��ȣ
    newTaxinvoice.invoiceeCorpName = "���޹޴��� ��ȣ"

    '[������� �ʼ�] ���޹޴��� ������ȣ(������� �ʼ�)
    newTaxinvoice.invoiceeMgtKey = "20220720-ASP-001"

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

    ' ������� �����ڿ��� ����ȳ����� ���ۿ���
    newTaxinvoice.invoiceeSMSSendYN = False



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
    '                       ���׸�(ǰ��) ����
    '**************************************************************
    Set newDetail = New TaxinvoiceDetail
    newDetail.serialNum = 1             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20220720"   '�ŷ�����  yyyyMMdd
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
    newDetail.purchaseDT = "20220720"   '�ŷ�����  yyyyMMdd
    newDetail.itemName = "ǰ��2��"
    newDetail.spec = "�԰�"
    newDetail.qty = "1" '����           ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.unitCost = "50000"       ' �Ҽ��� 2�ڸ����� ���ڿ��� ���簡��
    newDetail.supplyCost = "50000"
    newDetail.tax = "5000"
    newDetail.remark = "���"

    newTaxinvoice.AddDetail newDetail

    ' ��ÿ�û �޸�
    memo = "��ÿ�û �޸�"

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.RegistRequest(testCorpNum, newTaxinvoice, memo, userID)

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
                <legend>���ݰ�꼭 ��ÿ�û</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
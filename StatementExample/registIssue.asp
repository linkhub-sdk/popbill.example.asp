<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ۼ��� ���ڸ����� �����͸� �˺��� ����� ���ÿ� �����Ͽ�, "����Ϸ�" ���·� ó���մϴ�.
    ' - �˺� ����Ʈ [���ڸ�����] > [ȯ�漳��] > [���ڸ����� ����] �޴��� ����� �ڵ����� �ɼ� ������ ���� ���ڸ������� "����Ϸ�" ���°� �ƴ� "���δ��" ���·� ���� ó�� �� �� �ֽ��ϴ�.
    ' - https://developers.popbill.com/reference/statement/asp/api/issue#RegistIssue
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ȣ, 1~24�ڸ� ����, ����, '-', '_' �������� ����ں��� �ߺ����� �ʵ��� ����
    mgtKey = "20220720-ASP-001"

    ' �޸�
    memo = "��ù��� �޸�"

    ' �ȳ����� ����, ���� ����� �⺻������� ����
    emailSubject = ""


    ' ���ڸ����� ��ü ����
    Set newStatement = New Statement

    ' ����� �ۼ�����, ��¥����(yyyyMMdd)
    newStatement.writeDate = "20220720"

    ' {����, û��, ����} �� ����
    newStatement.purposeType = "����"

    ' ��������, {����, ����, �鼼} �� ����
    newStatement.taxType = "����"

    ' �������ڵ�, ����ó���� �⺻������� �ۼ�
    newStatement.formCode = ""

    ' ������ �����ڵ� - 121(�ŷ�������), 122(û����), 123(������), 124(���ּ�), 125(�Ա�ǥ), 126(������)
    newStatement.itemCode = "121"

    ' ������ȣ, ����, ����, '-', '_' ���� (�ִ�24�ڸ�)���� ����ں��� �ߺ����� �ʵ��� ����
    newStatement.mgtKey = mgtKey

    '**************************************************************
    '                             �߽��� ����
    '**************************************************************

    ' �߽��� ����ڹ�ȣ, '-' ���� 10�ڸ�
    newStatement.senderCorpNum = testCorpNum

    ' �߽��� ������� �ĺ���ȣ, �ʿ�� ����, ������ ���� 4�ڸ�
    newStatement.senderTaxRegID = ""

    ' �߽��� ��ȣ
    newStatement.senderCorpName = "�߽��� ��ȣ"

    ' �߽��� ��ǥ�ڼ���
    newStatement.senderCEOName = "�߽���"" ��ǥ�� ����"

    ' �߽��� �ּ�
    newStatement.senderAddr = "�߽��� �ּ�"

    ' �߽��� ����
    newStatement.senderBizClass = "�߽��� ����"

    ' �߽��� ����
    newStatement.senderBizType = "�߽��� ����,����2"

    ' �߽��� ����� ����
    newStatement.senderContactName = "�߽��� ����ڸ�"

    ' �߽��� �����ּ�
    newStatement.senderEmail = ""

    ' �߽��� ����ó
    newStatement.senderTEL = ""

    ' �߽��� �޴�����ȣ
    newStatement.senderHP = ""

    '**************************************************************
    '                      ������ ����
    '**************************************************************

    ' ������ ����ڹ�ȣ, '-' ���� 10�ڸ�
    newStatement.receiverCorpNum = "8888888888"

    ' ������ ��ȣ
    newStatement.receiverCorpName = "������ ��ȣ"

    ' ������ ��ǥ�� ����
    newStatement.receiverCEOName = "������ ��ǥ�� ����"

    ' ������ �ּ�
    newStatement.receiverAddr = "������ �ּ�"

    ' ������ ����
    newStatement.receiverBizClass = "������ ����"

    ' ������ ����
    newStatement.receiverBizType = "������ ����"

    ' ������ ����ڸ�
    newStatement.receiverContactName = "������ ����ڸ�"

    ' ������ �����ּ�
    ' �˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
    ' ���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
    newStatement.receiverEmail = ""

    ' ������ ����ó
    newStatement.receiverTEL = ""

    ' ������ �޴�����ȣ
    newStatement.receiverHP = ""

    '**************************************************************
    '                     ���ڸ����� �������
    '**************************************************************

    ' ���ް��� �հ�
    newStatement.supplyCostTotal = "100000"

    ' ���� �հ�
    newStatement.taxTotal = "10000"

    ' �հ�ݾ�, ���ް��� �հ� + ���� �հ�
    newStatement.totalAmount = "110000"

    ' ���� �� �Ϸù�ȣ �׸�
    newStatement.serialNum = "123"

    ' ���� �� ��� �׸�
    newStatement.remark1 = "���1"
    newStatement.remark2 = "���2"
    newStatement.remark3 = "���3"


    ' ����ڵ���� �̹��� ÷�ο���  (true / false �� �� 1)
    ' �� true = ÷�� , false = ��÷��(�⺻��)
    ' - �˺� ����Ʈ �Ǵ� �ΰ� �� ÷�ι��� ��� �˾� URL (GetSealURL API) �Լ��� �̿��Ͽ� ���
    newStatement.businessLicenseYN = False

    ' ����纻 �̹��� ÷�ο���  (true / false �� �� 1)
    ' �� true = ÷�� , false = ��÷��(�⺻��)
    ' - �˺� ����Ʈ �Ǵ� �ΰ� �� ÷�ι��� ��� �˾� URL (GetSealURL API) �Լ��� �̿��Ͽ� ���
    newStatement.bankBookYN = False

    ' ����� �˸����� ���ۿ���
    newStatement.smssendYN = True





    '**************************************************************
    '                      ���ڸ����� ��(ǰ��)
    '**************************************************************

    Set newDetail = New StatementDetail

    newDetail.serialNum = "1"             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20220720"   '�ŷ�����  yyyyMMdd
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
    newDetail.spare6 = "spare6"
    newDetail.spare7 = "spare7"
    newDetail.spare8 = "spare8"
    newDetail.spare9 = "spare9"
    newDetail.spare10 = "spare10"
    newDetail.spare11 = "spare11"
    newDetail.spare12 = "spare12"
    newDetail.spare13 = "spare13"
    newDetail.spare14 = "spare14"
    newDetail.spare15 = "spare15"
    newDetail.spare16 = "spare16"
    newDetail.spare17 = "spare17"
    newDetail.spare18 = "spare18"
    newDetail.spare19 = "spare19"
    newDetail.spare20 = "spare20"


    newStatement.AddDetail newDetail

    Set newDetail = New StatementDetail

    newDetail.serialNum = "2"             '�Ϸù�ȣ 1���� ���� ����
    newDetail.purchaseDT = "20220720"   '�ŷ�����  yyyyMMdd
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
    '                  ���ڸ����� �߰��Ӽ�
    '**************************************************************

    newStatement.propertyBag.Set "Balance", "150000"
    newStatement.propertyBag.Set "CBalance", "100000"


    On Error Resume Next

    Set result = m_StatementService.RegistIssue(testCorpNum, newStatement, memo, userID, emailSubject)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = result.code
        message = result.message
        invoiceNum = result.invoiceNum
    End If

    On Error GoTo 0

%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>���ڸ����� ��ù���</legend>
                <ul>
                    <li>�����ڵ� (Response.code) : <%=code%> </li>
                    <li>����޽��� (Response.message): <%=message%> </li>
                    <% If invoiceNum <> "" Then %>
                    <li>�˺����ι�ȣ (Response.invoiceNum) : <%=invoiceNum%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ִ� 100���� ���ݿ����� ������ �ѹ��� ��û���� �����մϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/issue#BulkSubmit
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    testCorpNum = "1234567890"

    ' ������̵�, �ִ� 36�ڸ� (����, ����, "-" ����)
    SubmitID = "20220720-ASP-BULK001"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    Dim cashbillList(99)
    for i = 0 to 99
        ' ���ݿ����� ���� ��ü ����
        Set CashbillObj = New Cashbill

        CashbillObj.mgtKey = SubmitID + CStr(i)

        ' ��������, [���ΰŷ�, ��Ұŷ�] �� ����
        CashbillObj.tradeType = "���ΰŷ�"

        ' [��Ұŷ��� �ʼ�] ���� ���ݿ����� ����û���ι�ȣ
        CashbillObj.orgConfirmNum = ""

        ' [��Ұŷ��� �ʼ�] ���� ���ݿ����� �ŷ�����
        CashbillObj.orgTradeDate = ""

        ' �ŷ�����, [�ҵ������, ����������] �� ����
        CashbillObj.tradeUsage = "�ҵ������"

        ' �ŷ�����, [�Ϲ�, ��������, ���߱���] �� ����
        ' ���Է½� �⺻�� '�Ϲ�' ó��
        CashbillObj.tradeOpt = "�Ϲ�"

        ' ��������, [����, �����] �� ����
        CashbillObj.taxationType = "����"

        ' ���ް���
        CashbillObj.supplyCost = "10000"

        ' �ΰ���
        CashbillObj.tax = "1000"

        ' �����
        CashbillObj.serviceFee = "0"

        ' �հ�ݾ�, ���ް��� + ����� + ����
        CashbillObj.totalAmount = "11000"

        ' ������ ����ڹ�ȣ, "-" ���� 10�ڸ�
        CashbillObj.franchiseCorpNum = testCorpNum

        ' ������ ������� �ĺ���ȣ
        CashbillObj.franchiseTaxRegID = ""

        ' ������ ��ȣ
        CashbillObj.franchiseCorpName = "������ ��ȣ"

        ' ������ ��ǥ�� ����
        CashbillObj.franchiseCEOName = "������ ��ǥ��"

        ' ������ �ּ�
        CashbillObj.franchiseAddr = "������ �ּ�"

        ' ������ ��ȭ��ȣ
        CashbillObj.franchiseTEL = "070-1234-1234"

        ' �ĺ���ȣ, �ŷ����п� ���� �ۼ�
        ' �� �ҵ������ - �ֹε��/�޴���/ī���ȣ(���ݿ����� ī��)/�����߱޿� ��ȣ(010-000-1234) ���簡��
        ' �� ���������� - ����ڹ�ȣ/�ֹε��/�޴���/ī���ȣ(���ݿ����� ī��) ���簡��
        ' �� �ֹε�Ϲ�ȣ 13�ڸ�, �޴�����ȣ 10~11�ڸ�, ī���ȣ 13~19�ڸ�, ����ڹ�ȣ 10�ڸ� �Է� ����
        CashbillObj.identityNum = "0100001234"

        ' �ֹ�������
        CashbillObj.customerName = "������"

        ' �ֹ���ǰ��
        CashbillObj.itemName = "��ǰ��"

        ' �ֹ���ȣ
        CashbillObj.orderNumber = "�ֹ���ȣ"

        ' �̸���
        ' �˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
        ' ���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
        CashbillObj.email = ""

        ' �޴���
        CashbillObj.hp = ""

        ' ����ȳ����� ���ۿ���
        ' �ȳ����� ���۽� ����Ʈ�� �����Ǹ�, ���۽��н� ȯ��ó���˴ϴ�.
        CashbillObj.smssendYN = False

        ' �ŷ��Ͻ�, ��¥(yyyyMMddHHmmss)
        ' ����, ���ϸ� ����, ���Է½� �⺻�� �����Ͻ� ó��
        CashbillObj.tradeDT = "20221108000000"

        Set cashbillList(i) =  CashbillObj
    Next

    On Error Resume Next

    Set Presponse = m_CashbillService.BulkSubmit(testCorpNum, SubmitID, cashbillList, userID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        receiptID = ""
        Err.Clears
    Else
        code = Presponse.code
        message =Presponse.message
        receiptID = Presponse.receiptID
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>���ݿ����� �ʴ뷮 ����</legend>
                <ul>
                    <li>�����ڵ� (Response.code) : <%=code%> </li>
                    <li>����޽��� (Response.message) : <%=message%> </li>
                    <% If receiptID <> "" Then %>
                    <li>�������̵� (Response.receiptID) : <%=receiptID%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ۼ��� ���ݿ����� �����͸� �˺��� ����� ���ÿ� �����Ͽ� "����Ϸ�" ���·� ó���մϴ�.
    ' - ���ݿ����� ����û ���� ��å : https://developers.popbill.com/guide/cashbill/asp/introduction/policy-of-send-to-nts
    ' - https://developers.popbill.com/reference/cashbill/asp/api/issue#RegistIssue
    '**************************************************************


    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ȣ, ������ ����ڴ��� ������ȣ �Ҵ�, 1~24�ڸ� ����,������������ �ߺ����� ����.
    mgtKey = "20221109-ASP-001"

    ' �޸�
    memo = "��ù��� �޸�"

    ' �ȳ����� ����, ���� ����� �⺻������� ����
    emailSubject = "���� �ȳ� ���� ����"

    ' ���ݿ����� ��ü ����
    Set CashbillObj = New CashBill

    CashbillObj.mgtKey = mgtKey

    ' ��������, [���ΰŷ�] ����
    CashbillObj.tradeType = "���ΰŷ�"

    ' �ŷ�����, [�ҵ������, ����������] �� ����
    CashbillObj.tradeUsage = "�ҵ������"

    ' �ŷ�����, [�Ϲ�, ��������, ���߱���] �� ����
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
    CashbillObj.identityNum = "0101112222"

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

    On Error Resume Next

    Set Presponse = m_CashbillService.RegistIssue(testCorpNum, CashbillObj, memo, userID, emailSubject)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message = Presponse.message
        confirmNum = Presponse.confirmNum
        tradeDate = Presponse.tradeDate
        tradeDT = Presponse.tradeDT
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
                    <% If confirmNum <> "" Then %>
                    <li> Response.confirmNum : <%=confirmNum%> </li>
                    <% End If %>
                    <% If tradeDate <> "" Then %>
                    <li> Response.tradeDate : <%=tradeDate%> </li>
                    <% End If %>
                    <% If tradeDT <> "" Then %>
                    <li> Response.tradeDT : <%=tradeDT%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>

<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' �ִ� 100���� ���ݿ����� ������ �ѹ��� ��û���� �����մϴ�.
    ' - https://docs.popbill.com/cashbill/asp/api#BulkSubmit
    '**************************************************************
    
    ' �˺�ȸ�� ����ڹ�ȣ
    testCorpNum = "1234567890"

    ' ������̵�, �ִ� 36�ڸ� (����, ����, "-" ����)
    SubmitID = "ASPBULK001"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"
    
    Dim cashbillList(99)  
    for i = 0 to 99
        ' ���ݿ����� ���� ��ü ����
        Set CashbillObj = New Cashbill

        CashbillObj.mgtKey = SubmitID + CStr(i)

        '��������, [���ΰŷ�, ��Ұŷ�] �� ����
        CashbillObj.tradeType = "���ΰŷ�"

        '[��Ұŷ��� �ʼ�] ���� ���ݿ����� ����û���ι�ȣ
        CashbillObj.orgConfirmNum = ""

        '[��Ұŷ��� �ʼ�] ���� ���ݿ����� �ŷ�����
        CashbillObj.orgTradeDate = ""

        '�ŷ�����, [�ҵ������, ����������] �� ����
        CashbillObj.tradeUsage = "�ҵ������"

        '�ŷ�����, [�Ϲ�, ��������, ���߱���] �� ����
        CashbillObj.tradeOpt = "�Ϲ�"

        '��������, [����, �����] �� ����
        CashbillObj.taxationType = "����"

        '���ް���
        CashbillObj.supplyCost = "10000"

        '�ΰ���
        CashbillObj.tax = "1000"

        '�����
        CashbillObj.serviceFee = "0"

        '�հ�ݾ�, ���ް��� + ����� + ����
        CashbillObj.totalAmount = "11000"

        '������ ����ڹ�ȣ, "-" ���� 10�ڸ�
        CashbillObj.franchiseCorpNum = testCorpNum

        '������ ������� �ĺ���ȣ
        CashbillObj.franchiseTaxRegID = ""

        '������ ��ȣ
        CashbillObj.franchiseCorpName = "������ ��ȣ"

        '������ ��ǥ�� ����
        CashbillObj.franchiseCEOName = "������ ��ǥ��"

        '������ �ּ�
        CashbillObj.franchiseAddr = "������ �ּ�"

        '������ ��ȭ��ȣ
        CashbillObj.franchiseTEL = "070-1234-1234"

        '�ŷ�ó �ĺ���ȣ, �ŷ������� ���� �ۼ�
        '�ҵ������ - �ֹε��/�޴���/ī���ȣ ���簡��
        '���������� - ����ڹ�ȣ/�ֹε��/�޴���/ī���ȣ ���簡��
        CashbillObj.identityNum = "0101112222"

        '�ֹ�������
        CashbillObj.customerName = "������"

        '�ֹ���ǰ��
        CashbillObj.itemName = "��ǰ��"

        '�ֹ���ȣ
        CashbillObj.orderNumber = "�ֹ���ȣ"

        '�̸���
        '�˺� ����ȯ�濡�� �׽�Ʈ�ϴ� ��쿡�� �ȳ� ������ ���۵ǹǷ�,
        '���� �ŷ�ó�� �����ּҰ� ������� �ʵ��� ����
        CashbillObj.email = "test@test.com"

        '�޴���
        CashbillObj.hp = "111-1234-1234"

        '�ѽ�
        CashbillObj.fax = "777-444-3333"

        '����ȳ����� ���ۿ���
        '�ȳ����� ���۽� ����Ʈ�� �����Ǹ�, ���۽��н� ȯ��ó���˴ϴ�.
        CashbillObj.smssendYN = False
        
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
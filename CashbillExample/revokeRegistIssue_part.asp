<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ۼ��� (�κ�)��� ���ݿ����� �����͸� �˺��� ����� ���ÿ� �����Ͽ� "����Ϸ�" ���·� ó���մϴ�.
    ' - ��� ���ݿ������� �ݾ��� ���� �ݾ��� ���� �� �����ϴ�.
    ' - ���ݿ����� ����û ���� ��å [https://developers.popbill.com/guide/cashbill/asp/introduction/policy-of-send-to-nts]
    ' - ��� ���ݿ����� ���� �� ������ �����ּҷ� ���� �ȳ� ������ ���۵Ǵ� �����Ͻñ� �ٶ��ϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/issue#RevokeRegistIssue_Part
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ȣ, ������ ����ڹ�ȣ ���� ������ȣ �Ҵ�, 1~24�ڸ� ����,������������ �ߺ����� ����.
    mgtKey = "20220720-ASP-005"

    ' ���� ���ݿ����� ����û���ι�ȣ
    orgConfirmNum = "TB0000105"

    ' ���� ���ݿ����� �ŷ�����
    orgTradeDate = "20221108"

    ' �ȳ� ���� ���ۿ��� , true / false �� �� 1
    ' �� true = ���� , false = ������
    ' �� ���� ���ݿ������� ������(����)�� �޴�����ȣ ���� ����
    smssendYN = False

    ' �޸�
    memo = "��ù��� �޸�"

    ' ���ݿ����� ������� , true / false �� �� 1
    ' �� true = �κ� ��� , false = ��ü ���
    ' �� ���Է½� �⺻�� false ó��
    isPartCancel = True

    ' ��һ��� , 1 / 2 / 3 �� �� 1
    ' �� 1 = �ŷ���� , 2 = �����߱���� , 3 = ��Ÿ
    ' �� ���Է½� �⺻�� 1 ó��
    cancelType = 1

    ' [���] ���ް���
    ' - ���ݿ����� ��������� true �� ��� ����� ���ް��� �Է�
    ' - ���ݿ����� ��������� false �� ��� ���Է�
    supplyCost = "5000"

    ' [���] �ΰ���
    ' - ���ݿ����� ��������� true �� ��� ����� �ΰ��� �Է�
    ' - ���ݿ����� ��������� false �� ��� ���Է�
    tax = "500"

    ' [���] �����
    ' - ���ݿ����� ��������� true �� ��� ����� ����� �Է�
    ' - ���ݿ����� ��������� false �� ��� ���Է�
    serviceFee = "0"

    ' [���] �ŷ��ݾ� (���ް���+�ΰ���+�����)
    ' - ���ݿ����� ��������� true �� ��� ����� �ŷ��ݾ� �Է�
    ' - ���ݿ����� ��������� false �� ��� ���Է�
    totalAmount = "5500"

    ' �ȳ����� ����, ����ó���� �⺻������� ����
    emailSubject = ""

    ' �ŷ��Ͻ�, ��¥(yyyyMMddHHmmss)
    ' ����, ���ϸ� ����, ���Է½� �⺻�� �����Ͻ� ó��
    tradeDT = ""

    On Error Resume Next

    Set Presponse = m_CashbillService.RevokeRegistIssue_Part(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID, isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount, emailSubject, tradeDT)

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
                <legend>(�κ�) ������ݿ����� ��ù���</legend>
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
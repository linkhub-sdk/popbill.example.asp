<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' �ۼ��� (�κ�)��� ���ݿ����� �����͸� �˺��� ����� ���ÿ� �����Ͽ� "����Ϸ�" ���·� ó���մϴ�.
    ' - ��� ���ݿ������� �ݾ��� ���� �ݾ��� ���� �� �����ϴ�.
    ' - ���ݿ����� ����û ���� ��å [https://docs.popbill.com/cashbill/ntsSendPolicy?lang=asp]
    ' - "����Ϸ�"�� ��� ���ݿ������� ����û ���� ������ �������(cancelIssue API) �Լ��� ����û �Ű� ��󿡼� ������ �� �ֽ��ϴ�.
    ' - ��� ���ݿ����� ���� �� ������ �����ּҷ� ���� �ȳ� ������ ���۵Ǵ� �����Ͻñ� �ٶ��ϴ�.
    ' - https://docs.popbill.com/cashbill/asp/api#RevokeRegistIssue_Part
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"	

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"				 

    ' ������ȣ, ������ ����ڹ�ȣ ���� ������ȣ �Ҵ�, 1~24�ڸ� ����,������������ �ߺ����� ����.
    mgtKey = "20171115-01"

    ' ���� ���ݿ����� ����û���ι�ȣ
    orgConfirmNum = "820116333"

    ' ���� ���ݿ����� �ŷ�����
    orgTradeDate = "20170711"

    ' �ȳ� ���� ���ۿ��� , true / false �� �� 1
    ' �� true = ���� , false = ������
    ' �� ���� ���ݿ������� ������(��)�� �޴�����ȣ ���� ����
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

    On Error Resume Next

    Set Presponse = m_CashbillService.RevokeRegistIssue_Part(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID, _
        isPartCancel, cancelType, supplyCost, tax, serviceFee, totalAmount)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else 
        code = Presponse.code
        message = Presponse.message
        confirmNum = Presponse.confirmNum
        tradeDate = Presponse.tradeDate		
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
                </ul>
            </fieldset>
         </div>
    </body>
</html>
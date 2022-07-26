<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' ��� ���ݿ����� �����͸� �˺��� ����� ���ÿ� �����Ͽ� "����Ϸ�" ���·� ó���մϴ�.
    ' - ��� ���ݿ������� �ݾ��� ���� �ݾ��� ���� �� �����ϴ�.
    ' - ���ݿ����� ����û ���� ��å [https://docs.popbill.com/cashbill/ntsSendPolicy?lang=asp]
    ' - "����Ϸ�"�� ��� ���ݿ������� ����û ���� ������ �������(cancelIssue API) �Լ��� ����û �Ű� ��󿡼� ������ �� �ֽ��ϴ�.
    ' - ��� ���ݿ����� ���� �� ������ �����ּҷ� ���� �ȳ� ������ ���۵Ǵ� �����Ͻñ� �ٶ��ϴ�.
    ' - https://docs.popbill.com/cashbill/asp/api#RevokeRegistIssue
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"	

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"				 

    ' ������ȣ, ������ ����ڹ�ȣ ���� ������ȣ �Ҵ�, 1~24�ڸ� ����,������������ �ߺ����� ����.
    mgtKey = "20220720-ASP-002"

    ' ���� ���ݿ����� ����û���ι�ȣ
    orgConfirmNum = "TB0000067"

    ' ���� ���ݿ����� �ŷ�����
    orgTradeDate = "20220720"

    ' ����ȳ� ���� ���ۿ���
    smssendYN = False

    ' �޸�
    memo = "��ù��� �޸�"

    On Error Resume Next

    Set Presponse = m_CashbillService.RevokeRegistIssue(testCorpNum, mgtKey, orgConfirmNum, orgTradeDate, smssendYN, memo, userID)

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
                <legend>������ݿ����� ��ù���</legend>
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
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ڼ��ݰ�꼭 1���� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�ϸ�, ������������ �μ� �������� "������" / "���޹޴���" / "������+���޹޴���"�� �� �ϳ��� ������ �� �ֽ��ϴ�.
    ' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/view#GetPrintURL
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ���ݰ�꼭 �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType= "SELL"

    ' ������ȣ
    MgtKey = "20220720-ASP-001"

    On Error Resume Next

    url = m_TaxinvoiceService.GetPrintURL(testCorpNum, KeyType, MgtKey, userID)

    If Err.Number <> 0 then
        code = Err.Number
        message =  Err.Description
        Err.Clears
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>���ݰ�꼭 �μ� URL</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>URL : <%=url%> </li>
                    </ul>
                <% Else %>
                    <ul>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    </ul>
                <% End If %>
            </fieldset>
         </div>
    </body>
</html>
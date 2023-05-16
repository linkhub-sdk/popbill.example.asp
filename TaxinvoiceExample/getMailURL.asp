<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ڼ��ݰ�꼭 �ȳ������� �󼼺��� ��ũ URL�� ��ȯ�մϴ�.
    ' �Լ� ȣ��� ��ȯ ���� URL���� ��ȿ�ð��� �����ϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/view#GetMailURL
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType= "SELL"

    ' ������ȣ
    MgtKey = "20220720-ASP-001"

    On Error Resume Next

    url = m_TaxinvoiceService.GetMailURL(testCorpNum, KeyType, MgtKey, userID)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>���ݰ�꼭 ���ϸ�ũ URL</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>URL : <%=url%> </li>
                    <% Else %>
                        <li> Response.code : <%=code%> </li>
                        <li> Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
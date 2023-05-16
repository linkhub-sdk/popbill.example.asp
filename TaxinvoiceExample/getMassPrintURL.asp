<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ټ����� ���ݰ�꼭�� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�մϴ�. (�ִ� 100��)
    ' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/view#GetMassPrintURL
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType= "SELL"

    ' �μ��� ���ݰ�꼭 ������ȣ �迭, �ִ� 100��
    Dim MgtKeyList(3)
    MgtKeyList(0) = "20220720-ASP-001"
    MgtKeyList(1) = "20220720-ASP-002"
    MgtKeyList(2) = "20220720-ASP-003"

    On Error Resume Next

    url = m_TaxinvoiceService.GetMassPrintURL(testCorpNum, KeyType, MgtKeyList, userID)

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
                <legend>���ݰ�꼭 �μ� URL - �뷮</legend>
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
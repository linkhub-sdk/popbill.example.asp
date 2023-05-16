<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ټ����� ���ݿ������� �μ��ϱ� ���� �������� �˾� URL�� ��ȯ�մϴ�. (�ִ� 100��)
    ' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/view#GetMassPrintURL
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ���ݿ����� ������ȣ, �ִ� 100��
    Dim mgtKeyList(2)
    mgtKeyList(0) = "20220720-ASP-001"
    mgtKeyList(1) = "20220720-ASP-002"

    On Error Resume Next

    url = m_CashbillService.GetMassPrintURL(testCorpNum, mgtKeyList, userID)

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
                <legend>���ݿ����� �μ� �˾� URL - �뷮</legend>
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
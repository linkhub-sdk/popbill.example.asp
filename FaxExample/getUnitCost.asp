<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ѽ� ���۽� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/fax/asp/api/point#GetUnitCost
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' ���Ź�ȣ ���� : "�Ϲ�" / "����" �� �� 1
    ' �� �Ϲݸ� : ���ɸ��� ������ ��ȣ
    ' �� ���ɸ� : 030*, 050*, 070*, 080*, ��ǥ��ȣ
    receiveNumType = "����"

    On Error Resume Next

    unitCost = m_FaxService.GetUnitCost(testCorpNum, receiveNumType)

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
                <legend>�ѽ� ���� �ܰ� Ȯ�� </legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>���� �ܰ� : <%=unitCost%> </li>
                    <% Else %>
                        <li> Response.code : <%=code%></li>
                        <li> Response.message : <%=message%></li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
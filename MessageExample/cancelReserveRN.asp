<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� ���������� ���� ������ ����մϴ�. (����ð� 10�� ������ ����)
    ' - https://developers.popbill.com/reference/sms/asp/api/send#CancelReserveRN
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ���๮�� ���ۿ�û�� ��Ʈ�ʰ� �Ҵ��� ���ۿ�û��ȣ
    requestNum = "20220720-ASP-001"

    On Error Resume Next

    Set result = m_MessageService.CancelReserveRN(testCorpNum, requestNum, userID)

    If Err.Number <> 0 then
        code = Err.Number
        message =  Err.Description
        Err.Clears
    Else
        code = result.code
        message = result.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>���ڿ������� ���</legend>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
            </fieldset>
         </div>
    </body>
</html>
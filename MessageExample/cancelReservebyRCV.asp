<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺����� ��ȯ���� ������ȣ�� ���Ź�ȣ�� ���� ���������� ���� �޽��� ������ ����մϴ�. (����ð� 10�� ������ ����)
    ' - https://developers.popbill.com/reference/sms/asp/api/send#CancelReservebyRCV
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' ���๮�� ���ۿ�û�� �˺��κ��� ��ȯ ���� ������ȣ
    ReceiptNum = "022102708000000003"

    ' ���๮�� ���ۿ�û�� �˺��� ��û�� ���Ź�ȣ
    ReceiveNum = "0102223333"

    On Error Resume Next

    Set result = m_MessageService.CancelReservebyRCV(CorpNum, ReceiptNum, ReceiveNum, UserID)

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

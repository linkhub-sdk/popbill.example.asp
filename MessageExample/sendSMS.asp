<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    '  �ִ� 90byte�� �ܹ�(SMS) �޽��� 1�� ������ �˺��� �����մϴ�.
    ' - https://developers.popbill.com/reference/sms/asp/api/send#SendSMS
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ �޽��� ���� ( true , false �� �� 1)
    ' �� true = ���� , false = �Ϲ�
    adsYN = False

    ' �������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
    reserveDT = ""

    Set msgList = CreateObject("Scripting.Dictionary")

    ' �޽��� ��������
    Set message = New Messages

    ' �߽��ڹ�ȣ
    message.sender = ""

    ' �߽��ڸ�
    message.senderName = "�߽��ڸ�"

    ' �����ڹ�ȣ
    message.receiver = ""

    ' �����ڸ�
    message.receivername = " �������̸�"

    ' �޽��� ����, 90byte �ʰ��ϴ� ��� ���̰� �����Ǿ� ���۵�
    message.content = "�˺� �޽��� API �׽�Ʈ��"

    msgList.Add 0, message

    ' ���ۿ�û��ȣ
    ' �˺��� ���� ������ �ĺ��� �� �ֵ��� ��Ʈ�ʰ� �Ҵ��� �ĺ���ȣ.
    ' 1~36�ڸ��� ����. ����, ����, ������(-), �����(_)�� �����Ͽ� �˺� ȸ������ �ߺ����� �ʵ��� �Ҵ�.
    requestNum = ""

    On Error Resume Next

    receiptNum = m_MessageService.SendSMS(testCorpNum, "", "", msgList, reserveDT, adsYN, requestNum, userID)

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
                <legend>�ܹ� ���ڸ޽��� 1�� ���� </legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>ReceiptNum(������ȣ) : <%=receiptNum%> </li>
                    </ul>
                <%	Else  %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <%	End If	%>
            </fieldset>
         </div>
    </body>
</html>
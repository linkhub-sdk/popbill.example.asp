<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ִ� 2,000byte�� �޽����� �̹����� ������ ���乮��(MMS) 1�� ������ �˺��� �����մϴ�.
    ' - �̹��� ���� ����/�԰� : �ִ� 300Kbyte(JPEG), ����/���� 1,000px ���� ����
    ' - https://developers.popbill.com/reference/sms/asp/api/send#SendMMS
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

    ' ������������
    Set message = New Messages

    ' �߽Ź�ȣ
    message.sender = ""

    ' �߽��ڸ�
    message.senderName = "�߽��ڸ�"

    ' ���Ź�ȣ
    message.receiver = ""

    ' �����ڸ�
    message.receivername = " �������̸�"

    ' �޽��� ����, 2000byte �ʰ��� ���̰� �����Ǿ����۵�.
    message.content = "MMS �޽��� �׽�Ʈ��"

    ' �޽��� ����
    message.subject = "MMS �޽��� �����Դϴ�"

    msgList.Add 0, message

    ' ���� �޽��� ÷������, 300KByte���� JPEG ���� ���۰���
    FilePaths = Array("C:\popbill.example.asp\test.jpg")

    ' ���ۿ�û��ȣ
    ' �˺��� ���� ������ �ĺ��� �� �ֵ��� ��Ʈ�ʰ� �Ҵ��� �ĺ���ȣ.
    ' 1~36�ڸ��� ����. ����, ����, ������(-), �����(_)�� �����Ͽ� �˺� ȸ������ �ߺ����� �ʵ��� �Ҵ�.
    requestNum = ""

    On Error Resume Next

    receiptNum = m_MessageService.SendMMS(testCorpNum, "", "", "", msgList, FilePaths, reserveDT, adsYN, requestNum, userID)

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
                <legend>MMS ���ڸ޽��� 1�� ���� </legend>
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
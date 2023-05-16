<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݰ�꼭�� ���õ� �ȳ� SMS(�ܹ�) ���ڸ� �������ϴ� �Լ���, �˺� ����Ʈ [���ڡ��ѽ�] > [����] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
    ' - �޽����� �ִ� 90byte���� �Է� �����ϰ�, �ʰ��� ������ �ڵ����� �����Ǿ� �����մϴ�. (�ѱ� �ִ� 45��)\
    ' - �Լ� ȣ��� ����Ʈ�� ���ݵ˴ϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#SendSMS
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    testUserID = "testkorea"

    ' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType = "SELL"

    ' ������ȣ
    MgtKey = "20220720-ASP-001"

    ' �߽Ź�ȣ
    Sender = ""

    ' ���Ź�ȣ
    Receiver = ""

    ' �޽��� ����, 90byte�ʰ��� ���̰� �����Ǿ� ���۵�
    Contents = "���� �׽�Ʈ�Դϴ� 90Bytes�� �ʰ��ѳ����� ���۵��� �ʽ��ϴ�"

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.SendSMS(testCorpNum, KeyType, MgtKey, Sender, Receiver, Contents, testUserID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message = Presponse.message
    End If
    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>�˸����� ����</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message : <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
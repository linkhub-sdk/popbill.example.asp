<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' �ִ� 90byte�� �ܹ�(SMS) �޽��� 1�� ������ �˺��� �����ϸ�, ������ ���� ���� ������ �����մϴ�. (�ִ� 1,000��)
    '  - �޽��� ������ 90Byte �ʰ��� �޽��� ������ �ڵ����� ���ŵ˴ϴ�.
    '  - https://docs.popbill.com/message/asp/api#SendSMS
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"		

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"					

    ' ���� �޽��� ���� ( true , false �� �� 1)
    ' �� true = ���� , false = �Ϲ�
    adsYN = False							

    ' �������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
    reserveDT = ""
    
    ' �������� �迭, �ִ� 1000��
    Set msgList = CreateObject("Scripting.Dictionary")

    For i=0 To 99
        Set message = New Messages

        ' �߽Ź�ȣ
        message.sender = ""

        ' �߽��ڸ�
        message.senderName = "�߽��ڸ�"

        ' ���Ź�ȣ
        message.receiver = ""

        ' �����ڸ�
        message.receivername = " �������̸�"+CStr(i)

        ' �޽�������, �ִ� 90byte�ʰ��� ���̰� �����Ǿ� ���۵�
        message.content = "This is Message �޽��� �׽�Ʈ��"

        msgList.Add i, message
    Next
    
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
                <legend>�ܹ� ���ڸ޽��� 100�� ���� </legend>
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
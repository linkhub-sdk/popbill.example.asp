<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' XMS(�ܹ�/�幮 �ڵ��ν�)�� �����մϴ�.
    ' - �޽��� ������ ����(90byte)�� ���� SMS/LMS(�ܹ�/�幮)�� �ڵ��ν��Ͽ� �����մϴ�.
    ' - 90byte �ʰ��� LMS(�幮)���� �ν� �մϴ�.
    ' - https://docs.popbill.com/message/asp/api#SendXMS
    '**************************************************************

    '�˺� ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    '�˺� ȸ�� ���̵�
    userID = "testkorea"

    '������ ���ۿ���
    adsYN = False

    '�������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
    reserveDT = ""
    
    Set msgList = CreateObject("Scripting.Dictionary")

    '���ڸ޽��� ��������
    Set message = New Messages

    '�߽Ź�ȣ
    message.sender = "07043042991"

    '�߽��ڸ�
    message.senderName = "�߽��ڸ�"

    '���Ź�ȣ
    message.receiver = "000111222"

    '�����ڸ�
    message.receivername = "�������̸�"

    '�޽�������, 90byte �������� ��/�幮�� �ڵ����� �νĵǾ� ����
    message.content = "��/�幮 �޽��� �ڵ��ν����� �׽�Ʈ�Դϴ�. �����ϴ� �޽����� ���̰� 90byte�̻��� ��� �幮(LMS)Ÿ������ �޽����� ���۵˴ϴ�. �������� �׽�Ʈ�Դϴ�."
    
    msgList.Add 0, message

    '���ۿ�û��ȣ (�˺� ȸ���� ���ߺ� ��ȣ �Ҵ�)
    '����,����,'-','_' ����, �ִ� 36��
    requestNum = ""	

    On Error Resume Next

    receiptNum = m_MessageService.SendXMS(testCorpNum, "", "", "", msgList, reserveDT, adsYN, requestNum, userID)

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
                <legend>��/�幮 �ڵ��ν����� 1�� ���� </legend>
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
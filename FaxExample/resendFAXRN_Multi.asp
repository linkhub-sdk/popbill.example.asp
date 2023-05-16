<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û��ȣ�� ���� �ټ��� �����ڿ��� �ѽ��� �������մϴ�. (�ִ� 1,000��)
    ' - �߽�/���� ���� ���Է½� ������ ������ ������ �ѽ��� ���۵ǰ�, ������ ���� �ִ� 60���� ������� �ʴ� �Ǹ� �������� �����մϴ�.
    ' - �ѽ� ������ ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
    ' - ��ȯ���� ������ ���۽����� �ѽ� �������� �������� �Ұ��մϴ�.
    ' - https://developers.popbill.com/reference/fax/asp/api/send#ResendFAXRN
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ���� �ѽ� ���۽� �Ҵ��� ���ۿ�û��ȣ(requestNum)
    orgRequestNum = "1"

    ' �߽��� ��ȣ
    sendNum = "07043042991"

    ' �߽��ڸ�
    sendName = "�߽��ڸ�"

    ' ���ۿ���ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
    reserveDT = ""

    ' �ѽ�����
    title = "�ѽ� ���� ������"

    ' ���������� �������������� ������ ���
    'ReDim receivers(-1)


    ' ���������� ������������ �ٸ� ��� �Ʒ� �ڵ� ����
    Dim receivers(1)
    Set receivers(0) = New FaxReceiver
    ' �ѽ� ���Ź�ȣ
    receivers(0).receiverNum = "010111222"
    ' �ѽ� �����ڸ�
    receivers(0).receiverName = "������ ��Ī"
    ' ��Ʈ�� ����Ű, ������ ������ �޸�
    receivers(0).interOPRefKey = "20220720-001"

    Set receivers(1) = New FaxReceiver
    ' �ѽ� ���Ź�ȣ
    receivers(1).receiverNum = "010111222"
    ' �ѽ� �����ڸ�
    receivers(1).receiverName = "������ ��Ī"
    ' ��Ʈ�� ����Ű, ������ ������ �޸�
    receivers(1).interOPRefKey = "20220720-002"


    ' ������ �ѽ��� ���ۿ�û��ȣ
    ' ��Ʈ�ʰ� ���� �ǿ� ���� ������ȣ�� �����Ͽ� �����ϴ� ��� ���.
    ' 1~36�ڸ��� ����. ����, ����, ������(-), �����(_)�� �����Ͽ� �˺� ȸ������ �ߺ����� �ʵ��� �Ҵ�.
    requestNum = ""

    On Error Resume Next

    url = m_FaxService.ResendFAXRN(testCorpNum, orgRequestNum, sendNum, sendName, receivers, reserveDT, userID, title, requestNum)

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
                <legend>�ѽ� ������</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>recepitNum : <%=url%> </li>
                    <% Else %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
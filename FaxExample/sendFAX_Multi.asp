<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%

    '**************************************************************
    ' ������ �ѽ������� �ټ��� �����ڿ��� �����ϱ� ���� �˺��� �����մϴ�. (�ִ� �������� ���� : 20��) (�ִ� 1,000��)
    ' - https://developers.popbill.com/reference/fax/asp/api/send#SendFAX
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' �߽Ź�ȣ
    sendNum = ""

    ' ���ۿ���ð� yyyyMMddHHmmss, reserveDT���� null ��� �������
    reserveDT = ""

    ' �������� �迭 �ִ� 1000��
    Dim receivers(1)
    Set receivers(0) = New FaxReceiver
    ' �ѽ� ���Ź�ȣ
    receivers(0).receiverNum = ""
    ' �ѽ� �����ڸ�
    receivers(0).receiverName = "������ ��Ī"
    ' ��Ʈ�� ����Ű, ������ ������ �޸�
    receivers(0).interOPRefKey = "20220720-001"

    Set receivers(1) = New FaxReceiver
    ' �ѽ� ���Ź�ȣ
    receivers(1).receiverNum = ""
    ' �ѽ� �����ڸ�
    receivers(1).receiverName = "������ ��Ī"
    ' ��Ʈ�� ����Ű, ������ ������ �޸�
    receivers(1).interOPRefKey = "20220720-002"

    ' �ѽ������� ���� (�ִ� 20��)
    FilePaths = Array("C:\popbill.example.asp\���ѹα����.doc","C:\popbill.example.asp\test.jpg")

    ' �����ѽ� ���ۿ��� , true / false �� �� 1
    ' �� true = ���� , false = �Ϲ�
    ' �� ���Է� �� �⺻�� false ó��
    adsYN = False

    ' �ѽ�����
    title = "�ѽ� �������� ����"

    ' ���ۿ�û��ȣ
    ' ��Ʈ�ʰ� ���� �ǿ� ���� ������ȣ�� �����Ͽ� �����ϴ� ��� ���.
    ' 1~36�ڸ��� ����. ����, ����, ������(-), �����(_)�� �����Ͽ� �˺� ȸ������ �ߺ����� �ʵ��� �Ҵ�.
    RequestNum = ""

    On Error Resume Next

    url = m_FaxService.SendFAX(CorpNum, sendNum, receivers, FilePaths, reserveDT, UserID, adsYN, title, RequestNum)

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
                <legend>�ѽ� ����</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>recepitNum (������ȣ) : <%=url%> </li>
                    <% Else %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>

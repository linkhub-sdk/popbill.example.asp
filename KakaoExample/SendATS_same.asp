<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ε� ���ø� ������ �ۼ��Ͽ� �ټ����� �˸��� ������ �˺��� �����ϸ�, ��� �����ڿ��� ���� ������ �����մϴ�. (�ִ� 1,000��)
    ' - ������ ���ε� ���ø��� ����� �˸��� ���۳���(content)�� �ٸ� ��� ���۽��� ó���˴ϴ�.
    ' - ���۽��н� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/send#SendATS
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    testUserID = "testkorea"

    ' ���ε� �˸��� ���ø��ڵ�
    ' �� �˸��� ���ø� ���� �˾� URL(GetATSTemplateMgtURL API) �Լ�, �˸��� ���ø� ��� Ȯ��(ListATStemplate API) �Լ��� ȣ���ϰų�
    '   �˺�����Ʈ���� ���ε� �˸��� ���ø� �ڵ带  Ȯ�� ����.
    templateCode = "019020000163"

    ' �˺��� ���� ��ϵ� �߽Ź�ȣ
    ' altSendType = 'C' / 'A' �� ���, ��ü���ڸ� ������ �߽Ź�ȣ
    ' altSendType = '' �� ���, null �Ǵ� ���� ó��
    ' �� ��ü���ڸ� �����ϴ� ��쿡�� ������ ��ϵ� �߽Ź�ȣ �Է� �ʼ�
    senderNum = ""

    ' �˸��� ����, �ִ� 1000��
    content = "[ �˺� ]" & vbCrLf
    content = content + "��û�Ͻ� #{���ø��ڵ�}�� ���� �ɻ簡 �Ϸ�Ǿ� ���� ó���Ǿ����ϴ�." & vbCrLf
    content = content + "�ش� ���ø����� ���� �����մϴ�." & vbCrLf & vbCrLf
    content = content + "���ǻ��� �����ø� ��Ʈ�ʼ��ͷ� ���ϰ� �����ֽñ� �ٶ��ϴ�. " & vbCrLf & vbCrLf
    content = content + "�˺� ��Ʈ�ʼ��� : 1600-8536" & vbCrLf
    content = content + "support@linkhub.co.kr"

    ' ��ü���� ����
    ' - �޽��� ����(90byte)�� ���� �幮(LMS)�� ��쿡�� ����.
    altSubject = "��ü���� ����"

    ' ��ü���� ����(altSendType)�� "A"�� ���, ��ü���ڷ� ������ ���� (�ִ� 2000byte)
    ' �� �˺��� �޽��� ���̿� ���� �ܹ�(90byte ����) �Ǵ� �幮(90byte �ʰ�)���� ����ó��
    altContent = "��ü���� �޽��� ����"

    ' ��ü���� ���� (null , "C" , "A" �� �� 1)
    ' null = ������, C = �˸���� ���� ���� ���� , A = ��ü���� ����(altContent)�� �Է��� ���� ����
    altSendType = "C"

    ' �������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
    reserveDT = ""

    Set receiverList = CreateObject("Scripting.Dictionary")

    ' �������� �迭, �ִ� 1000��
    For i =0 To 9
        Set rcvInfo = New KakaoReceiver

        ' �����ڹ�ȣ
        rcvInfo.rcv = "01011222"+ CStr(i)

        ' �����ڸ�
        rcvInfo.rcvnm = " �������̸�"

        ' ��Ʈ�� ����Ű, ������ ������ �޸�, �̻��� ����ó��
        rcvInfo.interOPRefKey = "20220720-" +CStr(i)

        receiverList.Add i, rcvInfo
    Next



    ' ���ۿ�û��ȣ
    ' �˺��� ���� ������ �ĺ��� �� �ֵ��� ��Ʈ�ʰ� �Ҵ��� �ĺ���ȣ.
    ' 1~36�ڸ��� ����. ����, ����, ������(-), �����(_)�� �����Ͽ� �˺� ȸ������ �ߺ����� �ʵ��� �Ҵ�.
    RequestNum = ""

    ' �˸��� ��ư������ ���ø� ��û�� ������ ��ư������ �����ϰ� �����ϴ� ��� btnList�� ���� �ϰ� �Լ�ȣ��.
    Set btnList = CreateObject("Scripting.Dictionary")

    ' �˸��� ��ư URL�� #{���ø�����}�� �����Ѱ�� ���ø����� ������ �����Ͽ� ��ư���� ����
    'Set btnInfo = New KakaoButton
    'btnInfo.n = "���ø� �ȳ�"
    'btnInfo.t = "WL"
    'btnInfo.u1 = "https://www.popbil.com"
    'btnInfo.u2 = "http://www.llinkhub.co.kr"
    'btnList.Add 0, btnInfo

    On Error Resume Next

    ReceiptNum = m_KakaoService.SendATS(CorpNum, templateCode, senderNum, content, altContent, altSendType, reserveDT, receiverList, RequestNum, testUserID, btnList, altSubject)

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
                <legend>�˸��� ���ϳ��� �뷮����</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>ReceiptNum(������ȣ) : <%=ReceiptNum%> </li>
                    </ul>
                <% Else %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <% End If %>
            </fieldset>
        </div>
    </body>
</html>

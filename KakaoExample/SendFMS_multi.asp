<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �̹����� ÷�ε� �ټ����� ģ���� ������ �˺��� �����ϸ�, ������ ���� ���� ������ �����մϴ�. (�ִ� 1,000��)
    ' - ģ������ ��� �߰� ������ ���ѵ˴ϴ�. (20:00 ~ ���� 08:00)
    ' - ���۽��н� ������ ������ ���� 'altSendType' ������ ��ü���ڸ� ������ �� �ְ�, �� ��� ����(SMS/LMS) ����� ���ݵ˴ϴ�.
    ' - ��ü������ ���, ���乮��(MMS) ������ �����ϰ� ���� �ʽ��ϴ�.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/send#SendFMS
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    testUserID = "testkorea"

    ' �˺��� ��ϵ� īī���� �˻��� ���̵�
    plusFriendID = "@�˺�"

    ' �˺��� ���� ��ϵ� �߽Ź�ȣ
    ' altSendType = 'C' / 'A' �� ���, ��ü���ڸ� ������ �߽Ź�ȣ
    ' altSendType = '' �� ���, null �Ǵ� ���� ó��
    ' �� ��ü���ڸ� �����ϴ� ��쿡�� ������ ��ϵ� �߽Ź�ȣ �Է� �ʼ�
    senderNum = ""

    ' ��ü���� ���� (null , "C" , "A" �� �� 1)
    ' null = ������, C = �˸���� ���� ���� ���� , A = ��ü���� ����(altContent)�� �Է��� ���� ����
    altSendType = "C"

    ' ��ü���� ����
    ' - �޽��� ����(90byte)�� ���� �幮(LMS)�� ��쿡�� ����.
    ' - �������� �迭�� ��ü���� ������ �Էµ��� ���� ��� ����.
    ' - ��� �����ڿ��� �ٸ� ������ ���� ��� 76�� ���ο� �ִ� altsjt �� �̿�.
    altSubject = "��ü���� ����"

    ' �������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
    reserveDT = ""

    ' ���� �޽��� ���� ( true , false �� �� 1)
    ' �� true = ���� , false = �Ϲ�
    ' - ���Է� �� �⺻�� false ó��
    adsYN = False

    ' ÷���̹��� ���� ���
    ' �̹��� ���� �԰�: ���� ���� - JPG ���� (.jpg, .jpeg), �뷮 - �ִ� 500 Kbyte, ũ�� - ���� 500px �̻�, ���� �������� ���� 0.5 ~ 1.3�� ���� ����
    filePaths = Array("C:\popbill.example.asp\test03.jpg")

    ' �̹��� ��ũ URL
    ' �� �����ڰ� ģ���� ��� �̹��� Ŭ���� ȣ��Ǵ� URL
    ' ���Է½� ÷�ε� �̹����� ��ũ ��� ���� ǥ��
    imageURL = "http://popbill.com"


    Set receiverList = CreateObject("Scripting.Dictionary")

    ' �������� �迭, �ִ� 1000��
    For i =0 To 9
        Set rcvInfo = New KakaoReceiver

        ' �����ڹ�ȣ
        rcvInfo.rcv = "01011222"+ CStr(i)

        ' �����ڸ�
        rcvInfo.rcvnm = " �������̸�"

        ' ģ���� ����, �ִ� 400��
        rcvInfo.msg = "ģ���� �޽��� ���� �����Դϴ�." +CStr(i)

        ' ��ü���� ����
        ' - �޽��� ����(90byte)�� ���� �幮(LMS)�� ��쿡�� ����.
        ' - ��� �����ڿ��� ������ ������ ���� ��� �迭�� ��� ���ҿ� ������ ���� �Է��ϰų�
        '   ���� �Է����� �ʰ� 37�� ���ο� �ִ� altSubject �� �̿�
        rcvInfo.altsjt = "��ü���� ����" +CStr(i)

        ' ��ü���� �޽��� ����
        rcvInfo.altmsg = "��ü���� �޽��� ����" +CStr(i)

        ' �����ں� ���� ��ư ���� ���۽� ���.
        ' �ִ� 5�� ��� ����.
        ' Set btnInfo = new KakaoButton
        ' btnInfo.n = "���ø� �ȳ�"
        ' btnInfo.t = "WL"
        ' btnInfo.u1 = "https://www.popbil.com" + Cstr(i)
        ' btnInfo.u2 = "http://www.llinkhub.co.kr"
        ' rcvInfo.AddBtn(btnInfo)

        ' Set btnInfo = new KakaoButton
        ' btnInfo.n = "���ø� �ȳ�"
        ' btnInfo.t = "WL"
        ' btnInfo.u1 = "https://www.TEST.com" + Cstr(i)
        ' btnInfo.u2 = "http://www.TEST.co.kr"
        ' rcvInfo.AddBtn(btnInfo)

        receiverList.Add i, rcvInfo
    Next


    'ģ���� ��ư���� ����
    '�����ں� ���� ��ư�� ����ϰų� ��ư�� ������� ������� btnList�� ���� �ϰ� �Լ�ȣ��.
    Set btnList = CreateObject("Scripting.Dictionary")
    Set btnInfo = New KakaoButton
    btnInfo.n = "��ư�̸�"
    btnInfo.t = "WL"
    btnInfo.u1 = "http://www.popbil.com"
    btnInfo.u2 = "http://www.llinkhub.co.kr"
    btnList.Add 0, btnInfo

    Set btnInfo = New KakaoButton
    btnInfo.n = "�޽��� ����"
    btnInfo.t = "MD"
    btnList.Add 1, btnInfo

    '���ۿ�û��ȣ
    ' ��Ʈ�ʰ� ���� �ǿ� ���� ������ȣ�� �����Ͽ� �����ϴ� ��� ���.
    ' 1~36�ڸ��� ����. ����, ����, ������(-), �����(_)�� �����Ͽ� �˺� ȸ������ �ߺ����� �ʵ��� �Ҵ�.
    RequestNum = ""

    On Error Resume Next

    ReceiptNum = m_KakaoService.SendFMS(CorpNum, plusFriendID, senderNum, "", "", _
        altSendType, reserveDT, adsYN, receiverList, btnList, filePaths, imageURL, RequestNum, testUserID, altSubject)

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
                <legend>ģ���� �������� �뷮 ����</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>ReceiptNum(������ȣ) : <%=ReceiptNum%> </li>
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

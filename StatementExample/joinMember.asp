<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' �˺� ����ȸ�� ������ ��û�մϴ�.
    ' - https://docs.popbill.com/statement/asp/api#JoinMember
    '**************************************************************

    ' ȸ������ ��ü ����
    Set joinInfo = New JoinForm

    '��ũ���̵� 
    joinInfo.LinkID = "TESTER"		   

    '����ڹ�ȣ, "-"���� 10�ڸ�
    joinInfo.CorpNum = "1234567890"    

    '��ǥ�ڼ���
    joinInfo.CEOName = "��ǥ�ڼ���"	
    
    '��ȣ��
    joinInfo.CorpName =  "��ȣ"	
    
    '�ּ�
    joinInfo.Addr =   "�ּ�"		   

    '����
    joinInfo.BizType =  "����"		   

    '����
    joinInfo.BizClass = "����"

    '���̵� (6�� �̻� 20�� �̸�)
    joinInfo.ID =  "userid"

    '��й�ȣ (8�� �̻� 20�� ����) ����, ���� ,Ư������ ����
    joinInfo.Password =  "asdf1234!@#$"

    '����ڸ�
    joinInfo.ContactName = "����ڸ�"    

    '����ڿ���ó
    joinInfo.ContactTEL = "02-999-9999"   

    '����� �޴�����ȣ
    joinInfo.ContactHP = "010-1234-5678"	

    '�ѽ���ȣ
    joinInfo.ContactFAX = "02-999-9999"		

    '����� �̸���
    joinInfo.ContactEmail = "test@test.com"

    On Error Resume Next

    Set Presponse = m_StatementService.JoinMember(joinInfo)
    
    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message =Presponse.message
    End If

    On Error GoTo 0

%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>����ȸ�� ����</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
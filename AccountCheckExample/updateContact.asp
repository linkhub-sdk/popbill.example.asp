<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' ����� ������ �����մϴ�.
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    CorpNum = "1234567890"		 

    ' �˺�ȸ�� ���̵� 
    UserID = "testkorea"				 

    ' ����� ���� ��ü ����
    Set contInfo = New ContactInfo
    
    ' ����� ���̵� 
    contInfo.id = UserID

    ' ����ڸ�
    contInfo.personName = "ASPTest"

    ' ����� ����ó
    contInfo.tel = "010-1234-1234"

    ' ����� �޴�����ȣ
    contInfo.hp = "010-1234-1234"

    ' ����� �̸����ּ�
    contInfo.email = "dev@linkhub.co.kr"

    ' ����� �ѽ���ȣ
    contInfo.fax = "02-6442-9700"

    ' ����� ��ȸ���� 1 - ���α��� / 2 - �б����  / 3 - ȸ�����
    contInfo.searchRole = 3

    On Error Resume Next

    Set Presponse = m_AccountCheckService.UpdateContact(CorpNum, contInfo, UserID)
    
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
                <legend>����� ��������</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
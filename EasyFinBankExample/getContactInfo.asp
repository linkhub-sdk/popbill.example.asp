<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ Ȯ���մϴ�.
	' - https://developers.popbill.com/reference/easyfinbank/asp/api/member#GetContactInfo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' Ȯ���� ����� ���̵�
    contactID = "testID"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    On Error Resume Next

    Set conInfo = m_EasyFinBankService.GetContactInfo(testCorpNum, contactID ,userID)

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
                <legend>����� ���� Ȯ��</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li> id(���̵�) : <%=conInfo.id%></li>
                        <li> personName(����� ����) : <%=conInfo.personName%></li>
                        <li> email(����� �̸���) : <%=conInfo.email%></li>
                        <li> tel(����� ����ó) : <%=conInfo.tel%></li>
                        <li> regDT(����Ͻ�) : <%=conInfo.regDT%></li>
                        <li> SearchRole(����� ��ȸ����) : <%=conInfo.SearchRole%></li>
                        <li> mgrYN(������ ����) : <%=conInfo.mgrYN%></li>
                        <li> state(����) : <%=conInfo.state%></li>
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
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/member#ListContact
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next
    Set result = m_EasyFinBankService.ListContact(testCorpNum, UserID)
    If Err.Number <> 0 Then
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
                <legend>����� ��� ��ȸ</legend>
                <%
                    If code = 0 Then
                        For i=0 To result.Count-1
                %>
                            <fieldset class="fieldset2">
                                <legend> ContactInfoList [ <%=i+1%> / <%=result.Count%> ] </legend>
                                    <ul>
                                        <li> id(���̵�) : <%=result.Item(i).id%></li>
                                        <li> personName(����� ����) : <%=result.Item(i).personName%></li>
                                        <li> email(����� �̸���) : <%=result.Item(i).email%></li>
                                        <li> tel(����� ����ó) : <%=result.Item(i).tel%></li>
                                        <li> regDT(����Ͻ�) : <%=result.Item(i).regDT%></li>
                                        <li> searchRole(����� ��ȸ����) : <%=result.Item(i).searchRole%></li>
                                        <li> mgrYN(������ ����) : <%=result.Item(i).mgrYN%></li>
                                        <li> state(����) : <%=result.Item(i).state%></li>
                                    </ul>
                                </fieldset>
                <%
                        Next
                    Else
                %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <%
                    End If
                %>
            </fieldset>
         </div>
    </body>
</html>
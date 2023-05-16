<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺��� ��ϵ� �������� ����� ��ȯ�մϴ�.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/manage#ListBankAccount
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_EasyFinBankService.ListBankAccount(testCorpNum, UserID)

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
                <legend>���� ���</legend>
                <%
                    If code = 0 Then
                        For i=0 To result.Count-1
                %>
                            <fieldset class="fieldset2">
                                <legend>ListBankAccount [ <%=i+1%> / <%=result.Count%> ] </legend>
                                    <ul>
                                        <li>accountNumber (���¹�ȣ) : <%=result.Item(i).accountNumber%></li>
                                        <li>bankCode (����ڵ�) : <%=result.Item(i).bankCode%></li>
                                        <li>accountName (���� ��Ī) : <%=result.Item(i).accountName%></li>
                                        <li>accountType (��������) : <%=result.Item(i).accountType%></li>
                                        <li>state (������ ����) : <%=result.Item(i).state%></li>
                                        <li>regDT (����Ͻ�) : <%=result.Item(i).regDT%></li>
                                        <li>contractDT (������ ���� �����Ͻ�) : <%=result.Item(i).contractDT %> </li>
                                        <li>useEndDate (������ ���� ������) : <%=result.Item(i).useEndDate %> </li>
                                        <li>baseDate (�ڵ����� ������) : <%=result.Item(i).baseDate %> </li>
                                        <li>contractState (������ ���� ����) : <%=result.Item(i).contractState%> </li>
                                        <li>closeRequestYN (������ ���� ������û ����) : <%=result.Item(i).closeRequestYN%> </li>
                                        <li>useRestrictYN (������ ���� ������� ����) : <%=result.Item(i).useRestrictYN%> </li>
                                        <li>closeOnExpired (������ ���� ���� �� ���� ����) : <%=result.Item(i).closeOnExpired %> </li>
                                        <li>unPaidYN (�̼��� ���� ����) : <%=result.Item(i).unPaidYN %> </li>
                                        <li>memo (�޸�) : <%=result.Item(i).memo%></li>

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

<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݿ����� ���� ���� �׸� ���� �߼ۼ����� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/etc#ListEmailConfig
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    '�˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set emailObj = m_CashbillService.listEmailConfig(testCorpNum, UserID)

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
                <legend>�˸����� ���۸�� ��ȸ</legend>
                        <ul>
                        <%
                            If code = 0 Then
                            For i=0 To emailObj.Count-1
                        %>
                            <% If emailObj.Item(i).emailType = "CSH_ISSUE" Then %>
                                    <li><%= emailObj.Item(i).emailType %> (�������� ���ݿ������� ���� �Ǿ����� �˷��ִ� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "CSH_CANCEL" Then %>
                                    <li><%= emailObj.Item(i).emailType %> (�������� ���ݿ������� ������� �Ǿ����� �˷��ִ� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                        <%
                            Next
                            Else
                        %>
                        </ul>
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

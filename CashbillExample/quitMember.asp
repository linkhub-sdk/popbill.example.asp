<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���Ե� ����ȸ���� Ż�� ��û�մϴ�.
    ' - ȸ��Ż�� ��û�� ���ÿ� �˺��� ��� ���� �̿��� �Ұ��ϸ�, �����ڸ� ������ ��� ����� ������ �ϰ�Ż�� �˴ϴ�.
    ' - ȸ��Ż��� ������ �����ʹ� ������ �Ұ����մϴ�.
    ' - ������ ������ ȸ��Ż�� �����մϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/member#QuitMember
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

	'Ż�� ����
	QuitReason = "Ż�����"

    '�˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result =m_CashbillService.QuitMember(CorpNum, QuitReason, UserID)

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
                <legend>�˺� ȸ�� Ż��</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                            <ul>
                                <li> code (���� �ڵ�) : <%=result.code%></li>
                                <li> message (���� �޽���) : <%=result.message%></li>
                            </ul>
                        </fieldset>
                <%
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

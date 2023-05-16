<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺��� ��ϵ� ���� ������ Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/manage#GetBankAccountInfo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' ����ڵ�
    ' �������-0002 / �������-0003 / ��������-0004 /��������-0007 / ��������-0011 / �츮����-0020
    ' SC����-0023 / �뱸����-0031 / �λ�����-0032 / ��������-0034 / ��������-0035 / ��������-0037
    ' �泲����-0039 / �������ݰ�-0045 / ��������-0048 / ��ü��-0071 / KEB�ϳ�����-0081 / ��������-0088 /��Ƽ����-0027
    BankCode = ""

    ' ���¹�ȣ ������('-') ����
    AccountNumber = ""

    On Error Resume Next
        Set result = m_EasyFinBankService.GetBankAccountInfo(testCorpNum, BankCode, AccountNumber, UserID)
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
                <legend>�������� ��ȸ</legend>
                <%
                    If code = 0 Then
                %>
                        <ul>

                            <li>accountNumber (���¹�ȣ) : <%=result.accountNumber%></li>
                            <li>bankCode (����ڵ�) : <%=result.bankCode%></li>
                            <li>accountName (���� ��Ī) : <%=result.accountName%></li>
                            <li>accountType (���� ����) : <%=result.accountType%></li>
                            <li>state (���� ����) : <%=result.state%></li>
                            <li>regDT (����Ͻ�) : <%=result.regDT%></li>
                            <li>contractDT (������ ���� �����Ͻ�) : <%=result.contractDT%></li>
                            <li>useEndDate (������ ���� ������) : <%=result.useEndDate%></li>
                            <li>baseDate (�ڵ����� ������) : <%=result.baseDate%></li>
                            <li>contractState (������ ���� ����) : <%=result.contractState%></li>
                            <li>closeRequestYN (������ ���� ������û ����) : <%=result.closeRequestYN%></li>
                            <li>useRestrictYN (������ ���� ������� ����) : <%=result.useRestrictYN%></li>
                            <li>closeOnExpired (������ ���� ���� �� ���� ����) : <%=result.closeOnExpired%></li>
                            <li>unPaidYN (�̼��� ���� ����) : <%=result.unPaidYN%></li>
                            <li>memo (�޸�) : <%=result.memo%></li>
                        </ul>
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
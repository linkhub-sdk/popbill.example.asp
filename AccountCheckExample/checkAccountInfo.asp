<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
        <title>��������ȸ API SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"-->
    <%
        '**************************************************************
        ' 1���� �����ּ����� ��ȸ�մϴ�.
        ' - https://developers.popbill.com/reference/accountcheck/asp/api/check#CheckAccountInfo
        '**************************************************************
        '�˺�ȸ�� ����ڹ�ȣ
        CorpNum = "1234567890"

        '�˺�ȸ�� ���̵�
        UserID = "testkorea"

        '����ڵ�
        BankCode = ""

        '���¹�ȣ
        AccountNumber = ""

        On Error Resume Next
            Set result = m_AccountCheckService.checkAccountInfo(CorpNum, BankCode, AccountNumber, UserID)

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
                <legend>���¼�����ȸ</legend>
            <%
                If Not IsEmpty(result) Then

            %>

                <ul>
                    <li>bankCode (����ڵ�) : <%= result.bankCode%></li>
                    <li>accountNumber (���¹�ȣ) : <%= result.accountNumber%></li>
                    <li>accountName (������ ����) : <%= result.accountName%></li>
                    <li>checkDate (Ȯ���Ͻ�) : <%= result.checkDate%></li>
                    <li>result (�����ڵ�) : <%= result.result%></li>
                    <li>resultMessage (����޽���) : <%= result.resultMessage%></li>
                </ul>

            <%
                End If
                If Not IsEmpty(code) then
            %>

            <ul>
                <li>Response.code : <%= code %> </li>
                <li>Response.message : <%= message %></li>
            </ul>
            <%
                End If
            %>

            </fieldset>
    </body>
</html>
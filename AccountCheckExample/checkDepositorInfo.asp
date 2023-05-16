<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
        <title>��������ȸ API SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"-->
    <%
        '**************************************************************
        ' 1���� �����ֽǸ��� ��ȸ�մϴ�.
        ' - https://developers.popbill.com/reference/accountcheck/asp/api/check#CheckDepositorInfo
        '**************************************************************
        '�˺�ȸ�� ����ڹ�ȣ
        CorpNum = "1234567890"

        '�˺�ȸ�� ���̵�
        UserID = "testkorea"

        '����ڵ�
        BankCode = ""

        '���¹�ȣ
        AccountNumber = ""

        ' ��Ϲ�ȣ ���� ( P / B �� �� 1 ,  P = ����, B = �����)
        identityNumType = ""

        ' ��Ϲ�ȣ
        ' - IdentityNumType ���� "B" �� ��� (������ '-' ����  ����ڹ�ȣ(10)�ڸ� �Է� )
        ' - IdentityNumType ���� "P" �� ��� (�������(6)�ڸ� �Է� (���� : YYMMDD))
        identityNum = ""

        On Error Resume Next
            Set result = m_AccountCheckService.checkDepositorInfo(CorpNum, BankCode, AccountNumber, identityNumType, identityNum, UserID)

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
                <legend>���½Ǹ���ȸ</legend>
            <%
                If Not IsEmpty(result) Then
            %>
                <ul>
                    <li>bankCode (����ڵ�) : <%= result.bankCode%></li>
                    <li>accountNumber (���¹�ȣ) : <%= result.accountNumber%></li>
                    <li>accountName (������ ����) : <%= result.accountName%></li>
                    <li>checkDate (Ȯ���Ͻ�) : <%= result.checkDate%></li>
                    <li>identityNumType (��Ϲ�ȣ ����) : <%= result.identityNumType%></li>
                    <li>identityNum (��Ϲ�ȣ) : <%= result.identityNum%></li>
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
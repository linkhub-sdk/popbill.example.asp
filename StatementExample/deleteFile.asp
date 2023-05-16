<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' "�ӽ�����" ������ ���ڸ������� ÷�ε� 1���� ������ �����մϴ�.
    ' - ���� �ĺ��� ���� ÷�ν� �ο��Ǵ� 'FileID'�� ÷������ ��� Ȯ��(GetFiles API) �Լ��� ȣ���Ͽ� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/statement/asp/api/etc#DeleteFile
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ �ڵ� - 121(�ŷ�������), 122(û����), 123(������), 124(���ּ�), 125(�Ա�ǥ), 126(������)
    itemCode = "121"

    ' ������ȣ
    mgtKey = "20220720-ASP-001"

    ' ���Ͼ��̵�, ÷�����ϸ��(getFiles) API�� AttachedFile��
    FileID = "2556D18D-9380-4843-B748-5B8120C31BA5.PBF"

    On Error Resume Next

    Set result = m_StatementService.DeleteFile(testCorpNum, itemCode, mgtKey, FileID, userID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = result.code
        message = result.message
    End If

    On Error GoTo 0

%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>÷������ ����</legend>
                    <ul>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
            </fieldset>
         </div>
    </body>
</html>
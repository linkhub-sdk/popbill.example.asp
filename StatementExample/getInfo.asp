<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ڸ������� 1���� ���� �� ������� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/statement/asp/api/info#GetInfo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ �ڵ� - 121(�ŷ�������), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
    itemCode = "121"

    ' ������ȣ
    mgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set result = m_StatementService.GetInfo(testCorpNum, itemCode, mgtKey, userID)

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
                <legend>���ڸ����� ����/��� ����Ȯ��</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li> itemKey(������Ű) : <%=result.itemKey %></li>
                        <li> itemCode(���������ڵ�) : <%=result.itemCode %></li>
                        <li> itemKey(�˺���ȣ) : <%=result.itemKey %></li>
                        <li> invoiceNum(�˺����ι�ȣ) : <%=result.invoiceNum %></li>
                        <li> mgtKey(��Ʈ�� ������ȣ) : <%=result.mgtKey %></li>
                        <li> taxType(��������) : <%=result.taxType %></li>
                        <li> writeDate(�ۼ�����) : <%=result.writeDate %></li>
                        <li> regDT(����Ͻ�) : <%=result.regDT %></li>
                        <li> senderCorpName(�߽��� ��ȣ) : <%=result.senderCorpName %></li>
                        <li> senderCorpNum(�߽��� ����ڹ�ȣ) : <%=result.senderCorpNum %></li>
                        <li> senderPrintYN(�߽��� �μ⿩��) : <%=result.senderPrintYN %></li>
                        <li> receiverCorpName(������ ��ȣ) : <%=result.receiverCorpName %></li>
                        <li> receiverCorpNum(������ ����ڹ�ȣ) : <%=result.receiverCorpNum %></li>
                        <li> receiverPrintYN(������ �μ⿩��) : <%=result.receiverPrintYN %></li>
                        <li> supplyCostTotal(���ް��� �հ�) : <%=result.supplyCostTotal %></li>
                        <li> taxTotal(���� �հ�) : <%=result.taxTotal %></li>
                        <li> purposeType(����/û��) : <%=result.purposeType %></li>
                        <li> issueDT(�����Ͻ�) : <%=result.issueDT %></li>
                        <li> stateCode(�����ڵ�) : <%=result.stateCode %></li>
                        <li> stateDT(���� �����Ͻ�) : <%=result.stateDT %></li>
                        <li> stateMemo(���¸޸�) : <%=result.stateMemo %></li>
                        <li> openYN(���� ���� ����) : <%=result.openYN %></li>
                        <li> openDT(���� �Ͻ�) : <%=result.openDT %></li>
                    <% Else %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ټ����� ���ڸ����� ���� �� ������� Ȯ���մϴ�. (1ȸ ȣ�� �� �ִ� 1,000�� Ȯ�� ����)
    ' - https://developers.popbill.com/reference/statement/asp/api/info#GetInfos
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ �ڵ� - 121(�ŷ�������), 122(û����), 123(������), 124(���ּ�), 125(�Ա�ǥ), 126(������)
    itemCode = "121"

    ' ������ȣ �迭, �ִ� 1000��
    Dim mgtKeyList(2)
    mgtKeyList(0) = "20220720-ASP-001"
    mgtKeyList(1) = "20220720-ASP-002"

    On Error Resume Next

    Set result = m_StatementService.GetInfos(testCorpNum, itemCode, mgtKeyList, userID)

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
                <legend>���ڸ����� ����/������� Ȯ�� - �뷮 </legend>
                <ul>
                    <% If code = 0 Then
                        For i=0 To result.Count-1 %>

                        <fieldset class="fieldset2">
                            <legend> ���ڸ����� ���� [<%=i+1%>] </legend>
                            <ul>
                                <li> itemCode(���������ڵ�) : <%=result.Item(i).itemCode %></li>
                                <li> itemKey(�˺���ȣ) : <%=result.Item(i).itemKey %></li>
                                <li> invoiceNum(�˺����ι�ȣ) : <%=result.Item(i).invoiceNum %></li>
                                <li> mgtKey(��Ʈ�� ������ȣ) : <%=result.Item(i).mgtKey %></li>
                                <li> taxType(��������) : <%=result.Item(i).taxType %></li>
                                <li> writeDate(�ۼ�����) : <%=result.Item(i).writeDate %></li>
                                <li> regDT(����Ͻ�) : <%=result.Item(i).regDT %></li>
                                <li> senderCorpName(�߽��� ��ȣ) : <%=result.Item(i).senderCorpName %></li>
                                <li> senderCorpNum(�߽��� ����ڹ�ȣ) : <%=result.Item(i).senderCorpNum %></li>
                                <li> senderPrintYN(�߽��� �μ⿩��) : <%=result.Item(i).senderPrintYN %></li>
                                <li> receiverCorpName(������ ��ȣ) : <%=result.Item(i).receiverCorpName %></li>
                                <li> receiverCorpNum(������ ����ڹ�ȣ) : <%=result.Item(i).receiverCorpNum %></li>
                                <li> receiverPrintYN(������ �μ⿩��) : <%=result.Item(i).receiverPrintYN %></li>
                                <li> supplyCostTotal(���ް��� �հ�) : <%=result.Item(i).supplyCostTotal %></li>
                                <li> taxTotal(���� �հ�) : <%=result.Item(i).taxTotal %></li>
                                <li> purposeType(����/û��) : <%=result.Item(i).purposeType %></li>
                                <li> issueDT(�����Ͻ�) : <%=result.Item(i).issueDT %></li>
                                <li> stateCode(�����ڵ�) : <%=result.Item(i).stateCode %></li>
                                <li> stateDT(���� �����Ͻ�) : <%=result.Item(i).stateDT %></li>
                                <li> stateMemo(���¸޸�) : <%=result.Item(i).stateMemo %></li>
                                <li> openYN(���� ���� ����) : <%=result.Item(i).openYN %></li>
                                <li> openDT(���� �Ͻ�) : <%=result.Item(i).openDT %></li>
                            </ul>
                        </fieldset>
                    <%
                        Next
                        Else
                    %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>
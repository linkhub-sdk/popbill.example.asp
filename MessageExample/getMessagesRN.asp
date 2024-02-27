<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� ���� ���ۻ��� �� ����� Ȯ���մϴ�.
    '  - https://developers.popbill.com/reference/sms/asp/api/info#GetMessagesRN
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' �������� ��û �� �Ҵ��� ���ۿ�û��ȣ(RequestNum)
    RequestNum = "20220720-ASP-001"

    On Error Resume Next

    Set result = m_MessageService.GetMessagesRN(CorpNum, RequestNum, UserID)

    If Err.Number <> 0 then
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
                <legend>���ڸ޽��� ���۰�� Ȯ��</legend>
                <ul>
                    <% If code = 0 Then
                        For i=0 To result.Count-1
                    %>
                        <fieldset class="fieldset2">
                            <legend>���ڸ޽��� ���۰�� [<%=i+1%>]</legend>
                            <ul>
                                <li>state (���ۻ��� �ڵ�) : <%=result.Item(i).state%> </li>
                                <li>result (���۰�� �ڵ�) : <%=result.Item(i).result%> </li>
                                <li>subject (�޽��� ����) : <%=result.Item(i).subject%> </li>
                                <li>content (�޽��� ����) : <%=result.Item(i).content%> </li>
                                <li>type (�޽��� ����) : <%=result.Item(i).msgType%> </li>
                                <li>sendnum (�߽Ź�ȣ) : <%=result.Item(i).sendnum%> </li>
                                <li>senderName (�߽��ڸ�) : <%=result.Item(i).senderName%> </li>
                                <li>ReceiveNum (���Ź�ȣ) : <%=result.Item(i).ReceiveNum%> </li>
                                <li>receiveName (�����ڸ�) : <%=result.Item(i).receiveName%> </li>
                                <li>receiptDT (�����Ͻ�) : <%=result.Item(i).receiptDT%> </li>
                                <li>sendDT (�����Ͻ�) : <%=result.Item(i).sendDT%> </li>
                                <li>resultDT (���۰�� �����Ͻ�) : <%=result.Item(i).resultDT%> </li>
                                <li>reserveDT (�����Ͻ�) : <%=result.Item(i).reserveDT%> </li>
                                <li>tranNet (����ó�� �̵���Ż��) : <%=result.Item(i).tranNet%> </li>
                                <li>ReceiptNum (������ȣ) : <%=result.Item(i).ReceiptNum%> </li>
                                <li>RequestNum (��û��ȣ) : <%=result.Item(i).RequestNum%> </li>
                                <li>interOPRefKey (��Ʈ�� ����Ű) : <%=result.Item(i).interOPRefKey%> </li>
                            </ul>
                        </fieldset>
                    <%
                        Next
                        Else
                    %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
        </div>
    </body>
</html>

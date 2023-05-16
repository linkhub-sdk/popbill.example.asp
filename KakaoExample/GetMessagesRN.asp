<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� �˸���/ģ���� ���ۻ��� �� ����� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/info#GetMessagesRN
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ���� ��û�� �Ҵ��� ���ۿ�û��ȣ(requestNum)
    requestNum = "20220720-0011"

    On Error Resume Next

    Set result = m_KakaoService.GetMessagesRN(testCorpNum, requestNum, UserID)

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
                <legend>īī���� ���۰�� Ȯ��</legend>
                    <%
                        If code = 0 Then
                    %>
                    <ul>
                        <li>contentType (īī���� ����) : <%=result.contentType%></li>
                        <li>templateCode (�˸��� ���ø� �ڵ�) : <%=result.templateCode%></li>
                        <li>plusFriendID (�÷���ģ�� ���̵�) : <%=result.plusFriendID%></li>
                        <li>sendNum (�߽Ź�ȣ) : <%=result.sendNum%></li>
                        <li>altContent (��ü���� ����) : <%=result.altContent%></li>
                        <li>altSendType (��ü���� ����) : <%=result.altSendType%></li>
                        <li>reserveDT (�����Ͻ�) : <%=result.reserveDT%></li>
                        <li>adsYN (�������� ����) : <%=result.adsYN%></li>
                        <li>imageURL (ģ���� �̹��� URL) : <%=result.imageURL%></li>
                        <li>sendCnt (���۰Ǽ�) : <%=result.sendCnt%></li>
                        <li>successCnt (�����Ǽ�) : <%=result.successCnt%></li>
                        <li>failCnt (���аǼ�) : <%=result.failCnt%></li>
                        <li>altCnt (��ü���� �Ǽ�) : <%=result.altCnt%></li>
                        <li>cancelCnt (��ҰǼ�) : <%=result.cancelCnt%></li>
                    </ul>
                    <%
                        For i=0 To Ubound(result.btns)-1
                    %>
                        <fieldset class="fieldset2">
                            <legend>��ư���� [<%=i+1%>]</legend>
                            <ul>
                                <li>n (��ư��) : <%=result.btns(i).n%> </li>
                                <li>t (��ư����) : <%=result.btns(i).t%> </li>
                                <li>u1 (��ư��ũ1) : <%=result.btns(i).u1%> </li>
                                <li>u2 (��ư��ũ2) : <%=result.btns(i).u2%> </li>
                            </ul>
                        </fieldset>
                    <%
                        Next
                    %>
                    <fieldset class="fieldset2">
                        <legend>���۰�� ���� ���</legend>
                    <%
                        For i=0 To UBound(result.msgs) -1
                    %>
                        <fieldset class="fieldset3">
                            <legend>���۰�� ���� [<%=i+1%>]</legend>
                            <ul>
                                <li>state (���ۻ��� �ڵ�) : <%=result.msgs(i).state%> </li>
                                <li>sendDT (�����Ͻ�) : <%=result.msgs(i).sendDT%> </li>
                                <li>receiveNum (���Ź�ȣ) : <%=result.msgs(i).receiveNum%> </li>
                                <li>receiveName (�����ڸ�) : <%=result.msgs(i).receiveName%> </li>
                                <li>content (�˸���/ģ���� ����) : <%=result.msgs(i).content%> </li>
                                <li>result (�˸���/ģ���� ���۰�� �ڵ�) : <%=result.msgs(i).result%> </li>
                                <li>resultDT (�˸���/ģ���� ���۰�� �����Ͻ�) : <%=result.msgs(i).resultDT%> </li>
                                <li>altContent (��ü���� ����) : <%=result.msgs(i).altContent%> </li>
                                <li>altContentType (��ü���� ��������) : <%=result.msgs(i).altContentType%> </li>
                                <li>altSendDT (��ü���� �����Ͻ�) : <%=result.msgs(i).altSendDT%> </li>
                                <li>altResult (��ü���� ���۰�� �ڵ�) : <%=result.msgs(i).altResult%> </li>
                                <li>altResultDT (��ü���� ���۰�� �����Ͻ�) : <%=result.msgs(i).altResultDT%> </li>
                                <li>receiptNum (������ȣ) : <%=result.msgs(i).receiptNum%> </li>
                                <li>requestNum (��û��ȣ) : <%=result.msgs(i).requestNum%> </li>
                                <li>interOPRefKey (��Ʈ�� ����Ű) : <%=result.msgs(i).interOPRefKey%> </li>
                            </ul>
                        </fieldset>
                    <%
                        Next
                    %>

                    <%
                        Else
                    %>
                        <ul>
                            <li>Response.code : <%=code%> </li>
                            <li>Response.message : <%=message%> </li>
                        </ul>
                    <% End If %>
            </fieldset>
         </div>
    </body>
</html>
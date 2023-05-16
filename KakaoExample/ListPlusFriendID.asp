<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺��� ����� ����ȸ���� īī���� ä�� ����� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/channel#ListPlusFriendID
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    On Error Resume Next

    Set Presponse = m_KakaoService.ListPlusFriendID(testCorpNum)

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
                <legend>īī���� ä�� ���� ��� Ȯ��</legend>
                <%
                    For i=0 To Presponse.length -1
                %>
                <fieldset class="fieldset2">
                <ul>
                    <li>īī���� �˻��� ���̵� (plusFriendID) : <%=Presponse.Get(i).plusFriendID%> </li>
                    <li>īī���� ä�� �̸� (plusFriendName) : <%=Presponse.Get(i).plusFriendName%> </li>
                    <li>����Ͻ� (regDT) : <%=Presponse.Get(i).regDT%> </li>
                    <li>ä�� ���� (state) : <%=Presponse.Get(i).state%> </li>
                    <li>ä�� ���� �Ͻ� (stateDT) : <%=Presponse.Get(i).stateDT%> </li>
                </ul>
                </fieldset>
                <%
                    Next
                %>

         </div>
    </body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺��� ����� ����ȸ���� ���� �߽Ź�ȣ ����� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/sms/asp/api/sendnum#GetSenderNumberList
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    On Error Resume Next

    Set Presponse = m_MessageService.GetSenderNumberList(testCorpNum)

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
                <legend>���� �߽Ź�ȣ ��� Ȯ��</legend>
                <%
                    For i=0 To Presponse.length -1
                %>
                <fieldset class="fieldset2">
                <ul>
                    <li>�߽Ź�ȣ (number) : <%=Presponse.Get(i).number%> </li>
                    <li>��ǥ��ȣ �������� (representYN) : <%=Presponse.Get(i).representYN%> </li>
                    <li>��ϻ��� (state) : <%=Presponse.Get(i).state%> </li>
                    <li>�޸� (memo) : <%=Presponse.Get(i).memo%> </li>
                </ul>
                </fieldset>
                <%
                    Next
                %>

         </div>
    </body>
</html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݰ�꼭�� ÷�ε� ���ϸ���� Ȯ���մϴ�.
    ' - �����׸� �� ���Ͼ��̵�(AttachedFile) �׸��� ���ϻ���(DeleteFile API) ȣ��� �̿��� �� �ֽ��ϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#GetFiles
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ���� 10�ڸ�
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    testUserID = "testkorea"

    ' ���ݰ�꼭 �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType = "SELL"

    ' ������ȣ
    MgtKey = "20220720-ASP-002"

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.GetFiles(testCorpNum, KeyType ,MgtKey, testUserID)

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
                <legend>���ݰ�꼭 ÷������ ��� Ȯ��</legend>
                    <%
                        If code = 0 Then
                            For i=0 To Presponse.length -1
                    %>
                            <fieldset class="filedset2">
                            <legend> ÷������ : <%=i+1%> </legend>
                                <ul>
                                    <li> serialNum(����) : <%=Presponse.Get(i).serialNum%></li>
                                    <li> AttachedFile(���ϸ�) : <%=Presponse.Get(i).AttachedFile%></li>
                                    <li> DisplayName(���Ͼ��̵�) : <%=Presponse.Get(i).DisplayName%></li>
                                    <li> regDT(����Ͻ�) : <%=Presponse.Get(i).regDT%></li>
                                </ul>
                            </fieldset>
                    <%
                        Next
                        Else
                    %>
                            <ul>
                                <li>Response.dcode : <%=code%> </li>
                                <li>Response.message : <%=message%> </li>
                            </ul>
                    <%
                        End If
                    %>
            </fieldset>
         </div>
    </body>
</html>
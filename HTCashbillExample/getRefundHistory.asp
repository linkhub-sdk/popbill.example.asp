<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����ȸ���� ����Ʈ ȯ�ҽ�û������ Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/htcashbill/asp/api/point#GetRefundHistory
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    '��� ��������ȣ
    Page = 1

    ' �������� ǥ���� ��ϰ���
    PerPage = 500

    '�˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_HTCashbillService.GetRefundHistory(CorpNum, Page, PerPage, UserID)

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
            <legend>����ȸ�� ����Ʈ ȯ�ҳ��� Ȯ��</legend>
            <%
                If code = 0 Then
            %>

            <ul>
                <li> code (���� �ڵ�) : <%=result.code%></li>
                <li> total (�� �˻���� �Ǽ�) : <%=result.total%></li>
                <li> perPage (�������� �˻�����) : <%=result.perPage%></li>
                <li> pageNum (������ ��ȣ) : <%=result.pageNum%></li>
                <li> pageCount (������ ����) : <%=result.pageCount%></li>
            </ul>
                <%
                    Dim i
                    For i = 0 To UBound(result.list) - 1
                %>
                <fieldset class="fieldset2">
                    <legend> RefundHistory  [ <%= i+1%> / <%=UBound(result.list)%>]</legend>
                    <ul>
                        <li> reqDT (��û �Ͻ�) : <%=result.list(i).reqDT%></li>
                        <li> requestPoint (ȯ�� ��û����Ʈ) : <%=result.list(i).requestPoint%></li>
                        <li> accountBank (ȯ�Ұ��� �����) : <%=result.list(i).accountBank%></li>
                        <li> accountNum (ȯ�Ұ��¹�ȣ) : <%=result.list(i).accountNum%></li>
                        <li> accountName (ȯ�Ұ��� �����ָ�) : <%=result.list(i).accountName%></li>
                        <li> state (����) : <%=result.list(i).state%></li>
                        <li> reason (ȯ�һ���) : <%=result.list(i).reason%></li>
                    </ul>
                </fieldset>
                <%
                    Next
                %>
            <%
                Else
            %>
                <ul>
                    <li>Response.code: <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            <%
                End If
            %>
            </fieldset>
        </div>
    </body>
</html>

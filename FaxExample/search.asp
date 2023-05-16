<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˻������� ����Ͽ� �ѽ����� ������ ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 2����)
    ' - �ѽ� �����Ͻ÷κ��� 2���� �̳� �����Ǹ� ��ȸ�� �� �ֽ��ϴ�.
    ' - https://developers.popbill.com/reference/fax/asp/api/info#Search
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    '��������, yyyyMMdd
    SDate = "20220701"

    '��������, yyyyMMdd
    EDate = "20220720"

    ' ���ۻ��� �迭 ("1" , "2" , "3" , "4" �� ����, ���� ���� ����)
    ' �� 1 = ��� , 2 = ���� , 3 = ���� , 4 = ���
    ' - ���Է� �� ��ü��ȸ
    Dim State(4)
    State(0) = "1"
    State(1) = "2"
    State(2) = "3"
    State(3) = "4"

    ' ���࿩�� (false , true �� �� 1)
    ' false = ��ü��ȸ, true = �������۰� ��ȸ
    ' ���Է½� �⺻�� false ó��
    ReserveYN = False

    ' ������ȸ ���� (false , true �� �� 1)
    ' false = ������ �ѽ� ��ü ��ȸ (�����ڱ���)
    ' true = �ش� ����� �������� ������ �ѽ��� ��ȸ (���α���)
    ' ���Է½� �⺻�� false ó��
    SenderOnlyYN = False

    '���Ĺ���, A-��������, D-��������
    Order = "D"

    '������ ��ȣ
    Page = 1

    '�������� �˻�����
    PerPage = 20

    ' ��ȸ�ϰ��� �ϴ� �߽��ڸ� �Ǵ� �����ڸ�
    ' - ���Է½� ��ü��ȸ
    QString = ""

    On Error Resume Next

    Set result = m_FaxService.Search(testCorpNum, SDate, EDate, State, ReserveYN, SenderOnlyYN, Order, Page, PerPage, QString)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message = Presponse.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>�ѽ����� ���۳��� ��ȸ </legend>
                    <ul>
                        <li> code (�����ڵ�) : <%=result.code%></li>
                        <li> total (�� �˻���� �Ǽ�) : <%=result.total%></li>
                        <li> pageNum (������ ��ȣ) : <%=result.pageNum%></li>
                        <li> perPage (�������� ��ϰ���) : <%=result.perPage%></li>
                        <li> pageCount (������ ����) : <%=result.pageCount%></li>
                        <li> message (����޽���) : <%=result.message%></li>
                    </ul>
                <% If code = 0 Then
                        For i=0 To UBound(result.list)-1
                %>
                    <fieldset class="fieldset2">
                            <legend> �ѽ� ���۰�� [ <%=i+1%> /  <%=UBound(result.list)%> ] </legend>
                            <ul>
                                <li>state (���ۻ��� �ڵ�) : <%=result.list(i).state%> </li>
                                <li>result (���۰�� �ڵ�) : <%=result.list(i).result%> </li>
                                <li>sendNum (�߽Ź�ȣ) : <%=result.list(i).sendNum%> </li>
                                <li>senderName (�߽��ڸ�) : <%=result.list(i).senderName%> </li>
                                <li>receiveNum (���Ź�ȣ) : <%=result.list(i).receiveNum%> </li>
                                <li>receiveNumType (���Ź�ȣ ����) : <%=result.list(i).receiveNumType%> </li>
                                <li>receiveName (�����ڸ�) : <%=result.list(i).receiveName%> </li>
                                <li>title (�ѽ� ����) : <%=result.list(i).title %> </li>
                                <li>sendPageCnt (��������) : <%=result.list(i).sendPageCnt%></li>
                                <li>successPageCnt (���� ��������) : <%=result.list(i).successPageCnt%></li>
                                <li>failPageCnt (���� ��������) : <%=result.list(i).failPageCnt%></li>
                                <li>refundPageCnt (ȯ�� ��������) : <%=result.list(i).refundPageCnt%></li>
                                <li>cancelPageCnt (��� ��������) : <%=result.list(i).cancelPageCnt%></li>
                                <li>reserveDT (����ð�) : <%=result.list(i).reserveDT%></li>
                                <li>sendDT (�߼۽ð�) : <%=result.list(i).sendDT%></li>
                                <li>receiptDT (���� �����ð�) : <%=result.list(i).receiptDT%></li>
                                <li>fileNames (�������ϸ� �迭) : <%=result.list(i).fileNames%></li>
                                <li>interOPRefKey (��Ʈ�� ����Ű) : <%=result.list(i).interOPRefKey%> </li>
                                <li>receiptNum (������ȣ) : <%=result.list(i).receiptNum%> </li>
                                <li>requestNum (��û��ȣ) : <%=result.list(i).requestNum%> </li>
                                <li>chargePageCnt (���� ��������) : <%=result.list(i).chargePageCnt%> </li>
                                <li>tiffFileSize (��ȯ���Ͽ뷮 (���� : byte)) : <%=result.list(i).tiffFileSize%> </li>
                            </ul>
                        </fieldset>
                <%
                    Next
                    Else
                %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <%	End If	%>

            </fieldset>
         </div>
    </body>
</html>
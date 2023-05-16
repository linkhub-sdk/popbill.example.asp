
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˻������� ����Ͽ� īī���� ���۳��� ����� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 2����)
    ' - īī���� �����Ͻ÷κ��� 6���� �̳� �����Ǹ� ��ȸ�� �� �ֽ��ϴ�.
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/info#Search
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    '��������
    SDate = "20220701"

    '��������
    EDate = "20220720"

    ' ���ۻ��� �迭 ("0" , "1" , "2" , "3" , "4" , "5" �� ����, ���� ���� ����)
    ' �� 0 = ���۴�� , 1 = ������ , 2 = ���ۼ��� , 3 = ��ü���� ���� , 4 = ���۽��� , 5 = �������
    ' - ���Է� �� ��ü��ȸ
    Dim State(6)
    State(0) = "0"
    State(1) = "1"
    State(2) = "2"
    State(3) = "3"
    State(4) = "4"
    State(5) = "5"

    ' �˻���� �迭 ("ATS", "FTS", "FMS" �� ����, ���� ���� ����)
    ' �� ATS = �˸��� , FTS = ģ����(�ؽ�Ʈ) , FMS = ģ����(�̹���)
    ' - ���Է� �� ��ü��ȸ
    Dim Item(3)
    Item(0) = "ATS"
    Item(1) = "FTS"
    Item(2) = "FMS"

    ' ���������� ��ȸ (null , "0" , "1" �� �� 1)
    ' �� null = ��ü , 0 = ������۰� , 1 = �������۰�
    ' - ���Է� �� ��ü��ȸ
    ReserveYN = ""

    ' ����ڱ��Ѻ� ��ȸ (true / false �� �� 1)
    ' �� false = ������ īī���� ��ü ��ȸ (�����ڱ���)
    ' �� true = �ش� ����� �������� ������ īī���常 ��ȸ (���α���)
    ' ���Է½� �⺻�� false ó��
    SenderYN = False

    ' ���Ĺ���, D-��������, A-��������
    Order = "D"

    ' ������ ��ȣ
    Page = 1

    PerPage = 30

    ' ��ȸ�ϰ��� �ϴ� �����ڸ�
    ' - ���Է½� ��ü��ȸ
    QString = ""

    On Error Resume Next

    Set resultObj = m_KakaoService.Search(testCorpNum, SDate, EDate, State, Item, ReserveYN, SenderYN, Order, Page, PerPage, QString)

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

                    <legend>īī���� ���۳��� ��ȸ </legend>
                    <ul>
                    <% If code = 0 Then %>
                            <li> code (�����ڵ�) : <%=resultObj.code%></li>
                            <li> message (����޽���) : <%=resultObj.message%></li>
                            <li> total (�� �˻���� �Ǽ�) : <%=resultObj.total%></li>
                            <li> pageNum (������ ��ȣ) : <%=resultObj.pageNum%></li>
                            <li> pageCount (������ ����) : <%=resultObj.pageCount%></li>
                            <li> perPage (�������� �˻�����) : <%=resultObj.perPage%></li>
                    </ul>
                        <%
                            For i=0 To UBound(resultObj.list) -1
                        %>
                            <fieldset class="fieldset2">
                                <legend> īī���� ���۰�� [ <%=i+1%> / <%= UBound(resultObj.list)%> ] </legend>
                                <ul>
                                    <li>state (���ۻ��� �ڵ�) : <%=resultObj.list(i).state%> </li>
                                    <li>sendDT (�����Ͻ�) : <%=resultObj.list(i).sendDT%> </li>
                                    <li>result (���۰�� �ڵ�) : <%=resultObj.list(i).result%> </li>
                                    <li>resultDT (���۰�� �����Ͻ�) : <%=resultObj.list(i).resultDT%> </li>
                                    <li>contentType (īī���� ����) : <%=resultObj.list(i).contentType%> </li>
                                    <li>receiveNum (���Ź�ȣ) : <%=resultObj.list(i).receiveNum%> </li>
                                    <li>receiveName (�����ڸ�) : <%=resultObj.list(i).receiveName%> </li>
                                    <li>content (�˸���/ģ���� ����) : <%=resultObj.list(i).content%> </li>
                                    <li>altSubject (��ü���� ����) : <%=resultObj.list(i).altSubject%></li>
                                    <li>altContent (��ü���� ����) : <%=resultObj.list(i).altContent%></li>
                                    <li>altContentType (��ü���� ����Ÿ��) : <%=resultObj.list(i).altContentType%> </li>
                                    <li>altSendDT (��ü���� �����Ͻ�) : <%=resultObj.list(i).altSendDT%> </li>
                                    <li>altResult (��ü���� ���۰�� �ڵ�) : <%=resultObj.list(i).altResult%> </li>
                                    <li>altResultDT (��ü���� ���۰�� �����Ͻ�) : <%=resultObj.list(i).altResultDT%> </li>
                                    <li>receiptNum (������ȣ) : <%=resultObj.list(i).receiptNum%> </li>
                                    <li>requestNum (��û��ȣ) : <%=resultObj.list(i).requestNum%> </li>
                                    <li>interOPRefKey (��Ʈ�� ����Ű) : <%=resultObj.list(i).interOPRefKey%> </li>
                                </ul>
                            </fieldset>
                        <%
                            Next
                        Else
                        %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>

            </fieldset>
         </div>
    </body>
</html>
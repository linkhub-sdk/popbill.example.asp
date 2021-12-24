<%
Class StatementService

Private m_PopbillBase

'�׽�Ʈ �÷���
Public Property Let IsTest(ByVal value)
    m_PopbillBase.IsTest = value
End Property

Public Property Let IPRestrictOnOff(ByVal value)
    m_PopbillBase.IPRestrictOnOff = value
End Property

Public Property Let UseStaticIP(ByVal value)
    m_PopbillBase.UseStaticIP = value
End Property

Public Property Let UseGAIP(ByVal value)
    m_PopbillBase.UseGAIP = value
End Property

Public Property Let UseLocalTimeYN(ByVal value)
    m_PopbillBase.UseLocalTimeYN = value
End Property

Public Sub Class_Initialize
    Set m_PopbillBase = New PopbillBase
    m_PopbillBase.AddScope("121")
    m_PopbillBase.AddScope("122")
    m_PopbillBase.AddScope("123")
    m_PopbillBase.AddScope("124")
    m_PopbillBase.AddScope("125")
    m_PopbillBase.AddScope("126")
End Sub

Public Sub Initialize(linkID, SecretKey )
    m_PopbillBase.Initialize linkID,SecretKey
End Sub

'ȸ���ܾ���ȸ
Public Function GetBalance(CorpNum)
    GetBalance = m_PopbillBase.GetBalance(CorpNum)
End Function
'��Ʈ�� �ܾ���ȸ
Public Function GetPartnerBalance(CorpNum)
    GetPartnerBalance = m_PopbillBase.GetPartnerBalance(CorpNum)
End Function

'��Ʈ�� ����Ʈ ���� �˾� URL - 2017/08/29 �߰�
Public Function GetPartnerURL(CorpNum, TOGO)
    GetPartnerURL = m_PopbillBase.GetPartnerURL(CorpNum,TOGO)
End Function

'�˺� �⺻ URL
Public Function GetPopbillURL(CorpNum , UserID , TOGO )
    GetPopbillURL = m_PopbillBase.GetPopbillURL(CorpNum , UserID , TOGO )
End Function
'�˺� �α��� URL
Public Function GetAccessURL(CorpNum , UserID)
    GetAccessURL = m_PopbillBase.GetAccessURL(CorpNum , UserID )
End Function

'�˺� ����ȸ�� ����Ʈ ���� URL
Public Function GetChargeURL(CorpNum , UserID)
    GetChargeURL = m_PopbillBase.GetChargeURL(CorpNum , UserID )
End Function

'�˺� ����ȸ�� ����Ʈ �������� URL
Public Function GetPaymentURL(CorpNum, UserID)
    GetPaymentURL = m_PopbillBase.GetPaymentURL(CorpNum, UserID)
End Function

'�˺� ����ȸ�� ����Ʈ ��볻�� URL
Public Function GetUseHistoryURL(CorpNum, UserID)
    GetUseHistoryURL = m_PopbillBase.GetUseHistoryURL(CorpNum, UserID)
End Function

'ȸ������ ����
Public Function CheckIsMember(CorpNum , linkID)
    Set CheckIsMember = m_PopbillBase.CheckIsMember(CorpNum,linkID)
End Function
'ȸ������
Public Function JoinMember(JoinInfo)
    Set JoinMember = m_PopbillBase.JoinMember(JoinInfo)
End Function

'����� ���� Ȯ��
Public Function GetContactInfo(CorpNum, ContactID, UserID)
    Set GetContactInfo = m_PopbillBase.GetContactInfo(CorpNum, ContactID, UserID)
End Function 

'����� �����ȸ
Public Function ListContact(CorpNum, UserID)
    Set ListContact = m_popbillBase.ListContact(CorpNum,UserID)
End Function
'����� ��������
Public Function UpdateContact(CorpNum, contInfo, UserId)
    Set UpdateContact = m_popbillBase.UpdateContact(CorpNum, contInfo, UserId)
End Function
'����� �߰� 
Public Function RegistContact(CorpNum, contInfo, UserId)
    Set RegistContact = m_popbillBase.RegistContact(CorpNum, contInfo, UserId)
End Function
'ȸ������ ����
Public Function UpdateCorpInfo(CorpNum, corpInfo, UserId)
    Set UpdateCorpInfo = m_popbillBase.UpdateCorpInfo(CorpNum, corpInfo, UserId)
End Function
'ȸ������ Ȯ�� 
Public Function GetCorpInfo(CorpNum, UserId)
    Set GetCorpInfo = m_popbillBase.GetCorpInfo(CorpNum, UserId)
End Function
Public Function CheckID(id)
    Set CheckID = m_popbillBase.CheckID(id)
End Function

'�������� Ȯ��
Public Function GetChargeInfo ( CorpNum, ItemCode, UserID )
    Dim result : Set result = m_PopbillBase.httpGET ( "/Statement/ChargeInfo/" + CStr(ItemCode), m_PopbillBase.getSession_token(CorpNum), UserID )

    Dim chrgInfo : Set chrgInfo = New ChargeInfo
    chrgInfo.fromJsonInfo result
    
    Set GetChargeInfo = chrgInfo
End Function 
'''''''''''''  End of PopbillBase

''�ܰ�Ȯ��
Public Function GetUnitCost(CorpNum, itemCode)
    Dim result : Set result = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode) + "?cfg=UNITCOST", m_PopbillBase.getSession_token(CorpNum),"")
    GetUnitCost = result.unitCost
End Function

'����������ȣ ��뿩�� Ȯ��
Public Function CheckMgtKeyInUse(CorpNum, itemCode, mgtKey) 
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    On Error Resume Next
    Dim result : Set result = m_PopbillBase.httpGet("/Statement/"+CStr(itemCode)+"/"+mgtKey, m_PopbillBase.getSession_token(CorpNum),"")
    
    If Err.Number = -12000004 Then
        CheckMgtKeyInUse = False
        Err.Clears
    Else 
        CheckMgtKeyInUse = True
        Err.Clears
    End If
    On Error GoTo 0
End Function 


'�˺� SSO URLȮ��
Public Function GetURL(CorpNum, UserID, TOGO)
    Dim result : Set result = m_PopbillBase.httpGet("/Statement?TG=" + TOGO, m_PopbillBase.getSession_token(CorpNum),UserID)
    GetURL = result.url
End Function 

'�ӽ�����
Public Function Register(CorpNum, ByRef statement, UserID)
    Dim tmpJson : Set tmpJson = statement.toJsonInfo

    Dim postdata : postdata = m_PopbillBase.toString(tmpJson)
    
    Set Register = m_PopbillBase.httpPOST("/Statement", m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)
End Function 


'����
Public Function Update(CorpNum, itemCode, mgtKey, ByRef statement, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim tmpJson : Set tmpJson = statement.toJsonInfo

    Dim postdata : postdata = m_PopbillBase.toString(tmpJson)
    
    Set Update = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode)+"/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "PATCH", postdata, UserID)
End Function 

'����
Public Function Issue(CorpNum, itemCode, mgtKey, Memo, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "memo", Memo

    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Set Issue = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode)+"/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "ISSUE", postdata, UserID)
End Function 


'�������
Public Function CancelIssue(CorpNum, itemCode, mgtKey, Memo, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "memo", Memo

    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Set CancelIssue = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode)+"/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "CANCEL", postdata, UserID)

End Function

'����
Public Function Delete(CorpNum, itemCode, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Set Delete = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode)+"/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "DELETE", "", UserID)
End Function


'�˺� �ΰ� �� ÷�ι��� ���  URLȮ��
Public Function GetSealURL(CorpNum, UserID)
    Dim result : Set result = m_PopbillBase.httpGET("/?TG=SEAL", _
                        m_PopbillBase.getSession_token(CorpNum), UserID)
    GetSealURL = result.url
End Function

'���� ÷��
Public Function AttachFile(CorpNum, itemCode, mgtKey, filePath, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Set AttachFile = m_PopbillBase.httpPOST_File("/Statement/"+CStr(itemCode)+"/"+mgtKey+"/Files", m_PopbillBase.getSession_token(CorpNum), filePath, UserID)
End Function 


'÷������ ���
Public Function GetFiles(CorpNum, itemCode, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Set GetFiles = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode)+"/"+mgtKey+"/Files", m_PopbillBase.getSession_token(CorpNum), UserID)
End Function 


'÷������ ����
Public Function DeleteFile(CorpNum, itemCode, mgtKey, FileID, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Set DeleteFile = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode)+"/"+mgtKey+"/Files/"+FileID, m_PopbillBase.getSession_token(CorpNum), "DELETE", "", UserID)
End Function 


'���� ���� ������� Ȯ��
Public Function GetInfo(CorpNum, itemCode, mgtKey, UserID)
    If mgtKey = "" Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode)+"/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), UserID)

    Dim infoObj : Set infoObj = New StatementInfo
    infoObj.fromJsonInfo result

    Set GetInfo = infoObj

End Function 

'�ٷ� ���� ������� Ȯ��
Public Function GetInfos(CorpNum, itemCode, mgtKeyList, UserID)
    If isNull(mgtKeyList) Or isEmpty(mgtKeyList)  Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim tmp : Set tmp = JSON.parse("[]")

    Dim i
    For i=0 To Ubound(mgtKeyList)-1
        tmp.Set i, mgtKeyList(i)
    Next

    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Dim infoList : Set infoList = CreateObject("Scripting.Dictionary")

    Dim result : Set result = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode), m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)

    For i=0 To result.length-1
        Dim tmpObj : Set tmpObj = New StatementInfo
        tmpObj.fromJsonInfo result.Get(i)
        infoList.Add i, tmpObj
    Next

    Set GetInfos = infoList
End Function 

'�̷� Ȯ��
Public Function GetLogs(CorpNum, itemCode, mgtKey, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey)  Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode)+"/"+mgtKey+"/Logs", m_PopbillBase.getSession_token(CorpNum), UserID)
    
    Dim logList : Set logList = CreateObject("Scripting.Dictionary")

    Dim i
    For i=0 To result.length-1
        Dim logObj : Set logObj = New StatementLog
        logObj.fromJsonInfo result.Get(i)
        logList.Add i, logObj
    Next

    Set GetLogs = logList
End Function 



'������ Ȯ��
Public Function GetDetailInfo(CorpNum, itemCode, mgtKey, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey)  Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode)+"/"+mgtKey+"?Detail", m_PopbillBase.getSession_token(CorpNum), UserID)

    Dim infoObj : Set infoObj = New Statement
    
    infoObj.fromJsonInfo result

    Set GetDetailInfo = infoObj
End Function 


'�˸����� ����
Public Function SendEmail(CorpNum, itemCode, mgtKey, receiver, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey)  Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "receiver", receiver

    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Set SendEmail = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode)+"/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "EMAIL", postdata, UserID)
End Function 


'�˸����� ����
Public Function SendSMS(CorpNum, itemCode, mgtKey, sender, receiver, contents, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "receiver", receiver
    tmp.Set "sender", sender
    tmp.Set "contents", contents

    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Set SendSMS = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode)+"/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "SMS", postdata, UserID)
End Function


'���ڸ��� �ѽ� ����
Public Function SendFAX(CorpNum, itemCode, mgtKey, sender, receiver, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "receiver", receiver
    tmp.Set "sender", sender
    
    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Set SendFAX = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode)+"/"+mgtKey, m_PopbillBase.getSession_token(CorpNum), "FAX", postdata, UserID)
End Function 


'���ڸ��� ���� URL
Public Function GetPopUpURL(CorpNum, itemCode, mgtKey, UserID)
    Dim result : Set result = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode)+"/"+mgtKey+"?TG=POPUP", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetPopUpURL = result.url
End Function 

'���ڸ��� ���� URL
Public Function GetViewURL(CorpNum, itemCode, mgtKey, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result : Set result = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode)+"/"+mgtKey+"?TG=VIEW", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetViewURL = result.url
End Function 

'�μ� URL ȣ��
Public Function GetPrintURL(CorpNum, itemCode, mgtKey, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result : Set result = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode)+"/"+mgtKey+"?TG=PRINT", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetPrintURL = result.url
End Function 


'�μ� URL ȣ��(���޹޴��ڿ�)
Public Function GetEPrintURL(CorpNum, itemCode, mgtKey, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode)+"/"+mgtKey+"?TG=EPRINT", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetEPrintURL = result.url
End Function 


'���� ��ũ URL ȣ��
Public Function GetMailURL(CorpNum, itemCode, mgtKey, UserID)
    If isNull(mgtKey) Or isEmpty(mgtKey) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim result : Set result = m_PopbillBase.httpGET("/Statement/"+CStr(itemCode)+"/"+mgtKey+"?TG=MAIL", m_PopbillBase.getSession_token(CorpNum), UserID)
    GetMailURL = result.url
End Function


'�ٷ� �μ� URL ȣ��
Public Function GetMassPrintURL(CorpNum, itemCode, mgtKeyList, UserID)
    If isNull(mgtKeyList) Or isEmpty(mgtKeyList) Then 
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim tmp : Set tmp = JSON.parse("[]")
    Dim i
    For i=0 To UBound(mgtKeyList)-1
        tmp.Set i, mgtKeyList(i)
    Next

    Dim postdata : postdata = m_PopbillBase.toString(tmp)

    Dim result : Set result = m_PopbillBase.httpPOST("/Statement/"+CStr(itemCode)+"?Print", m_PopbillBase.getSession_token(CorpNum), "", postdata, UserID)
    GetMassPrintURL = result.url
End Function

'���ڸ��� ��ù���
Public Function RegistIssue(CorpNum, ByRef statement, Memo, UserID, EmailSubject)
    If statement Is Nothing Then Err.raise -99999999,"POPBILL","����� ���ڸ��� ������ �Էµ��� �ʾҽ��ϴ�."

    Dim tmpDic : Set tmpDic = statement.toJsonInfo

    If Memo <> "" Then
        tmpDic.Set "memo", Memo
    End If

    If EmailSubject <> "" Then
        tmpDic.Set "emailSubject", EmailSubject
    End If


    Dim postdata : postdata = m_PopbillBase.toString(tmpDic)

    Set RegistIssue = m_PopbillBase.httpPOST("/Statement", m_PopbillBase.getSession_token(CorpNum), _
                    "ISSUE", postdata, UserID)

End Function

'���ѽ� ����
Public Function FAXSend(CorpNum, ByRef statement, SendNum, ReceiveNum, UserID)
    If statement Is Nothing Then Err.raise -99999999,"POPBILL","������ ���ڸ��� ������ �Էµ��� �ʾҽ��ϴ�."
    
    If isNull(ReceiveNum) Or isEmpty(ReceiveNum) Then 
        Err.Raise -99999999, "POPBILL", "�����ѽ���ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim tmpDic : Set tmpDic = statement.toJsonInfo
    tmpDic.Set "sendNum", SendNum
    tmpDic.Set "receiveNum", ReceiveNum

    Dim postdata : postdata = m_PopbillBase.toString(tmpDic)
    
    Dim result : Set result = m_PopbillBase.httpPOST("/Statement", m_PopbillBase.getSession_token(CorpNum), "FAX", postdata, UserID)
    FAXSend = result.receiptNum

End Function 


'���ڸ��� ��� ��ȸ
Public Function Search(CorpNum, DType, SDate, EDate, State, ItemCode, Order, Page, PerPage, QString)
    If DType = "" Then
        Err.Raise -99999999, "POPBILL", "�˻����� ������ �Էµ��� �ʾҽ��ϴ�."
    End If
    If SDate = "" Then
        Err.Raise -99999999, "POPBILL", "�������ڰ� �Էµ��� �ʾҽ��ϴ�."
    End If
    If EDate = "" Then
        Err.Raise -99999999, "POPBILL", "�������ڰ� �̷µ��� �ʾҽ��ϴ�."
    End If
    Dim uri
    uri = "/Statement/Search"
    uri = uri & "?DType=" & DType
    uri = uri & "&SDate=" & SDate
    uri = uri & "&EDate=" & EDate

    uri = uri & "&State="
    Dim i
    For i=0 To UBound(State) -1	
        If i = UBound(State) -1 then
            uri = uri & State(i)
        Else
            uri = uri & State(i) & ","
        End If
    Next

    uri = uri & "&ItemCode="
    For i=0 To UBound(Itemcode) -1
        If i = UBound(Itemcode) -1  then	
            uri = uri & Itemcode(i)
        Else
            uri = uri & Itemcode(i) & ","
        End If
    Next

    uri = uri & "&QString=" & QString
    uri = uri & "&Order=" & Order
    uri = uri & "&Page=" & CStr(Page)
    uri = uri & "&PerPage=" & CStr(PerPage)
    
    Dim searchResult : Set searchResult = New StmtSearchResult
    Dim tmpObj : Set tmpObj = m_PopbillBase.httpGET(uri, m_PopbillBase.getSession_token(CorpNum), "")

    searchResult.fromJsonInfo tmpObj
    
    Set Search = searchResult
End Function

Public Function AttachStatement(CorpNum, ItemCode, MgtKey, SubItemCode, SubMgtKey)
    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "ItemCode", SubItemCode
    tmp.Set "MgtKey", SubMgtKey
    
    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    Set AttachStatement = m_PopbillBase.httpPOST("/Statement/"+CStr(ItemCode)+"/"+mgtKey+"/AttachStmt", _
                        m_PopbillBase.getSession_token(CorpNum), "", postdata, "")
End Function

Public Function DetachStatement(CorpNum, ItemCode, MgtKey, SubItemCode, SubMgtKey)
    Dim tmp : Set tmp = JSON.parse("{}")
    tmp.Set "ItemCode", SubItemCode
    tmp.Set "MgtKey", SubMgtKey
    
    Dim postdata : postdata = m_PopbillBase.toString(tmp)
    Set DetachStatement = m_PopbillBase.httpPOST("/Statement/"+CStr(ItemCode)+"/"+mgtKey+"/DetachStmt", _
                        m_PopbillBase.getSession_token(CorpNum), "", postdata, "")
End Function


'�˸����� ���۸�� ��ȸ
Public Function listEmailConfig(CorpNum, UserID)
    If CorpNum = "" Or isEmpty(CorpNum) Then 
        Err.Raise -99999999, "POPBILL", "����ڵ�Ϲ�ȣ�� �ùٸ��� �ʽ��ϴ�."
    End If

    Dim result : Set result = m_PopbillBase.httpGet("/Statement/EmailSendConfig", m_PopbillBase.getSession_token(CorpNum), UserID)
    
    Dim tmpDic : Set tmpDic = CreateObject("Scripting.Dictionary")

    Dim i
    For i=0 To result.length-1
        Dim emailObj : Set emailObj = New EmailSendConfig	
        emailObj.fromJsonInfo result.Get(i)
        tmpDic.Add i, emailObj
    Next
    
    Set listEmailConfig = tmpDic
End Function 

'�˸����� ���ۼ��� ����
Public Function updateEmailConfig(CorpNum, mailType, sendYN, UserID)
    If CorpNum = "" Or isEmpty(CorpNum) Then 
        Err.Raise -99999999, "POPBILL", "����ڵ�Ϲ�ȣ�� �ùٸ��� �ʽ��ϴ�."
    End If

    If mailType = "" Or isEmpty(mailType) Then 
        Err.Raise -99999999, "POPBILL", "�������� Ÿ���� �Էµ��� �ʾҽ��ϴ�."
    End If

    If sendYN = "" Or isEmpty(sendYN) Then 
        Err.Raise -99999999, "POPBILL", "�������� ���� �׸��� �Էµ��� �ʾҽ��ϴ�."
    End If

    If (sendYN) Then
        sendYN="true"
    Else
        sendYN="false"
    End If
    
    Dim uri : uri = "/Statement/EmailSendConfig?EmailType="+mailType+"&SendYN="+sendYN

    Set updateEmailConfig = m_PopbillBase.httpPOST(uri, m_PopbillBase.getSession_token(CorpNum), "", "", UserID)
End Function

End Class

Class StatementLog
Public docLogType
Public log
Public procType
Public procCorpName
Public procMemo
Public regDT
Public ip

Public Sub fromJsonInfo(jsonInfo)
    On Error Resume Next
    docLogType = jsonInfo.docLogType
    log = jsonInfo.docLogType
    procType = jsonInfo.procType
    procCorpName = jsonInfo.procCorpName
    procMemo = jsonInfo.procMemo
    regDT = jsonInfo.regDT
    ip = jsonInfo.ip
    On Error GoTo 0 
End Sub

End Class

Class StatementInfo
Public itemKey
Public stateCode
Public itemCode
Public taxType
Public purposeType
Public writeDate
Public senderCorpName
Public senderCorpNum
Public senderPrintYN
Public receiverCorpName
Public receiverCorpNum
Public receiverPrintYN
Public supplyCostTotal
Public taxTotal
Public issueDT
Public stateDT
Public openYN
Public openDT
Public stateMemo
Public regDT

Public Sub fromJsonInfo(jsonInfo)
    On Error Resume Next
    itemKey = jsonInfo.itemKey
    stateCode = jsonInfo.stateCode
    itemCode = jsonInfo.itemCode
    taxType = jsonInfo.taxType
    purposeType = jsonInfo.purposeType
    writeDate = jsonInfo.writeDate
    senderCorpName = jsonInfo.senderCorpName
    senderCorpNum = jsonInfo.senderCorpNum
    senderPrintYN = jsonInfo.senderPrintYN
    receiverCorpName = jsonInfo.receiverCorpName
    receiverCorpNum = jsonInfo.receiverCorpNum
    receiverPrintYN = jsonInfo.receiverPrintYN
    supplyCostTotal = jsonInfo.supplyCostTotal
    taxTotal = jsonInfo.taxTotal
    issueDT = jsonInfo.issueDT
    stateDT = jsonInfo.stateDT
    openYN = jsonInfo.openYN
    openDT = jsonInfo.openDT
    stateMemo = jsonInfo.stateMemo
    regDT = jsonInfo.regDT
    On Error GoTo 0
End Sub

End Class

Class Statement
Public itemCode             
Public mgtKey               
Public invoiceNum           
Public formCode             
Public writeDate            
Public taxType              

Public senderCorpNum      
Public senderTaxRegID     
Public senderCorpName     
Public senderCEOName      
Public senderAddr         
Public senderBizClass     
Public senderBizType      
Public senderContactName  
Public senderDeptName     
Public senderTEL          
Public senderHP           
Public senderEmail        
Public senderFAX          

Public receiverCorpNum    
Public receiverTaxRegID   
Public receiverCorpName   
Public receiverCEOName    
Public receiverAddr       
Public receiverBizClass   
Public receiverBizType    
Public receiverContactName
Public receiverDeptName   
Public receiverTEL        
Public receiverHP         
Public receiverEmail      
Public receiverFAX        

Public taxTotal           
Public supplyCostTotal    
Public totalAmount        
Public purposeType        
Public serialNum          
Public remark1            
Public remark2            
Public remark3            
Public businessLicenseYN  
Public bankBookYN         
Public faxsendYN          
Public smssendYN          
Public autoacceptYN       

Public detailList()
Public propertyBag

Public Sub AddDetail(detail)
    ReDim Preserve detailList(UBound(detailList) + 1)
    Set detailList(Ubound(detailList)) = detail
End Sub

Public Sub Class_Initialize
    ReDim detailList(-1)
    Set properTyBag = JSON.parse("{}")
End Sub

Public Function toJsonInfo()
    Set toJsonInfo = JSON.parse("{}")
    toJsonInfo.Set "itemCode", itemCode     
    toJsonInfo.Set "mgtKey", mgtKey               
    toJsonInfo.Set "invoiceNum", invoiceNum           
    toJsonInfo.Set "formCode", formCode             
    toJsonInfo.Set "writeDate", writeDate            
    toJsonInfo.Set "taxType", taxType               

    toJsonInfo.Set "senderCorpNum", senderCorpNum      
    toJsonInfo.Set "senderTaxRegID", senderTaxRegID     
    toJsonInfo.Set "senderCorpName", senderCorpName     
    toJsonInfo.Set "senderCEOName", senderCEOName      
    toJsonInfo.Set "senderAddr", senderAddr         
    toJsonInfo.Set "senderBizClass", senderBizClass     
    toJsonInfo.Set "senderBizType", senderBizType      
    toJsonInfo.Set "senderContactName", senderContactName  
    toJsonInfo.Set "senderDeptName", senderDeptName     
    toJsonInfo.Set "senderTEL", senderTEL          
    toJsonInfo.Set "senderHP", senderHP           
    toJsonInfo.Set "senderEmail", senderEmail        
    toJsonInfo.Set "senderFAX", senderFAX          

    toJsonInfo.Set "receiverCorpNum", receiverCorpNum    
    toJsonInfo.Set "receiverTaxRegID", receiverTaxRegID   
    toJsonInfo.Set "receiverCorpName", receiverCorpName   
    toJsonInfo.Set "receiverCEOName", receiverCEOName    
    toJsonInfo.Set "receiverAddr", receiverAddr       
    toJsonInfo.Set "receiverBizClass", receiverBizClass   
    toJsonInfo.Set "receiverBizType", receiverBizType    
    toJsonInfo.Set "receiverContactName", receiverContactName
    toJsonInfo.Set "receiverDeptName", receiverDeptName   
    toJsonInfo.Set "receiverTEL", receiverTEL        
    toJsonInfo.Set "receiverHP", receiverHP         
    toJsonInfo.Set "receiverEmail", receiverEmail      
    toJsonInfo.Set "receiverFAX", receiverFAX        

    toJsonInfo.Set "taxTotal", taxTotal           
    toJsonInfo.Set "supplyCostTotal", supplyCostTotal    
    toJsonInfo.Set "totalAmount", totalAmount        
    toJsonInfo.Set "purposeType", purposeType        
    toJsonInfo.Set "serialNum", serialNum          
    toJsonInfo.Set "remark1", remark1            
    toJsonInfo.Set "remark2", remark2            
    toJsonInfo.Set "remark3", remark3            
    toJsonInfo.Set "businessLicenseYN", businessLicenseYN
    toJsonInfo.Set "bankBookYN", bankBookYN         
    toJsonInfo.Set "faxsendYN", faxsendYN          
    toJsonInfo.Set "smssendYN", smssendYN          
    toJsonInfo.Set "autoacceptYN", autoacceptYN    	

    Dim detailJsonInfo()
    ReDim detailJsonInfo(UBound(detailList))
    Dim i, detail
    i = 0
    For Each detail In detailList
        Set detailJsonInfo(i) = detailList(i).toJsonInfo
        i = i + 1
    Next
    toJsonInfo.Set "detailList", detailJsonInfo
    toJsonInfo.Set "propertyBag", propertyBag
End Function

Public Sub fromJsonInfo(jsonInfo)
    On Error Resume Next

    itemCode = jsonInfo.itemCode
    mgtKey = jsonInfo.mgtKey               
    invoiceNum = jsonInfo.invoiceNum           
    formCode = jsonInfo.formCode             
    writeDate = jsonInfo.writeDate             
    taxType = jsonInfo.taxType               

    senderCorpNum = jsonInfo.senderCorpNum
    senderTaxRegID = jsonInfo.senderTaxRegID
    senderCorpName = jsonInfo.senderCorpName     
    senderCEOName = jsonInfo.senderCEOName      
    senderAddr = jsonInfo.senderAddr         
    senderBizClass = jsonInfo.senderBizClass     
    senderBizType = jsonInfo.senderBizType      
    senderContactName = jsonInfo.senderContactName  

    senderDeptName = jsonInfo.senderDeptName     
    senderTEL = jsonInfo.senderTEL         
    senderHP = jsonInfo.senderHP           
    senderEmail = jsonInfo.senderEmail        
    senderFAX = jsonInfo.senderFAX          

    receiverCorpNum = jsonInfo.receiverCorpNum    
    receiverTaxRegID = jsonInfo.receiverTaxRegID   
    receiverCorpName = jsonInfo.receiverCorpName   
    receiverCEOName = jsonInfo.receiverCEOName    
    receiverAddr = jsonInfo.receiverAddr       
    receiverBizClass = jsonInfo.receiverBizClass   
    receiverBizType = jsonInfo.receiverBizType    
    receiverContactName = jsonInfo.receiverContactName
    receiverDeptName = jsonInfo.receiverDeptName   
    receiverTEL = jsonInfo.receiverTEL        
    receiverHP = jsonInfo.receiverHP         
    receiverEmail = jsonInfo.receiverEmail      
    receiverFAX = jsonInfo.receiverFAX        

    taxTotal = jsonInfo.taxTotal           
    supplyCostTotal = jsonInfo.supplyCostTotal    
    totalAmount = jsonInfo.totalAmount        
    purposeType = jsonInfo.purposeType        
    serialNum = jsonInfo.serialNum          
    remark1 = jsonInfo.remark1            
    remark2 = jsonInfo.remark2            
    remark3 = jsonInfo.remark3            
    businessLicenseYN = jsonInfo.businessLicenseYN  
    bankBookYN = jsonInfo.bankBookYN         
    faxsendYN = jsonInfo.faxsendYN          
    smssendYN = jsonInfo.smssendYN          
    autoacceptYN = jsonInfo.autoacceptYN       

    ReDim detailList(jsonInfo.detailList.length)
    Dim i
    For i = 0 To jsonInfo.detailList.length-1
        Dim tmpDetail : Set tmpDetail = New StatementDetail
        tmpDetail.fromJsonInfo jsonInfo.detailList.Get(i)
        Set detailList(i) = tmpDetail
    Next

    Set propertyBag = jsonInfo.propertyBag
     
    On Error GoTo 0 

End Sub
End Class

Class StatementDetail
Public serialNum
Public purchaseDT
Public itemName
Public spec
Public unit
Public qty
Public unitCost
Public supplyCost
Public tax
Public remark
Public spare1
Public spare2
Public spare3
Public spare4
Public spare5
Public spare6
Public spare7
Public spare8
Public spare9
Public spare10
Public spare11
Public spare12
Public spare13
Public spare14
Public spare15
Public spare16
Public spare17
Public spare18
Public spare19
Public spare20

Public Function toJsonInfo()
    Set toJsonInfo = JSON.parse("{}")
    toJsonInfo.Set "serialNum", serialNum
    toJsonInfo.Set "purchaseDT", purchaseDT
    toJsonInfo.Set "itemName",itemName
    toJsonInfo.Set "spec", spec
    toJsonInfo.Set "unit", unit
    toJsonInfo.Set "qty", qty
    toJsonInfo.Set "unitCost", unitCost
    toJsonInfo.Set "supplyCost", supplyCost
    toJsonInfo.Set "tax", tax
    toJsonInfo.Set "remark", remark
    toJsonInfo.Set "spare1", spare1
    toJsonInfo.Set "spare2", spare2
    toJsonInfo.Set "spare3", spare3
    toJsonInfo.Set "spare4", spare4
    toJsonInfo.Set "spare5", spare5
    toJsonInfo.Set "spare6", spare6
    toJsonInfo.Set "spare7", spare7
    toJsonInfo.Set "spare8", spare8
    toJsonInfo.Set "spare9", spare9
    toJsonInfo.Set "spare10", spare10
    toJsonInfo.Set "spare11", spare11
    toJsonInfo.Set "spare12", spare12
    toJsonInfo.Set "spare13", spare13
    toJsonInfo.Set "spare14", spare14
    toJsonInfo.Set "spare15", spare15
    toJsonInfo.Set "spare16", spare16
    toJsonInfo.Set "spare17", spare17
    toJsonInfo.Set "spare18", spare18
    toJsonInfo.Set "spare19", spare19
    toJsonInfo.Set "spare20", spare20


End Function

Public Sub fromJsonInfo(jsonInfo)
    On Error Resume Next
    serialNum = jsonInfo.serialNum
    purchaseDT = jsonInfo.purchaseDT
    itemName = jsonInfo.itemName
    spec = jsonInfo.spec
    unit = jsonInfo.unit
    qty = jsonInfo.qty
    unitCost = jsonInfo.unitCost
    supplyCost = jsonInfo.supplyCost
    tax = jsonInfo.tax
    remark = jsonInfo.remark
    spare1 = jsonInfo.spare1
    spare2 = jsonInfo.spare2
    spare3 = jsonInfo.spare3
    spare4 = jsonInfo.spare4
    spare5 = jsonInfo.spare5
    spare6 = jsonInfo.spare6
    spare7 = jsonInfo.spare7
    spare8 = jsonInfo.spare8
    spare9 = jsonInfo.spare9
    spare10 = jsonInfo.spare10
    spare11 = jsonInfo.spare11
    spare12 = jsonInfo.spare12
    spare13 = jsonInfo.spare13
    spare14 = jsonInfo.spare14
    spare15 = jsonInfo.spare15
    spare16 = jsonInfo.spare16
    spare17 = jsonInfo.spare17
    spare18 = jsonInfo.spare18
    spare19 = jsonInfo.spare19
    spare20 = jsonInfo.spare20
    On Error GoTo 0 
End Sub
End Class

Class StmtSearchResult
    Public code
    Public total
    Public perPage
    Public pageNum
    Public pageCount
    Public message
    Public list()

    Public Sub Class_Initialize
        ReDim list(-1)
    End Sub

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
        code = jsonInfo.code
        total = jsonInfo.total
        perPage = jsonInfo.perPage
        pageNum = jsonInfo.pageNum
        pageCount = jsonInfo.pageCount
        message = jsonInfo.message
        
        ReDim list(jsonInfo.list.length)
        Dim i
        For i = 0 To jsonInfo.list.length -1
            Dim tmpObj : Set tmpObj = New StatementInfo
            tmpObj.fromJsonInfo jsonInfo.list.Get(i)
            Set list(i) = tmpObj
        Next

        On Error GoTo 0
    End Sub
End Class

Class EmailSendConfig
    Public emailType
    Public sendYN

    Public Sub fromJsonInfo(jsonInfo)
        On Error Resume Next
        emailType = jsonInfo.emailType
        sendYN = jsonInfo.sendYN
        On Error GoTo 0 
    End Sub 

    Public Function toJsonInfo()
        Set toJsonInfo = JSON.parse("{}")
        toJsonInfo.Set "emailType", emailType
        toJsonInfo.Set "sendYN", sendYN
    End Function 
End Class
%>
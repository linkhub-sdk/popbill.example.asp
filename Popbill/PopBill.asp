<!--#include file="Linkhub/Linkhub.asp"--> 
<%

Application("LINKHUB_TOKEN_SCOPE_POPBILL") = Array("member")
Const ServiceID_REAL = "POPBILL"
Const ServiceID_TEST = "POPBILL_TEST"
Const ServiceURL_REAL = "https://popbill.linkhub.co.kr"
Const ServiceURL_TEST = "https://popbill-test.linkhub.co.kr"
Const APIVersion = "1.0"
Const adTypeBinary = 1
Const adTypeText = 2

Class PopbillBase

Private m_IsTest
Private m_TokenDic
Private m_Linkhub

'테스트 플래그
Public Property Let IsTest(ByVal value)
    m_IsTest = value
End Property


Public Sub Class_Initialize
	On Error Resume next
	If  Not(POPBILL_TOKEN_CACHE Is Nothing) Then
		Set m_TokenDic = POPBILL_TOKEN_CACHE
	Else
		Set m_TokenDic = server.CreateObject("Scripting.Dictionary")
	End If
	On Error GoTo 0

	If isEmpty( m_TokenDic) Then
		Set m_TokenDic = server.CreateObject("Scripting.Dictionary")
	End If
	
	m_IsTest = False
    Set m_Linkhub = New Linkhub

	
End Sub

Public Sub Class_Terminate
	Set m_Linkhub = Nothing 
End Sub 

Private Property Get m_scope
	m_scope = Application("LINKHUB_TOKEN_SCOPE_POPBILL")
End Property

Public Sub AddScope(scope)
	t = Application("LINKHUB_TOKEN_SCOPE_POPBILL")
	ReDim Preserve t(Ubound(t)+1)
	t(Ubound(t)) = scope
	Application("LINKHUB_TOKEN_SCOPE_POPBILL") = t
End Sub


Public Sub Initialize(linkID, SecretKey )
    m_Linkhub.LinkID = linkID
    m_Linkhub.SecretKey = SecretKey
End Sub

Public Function getSession_token(CorpNum)
    refresh = False
    Set m_Token = Nothing
	
	If m_TokenDic.Exists(CorpNum) Then 
		Set m_Token = m_TokenDic.Item(CorpNum)
	End If
	
    If m_Token Is Nothing Then
        refresh = True
    Else
		'CheckScope
		For Each scope In m_scope
			If InStr(m_Token.strScope,scope) = 0 Then
				refresh = True
				Exit for
			End if
		Next
		If refresh = False then
			Dim utcnow
			utcnow = CDate(Replace(left(m_linkhub.getTime,19),"T" , " " ))
			refresh = CDate(Replace(left(m_Token.expiration,19),"T" , " " )) < utcnow
		End if
    End If
    
    If refresh Then
		If m_TokenDic.Exists(CorpNum) Then m_TokenDic.remove CorpNum
        Set m_Token = m_Linkhub.getToken(IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), CorpNum, m_scope)
		m_Token.set "strScope", Join(m_scope,"|")
		m_TokenDic.Add CorpNum, m_Token
	End If
    
    getSession_token = m_Token.session_token

End Function

'회원잔액조회
Public Function GetBalance(CorpNum)
    GetBalance = m_Linkhub.GetBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
End Function
'파트너 잔액조회
Public Function GetPartnerBalance(CorpNum)
    GetPartnerBalance = m_Linkhub.GetPartnerBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
End Function

'파트너 포인트 충전 URL - 2017/08/29 추가
Public Function GetPartnerURL(CorpNum, TOGO)
    GetPartnerURL = m_Linkhub.GetPartnerURL(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), TOGO)
End Function

'팝빌 기본 URL
Public Function GetPopbillURL(CorpNum , UserID , TOGO )

    Set result = httpGET("/?TG=" + TOGO, getSession_token(CorpNum), UserID)
    GetPopbillURL = result.url
End Function
'회원가입 여부
Public Function CheckIsMember(CorpNum , linkID)
    
    Set CheckIsMember = httpGET("/Join?CorpNum=" + CorpNum + "&LID=" + linkID, "","")

End Function
'회원가입
Public Function JoinMember(JoinInfo)
   
    Set tmp = JSON.parse("{}")
    
    tmp.set "LinkID", JoinInfo.linkID
    tmp.set "CorpNum", JoinInfo.CorpNum
    tmp.set "CorpName", JoinInfo.CorpName
    tmp.set "CEOName", JoinInfo.CEOName
    tmp.set "Addr", JoinInfo.Addr
    tmp.set "ZipCode", JoinInfo.ZipCode
    tmp.set "BizClass", JoinInfo.BizClass
    tmp.set "BizType", JoinInfo.BizType
    tmp.set "ContactName", JoinInfo.contactName
    tmp.set "ContactEmail", JoinInfo.ContactEmail
    tmp.set "ContactFAX", JoinInfo.ContactFAX
    tmp.set "ContactHP", JoinInfo.ContactHP
    tmp.set "ContactTEL", JoinInfo.ContactTEL
    tmp.set "ID", JoinInfo.ID
    tmp.set "PWD", JoinInfo.PWD
    
    postdata = m_Linkhub.toString(tmp)
   
    Set JoinMember = httpPOST("/Join", "", "", postdata, "")


End Function

' 담당자 목록조회
Public Function ListContact(CorpNum, UserID)

	Set result = httpGET("/IDs",getSession_token(CorpNum), UserID)

	Set infoObj = CreateObject("Scripting.Dictionary")

	For i = 0 To result.length - 1
		Set contInfo = New ContactInfo
		contInfo.fromJsonInfo result.Get(i)
		infoObj.Add i, contInfo
	Next

	Set ListContact = infoObj
End Function

'담당자 수정 
Public Function UpdateContact(CorpNum, ContactInfo, UserID)
	Set tmp = ContactInfo.toJsonInfo
	postdata = m_Linkhub.toString(tmp)

	Set UpdateContact = httpPOST("/IDs", getSession_token(CorpNum), "", postdata, UserID)
End Function

'담당자 추가
Public Function RegistContact(CorpNum, ContactInfo, UserId)
	Set tmp = ContactInfo.toJsonInfo
	postdata = m_Linkhub.toString(tmp)
	
	Set RegistContact = httpPOST("/IDs/New", getSession_token(CorpNum), "", postdata, UserId)
End Function 

'회사정보 확인
Public Function GetCorpInfo(CorpNum, UserID)
	Set result = httpGET("/CorpInfo",getSession_token(CorpNum), UserID)

	Set infoObj = New CorpInfo
	infoObj.fromJsonInfo result
	
	Set GetCorpInfo = infoObj
End Function

'회사정보 수정
Public Function UpdateCorpInfo(CorpNum, CorpInfo, UserID)
	Set tmp = CorpInfo.toJsonInfo
	postdata = m_Linkhub.toString(tmp)

	Set UpdateCorpInfo = httpPOST("/CorpInfo", getSession_token(CorpNum), "", postdata, UserId)
End Function

'아이디 중복확인
Public Function CheckID(id)
	Set CheckID = httpGET("/IDCheck?ID="+id, "", "")
End Function

'''''''''''''  End of PopbillBase

'Private Functions
Public Function httpGET(url , BearerToken , UserID )

    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("GET", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url, false)
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
	
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If

	winhttp1.Send
    winhttp1.WaitForResponse
	result = winhttp1.responseText

	If winhttp1.Status <> 200 Then
		Set winhttp1 = Nothing
        Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If
    
    Set winhttp1 = Nothing
    
    Set httpGET = m_Linkhub.parse(result)

End Function


Public Function httpPOST(url , BearerToken , override , postdata ,  UserID)
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If

    winhttp1.Send (postdata)
    winhttp1.WaitForResponse
    result = winhttp1.responseText
    
    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
		Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If
    
    Set winhttp1 = Nothing
    Set httpPOST = m_Linkhub.parse(result)

End Function


Public Function httpPOST_ContentsType(url , BearerToken , override , postdata , UserID, ContentsType)
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If

	If ContentsType <> "" Then
		Call winhttp1.setRequestHeader("Content-Type", ContentsType)
    Else
		Call winhttp1.setRequestHeader("Content-Type", "Application/json")
	End If

    winhttp1.Send (postdata)
    winhttp1.WaitForResponse
    result = winhttp1.responseText
    
    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
		Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If
    
    Set winhttp1 = Nothing
    Set httpPOST_ContentsType = m_Linkhub.parse(result)
End Function



Public Function httpPOST_File(url , BearerToken , FilePath , UserID )
     
    boundary = "---------------------popbill"
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary)
    
    Dim Stream : Set Stream = Server.CreateObject("ADODB.Stream")
	Stream.Type = adTypeBinary
	Stream.Open
	
    fileHead = vbCrLf & "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""Filedata""; filename=""" & GetOnlyFileName(FilePath) & """" + vbCrLf & _
           "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
	Stream.Write StringToBytes(fileHead)
	Stream.Write GetFile(FilePath)
           
    tail = vbCrLf & "--" & boundary & "--" & vbCrLf
	Stream.Write  StringToBytes(tail)

	Stream.Flush
	Stream.Position = 0
	postData = Stream.Read
	Stream.Close : Set Stream = Nothing
	
    winhttp1.Send (postData)
	winhttp1.WaitForResponse
    
    result = winhttp1.responseText
       
    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
		Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_File = m_Linkhub.parse(Result)

End Function


Public Function httpPOST_Files(url , BearerToken ,postData, FilePaths , UserID )
     
    boundary = "---------------------popbill"
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary)
    
    Dim Stream : Set Stream = Server.CreateObject("ADODB.Stream")
	Stream.Type = adTypeBinary
	Stream.Open
	
	If postdata <> "" Then
        applicationform = vbCrLf & "--" & boundary & vbCrLf & _
                          "Content-Disposition: form-data; name=""form""" & vbCrLf & _
                          "Content-Type: Application/json" & vbCrLf & vbCrLf & _
						  postdata
        Stream.Write StringToBytes(applicationform)
    End If
	
	For Each FilePath In FilePaths
		fileHead = vbCrLf & "--" & boundary & vbCrLf & _
			   "Content-Disposition: form-data; name=""file""; filename=""" & GetOnlyFileName(FilePath) & """" + vbCrLf & _
			   "Content-Type: application/octet-stream" & vbCrLf & vbCrLf

		Stream.Write StringToBytes(fileHead)
		Stream.Write GetFile(FilePath)
    Next
    
    tail = vbCrLf & "--" & boundary & "--" & vbCrLf
	Stream.Write  StringToBytes(tail)

	Stream.Flush
	Stream.Position = 0
	btPostData = Stream.Read
	Stream.Close : Set Stream = Nothing
	
    winhttp1.Send (btPostData)
	winhttp1.WaitForResponse
    
    result = winhttp1.responseText
       
    If winhttp1.Status <> 200 Then
        Set winhttp1 = Nothing
		Set parsedDic = m_Linkhub.parse(result)
        Err.raise parsedDic.code, "POPBILL", parsedDic.message
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_Files = m_Linkhub.parse(Result)

End Function

Private Function StringToBytes(Str)
  Dim Stream : Set Stream = Server.CreateObject("ADODB.Stream")
  Stream.Type = adTypeText
  Stream.Charset = "UTF-8"
  Stream.Open
  Stream.WriteText Str
  Stream.Flush
  Stream.Position = 0
  Stream.Type = adTypeBinary
  buffer= Stream.Read
  Stream.Close
  'Remove BOM.
  Set Stream = Server.CreateObject("ADODB.Stream")
  Stream.Type = adTypeBinary
  Stream.Open
  Stream.write buffer
  Stream.Flush
  Stream.Position = 3
  StringToBytes= Stream.Read
  Stream.Close
  Set Stream = Nothing
 
End Function

Private Function GetFile(FileName)
	Dim Stream: Set Stream = CreateObject("ADODB.Stream")
	Stream.Type = adTypeBinary
	Stream.Open
	Stream.LoadFromFile FileName
	GetFile = Stream.Read
	Stream.Close
End Function

Private Function GetOnlyFileName(ByVal FilePath ) 
     Temp = Split(FilePath, "\")
     GetOnlyFileName = Split(FilePath, "\")(UBound(Temp))
End Function

Private Function IIf(condition , trueState,falseState)
	If condition Then 
		IIf = trueState
	Else
		IIf = falseState
	End if
End Function
public Function toString(object)
	toString = m_Linkhub.toString(object)
End Function

Public Function parse(jsonString)
	Set parse = m_Linkhub.parse(jsonString)
End Function
End Class

''회원가입 정보
Class JoinForm
Public LinkID          
Public CorpNum         
Public CEOName         
Public CorpName        
Public Addr            
Public ZipCode         
Public BizType         
Public BizClass        
Public ID              
Public PWD             
Public ContactName     
Public ContactTEL      
Public ContactHP       
Public ContactFAX      
Public ContactEmail    
End Class

'담당자 정보
Class ContactInfo
	Public id
	Public pwd
	Public email
	Public hp
	Public personName
	Public searchAllAllowYN
	Public tel
	Public fax
	Public mgrYN
	Public regDT
	Public state
	
	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			
		id = jsonInfo.id
		email = jsonInfo.email
		hp = jsonInfo.hp
		personName = jsonInfo.personName
		searchAllAllowYN = jsonInfo.searchAllAllowYN
		tel = jsonInfo.tel
		fax = jsonInfo.fax
		mgrYN = jsonInfo.mgrYN
		regDT = jsonInfo.regDT
		State = jsonInfo.state

		On Error GoTo 0
	End Sub

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.set "id", id
		toJsonInfo.set "pwd", pwd
		toJsonInfo.set "email", email
		toJsonInfo.set "hp", hp
		toJsonInfo.set "personName", personName
		toJsonInfo.set "searchAllAllowYN", searchAllAllowYN
		toJsonInfo.set "tel", tel
		toJsonInfo.set "fax", fax
		toJsonInfo.set "mgrYN", mgrYN
	End Function

End Class

'회사정보 
Class CorpInfo
	Public ceoname
	Public corpName
	Public addr
	Public bizType
	Public bizClass

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
		ceoname = jsonInfo.ceoname
		corpName = jsonInfo.corpName
		addr = jsonInfo.addr
		bizType = jsonInfo.bizType
		bizClass = jsonInfo.bizClass
		On Error GoTo 0
	End Sub

	Public Function toJsonINfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "ceoname", ceoname
		toJsonInfo.Set "corpName", corpName
		toJsonInfo.Set "addr", addr
		toJsonInfo.Set "bizType", bizType
		toJsonInfo.Set "bizClass", bizClass
	End Function

End Class

'과금정보
Class ChargeInfo
	Public unitCost
	Public chargeMethod
	Public rateSystem

	Public Sub fromJsonInfo ( jsonInfo )
		On Error Resume Next
		unitCost = jsonInfo.unitCost
		chargeMethod = jsonInfo.chargeMethod
		rateSystem = jsonInfo.rateSystem
		On Error GoTo 0
	End Sub 

End Class
%>
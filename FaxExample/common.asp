<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/FaxService.asp"-->
<%
	'**************************************************************
	' 팝빌 팩스 API ASP SDK Example
	'
	' - ASP SDK 연동환경 설정방법 안내 : http://blog.linkhub.co.kr/577
	' - 업데이트 일자 : 2017-05-24
	' - 연동 기술지원 연락처 : 1600-8536 / 070-4304-2991
	' - 연동 기술지원 이메일 : code@linkhub.co.kr
	'
	' <테스트 연동개발 준비사항>
	' 1) 19, 22번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
	'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
	' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
	'**************************************************************

	'링크아이디 
	LinkID = "TESTER"

	'비밀키
	SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_FaxService = new FaxService

	m_FaxService.Initialize LinkID, SecretKey

	'연동환경설정값, 개발용(True), 상업용(False)
	m_FaxService.IsTest = True
%>
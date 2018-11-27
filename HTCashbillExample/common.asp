<!--#include virtual="/Popbill/Popbill.asp"--> 
<!--#include virtual="/Popbill/HTCashbillService.asp"-->
<%

	'**************************************************************'
	' 팝빌 홈택스 현금영수증 연동 API ASP SDK Example
	'
	' - ASP SDK 연동환경 설정방법 안내 : http://blog.linkhub.co.kr/577/
	' - 업데이트 일자 : 2018-11-22
	' - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991
	' - 연동 기술지원 이메일 : code@linkhub.co.kr
	'
	' <테스트 연동개발 준비사항>
	' 1) 24, 27번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
	'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
	' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
	' 3) 홈택스에서 이용가능한 공인인증서를 등록합니다.
	'    - 팝빌로그인 > [홈택스연계] > [환경설정] > [공인인증서 관리] 메뉴
	'    - 공인인증서 등록(GetCertificatePopUpURL API) 반환된 URL을 이용하여
	'      팝업 페이지에서 공인인증서 등록
	'**************************************************************

	'링크아이디 
	LinkID = "TESTER"

	'연동상담시 발급받은 비밀키, 유출에 주의
	SecretKey ="SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

	set m_HTCashbillService = new HTCashbillService

	m_HTCashbillService.Initialize LinkID, SecretKey

	'연동환경 설정값 개발용(True), 상업용(False)
	m_HTCashbillService.IsTest = True
%>
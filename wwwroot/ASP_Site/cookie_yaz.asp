<%
response.Cookies("Ziyaret�i")("Adi")="Bedri"
response.Cookies("Ziyaret�i")("Soyadi")="Akay"
response.Cookies("Ziyaret�i")("Email")="Bedri@yasalegitim.com"
response.Cookies("Ziyaret�i")("yasi")=25
response.Cookies("Font")="Arial"
response.Cookies("�ablon")="�ablon 1"
response.Cookies("Ad�")="Bedrettin" 


response.Cookies("Ziyaret�i").Expires = "30/01/2006"
response.Cookies("Font").Expires = "30/01/2006 10:47:00"
response.Cookies("�ablon").expires="30/01/2006"
response.Cookies("Ad�").Expires = "30/01/2006"
response.Cookies("Ad�").domain = "yasalegitim.com"
response.Cookies("Ad�").path = "/data/"
response.Cookies("Ad�").secure = true
%>
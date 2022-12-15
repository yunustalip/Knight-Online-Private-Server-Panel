<!--#include file="tarayicitanima.asp"--><%
dim zmn
SUB AktifKullanici
  DIM strAktifKullaniciListesi
  DIM intKullaniciBaslangic, intKullaniciBitis
  DIM strKullanici
  DIM strTarih

  strAktifKullaniciListesi = APPLICATION("AktifKullaniciListesi")

  IF Instr(1, strAktifKullaniciListesi, Session.SessionID) > 0 Then
    Application.LOCK
    intKullaniciBaslangic = INSTR(1, strAktifKullaniciListesi, Session.SessionID)
    intKullaniciBitis = INSTR(intKullaniciBaslangic, strAktifKullaniciListesi, "|")
    strKullanici = MID(strAktifKullaniciListesi, intKullaniciBaslangic, intKullaniciBitis - intKullaniciBaslangic)
    strAktifKullaniciListesi = REPLACE(strAktifKullaniciListesi, strKullanici, Session.SessionID & "#" & NOW()&"#"&Request.ServerVariables("REMOTE_ADDR")& "#"& Session("sayfa")&"#"&BrowserType)
    Application("AktifKullaniciListesi") = strAktifKullaniciListesi
    Application.UNLOCK
  ELSE
    Application.LOCK
    Application("AktifKullaniciListesi") = Application("AktifKullaniciListesi") & Session.SessionID & "#" & NOW()&"#"&Request.ServerVariables("REMOTE_ADDR") & "#" &Session("sayfa")&"#"&BrowserType &"|"
    Application.UNLOCK
	End If
	END SUB

	SUB AktifKullanicilariSil
 	 DIM ix
 	 DIM strAktifKullaniciListesi
 	 DIM aAktifKullanicilar
 	 DIM intAktifKullaniciSilmeZamani
 	 DIM intAktifKullaniciTimeout

  intAktifKullaniciSilmeZamani = 1 'dakika olarak AktifKullaniciListe'sinin ne kadar zamanda bir silineceði.
  intAktifKullaniciTimeout = 1 'dakika olarak, kullanýcý ne zaman hareketsiz kabul edilecek ve silinecek

  IF Application("AktifKullaniciListesi") = "" Then EXIT SUB

  IF DATEDIFF("n", Application("AktifKullaniciSonSilme"), NOW()) > intAktifKullaniciSilmeZamani Then
    Application.LOCK
    Application("AktifKullaniciSonSilme") = NOW()
    Application.UNLOCK

    strAktifKullaniciListesi = Application("AktifKullaniciListesi")
    strAktifKullaniciListesi = LEFT(strAktifKullaniciListesi, LEN(strAktifKullaniciListesi) - 1)

    aAktifKullanicilar = SPLIT(strAktifKullaniciListesi, "|")

    FOR ix = 0 TO UBOUND(aAktifKullanicilar)
    zmn=split(aAktifKullanicilar(ix),"#")
    IF DATEDIFF("n", MID(aAktifKullanicilar(ix), INSTR(1, aAktifKullanicilar(ix), "#")+1, len(zmn(1))), NOW()) > intAktifKullaniciTimeout Then
        aAktifKullanicilar(ix) = "XXXX"
      End If
    NEXT

    strAktifKullaniciListesi = JOIN(aAktifKullanicilar, "|") & "|"
    strAktifKullaniciListesi = REPLACE(strAktifKullaniciListesi, "XXXX|", "")

    Application.LOCK
    Application("AktifKullaniciListesi") = strAktifKullaniciListesi
    Application.UNLOCK

  End If
END SUB
%>
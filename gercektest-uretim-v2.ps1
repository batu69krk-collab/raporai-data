$ErrorActionPreference = 'Stop'
$outPath  = Join-Path $PSScriptRoot 'rapor-korpus.html'
$cs = 'Data Source=localhost\BOTANIKSQL;Initial Catalog=eczane;Integrated Security=True;Connect Timeout=30'

$conn = New-Object System.Data.SqlClient.SqlConnection($cs)
$conn.Open()

function Fill([string]$q) {
    $da = New-Object System.Data.SqlClient.SqlDataAdapter($q, $conn)
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    return , $dt
}

# ---- V3: kesif INFORMATION_SCHEMA ile dogrulanmis sabit isimler ----
# Hasta = Musteri (MusteriId / MusteriDogumTarihi); FK: RaporAnaMusteriId, RxMusteriId, EReceteMusteriId
# Barkod = Barkod.BarkodAdi (BarkodUrunId) + FIBBarkodUrunId.FBBarkod (FBUrunId)
# Kamu   = Urun.UrunSGKKodId -> SGKKodlari.SGKKodu
$chk = Fill "SELECT TABLE_NAME, COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE (TABLE_NAME='Musteri' AND COLUMN_NAME IN ('MusteriId','MusteriDogumTarihi')) OR (TABLE_NAME='RaporAna' AND COLUMN_NAME='RaporAnaMusteriId') OR (TABLE_NAME='ReceteAna' AND COLUMN_NAME='RxMusteriId') OR (TABLE_NAME='ERecete' AND COLUMN_NAME='EReceteMusteriId')"
if ($chk.Rows.Count -lt 5) { throw ('V3 kesif hatasi: beklenen kolonlar bulunamadi (bulunan=' + $chk.Rows.Count + '/5)') }
$hastaTablo   = 'Musteri'
$hastaIdKolon = 'MusteriId'
$dogumKolon   = 'MusteriDogumTarihi'
Write-Host 'V3 kesif: hasta=Musteri (MusteriId, MusteriDogumTarihi); FK=RaporAnaMusteriId/RxMusteriId/EReceteMusteriId dogrulandi'

# Hasta dogum tarihleri YALNIZ bellegde tutulur (ciktiya DOB yazilmaz; yalniz yas turetilir)
$hastaDogum = @{}
if ($hastaTablo -and $hastaIdKolon -and $dogumKolon) {
    foreach ($r in (Fill ("SELECT " + $hastaIdKolon + " AS Hid, " + $dogumKolon + " AS Dob FROM " + $hastaTablo + " WHERE " + $dogumKolon + " IS NOT NULL"))) {
        try { $hastaDogum[[int]$r.Hid] = [datetime]$r.Dob } catch { }
    }
}
function YasYaz([object]$dob, [object]$tarih) {
    if (-not $dob -or -not $tarih) { return '' }
    try {
        $d = [datetime]$tarih; $b = [datetime]$dob
        $yas = $d.Year - $b.Year
        if ($d.Month -lt $b.Month -or ($d.Month -eq $b.Month -and $d.Day -lt $b.Day)) { $yas-- }
        if ($yas -lt 0 -or $yas -gt 130) { return '' }
        return [string]$yas
    } catch { return '' }
}

$inv = [System.Globalization.CultureInfo]::InvariantCulture
function Num($x) { if ($null -eq $x -or $x -eq [System.DBNull]::Value) { '' } else { ([double]$x).ToString($inv) } }
function Dte($x) { if ($null -eq $x -or $x -eq [System.DBNull]::Value) { '' } else { ([datetime]$x).ToString('yyyy-MM-dd') } }
function Dts($x) { if ($null -eq $x -or $x -eq [System.DBNull]::Value) { '' } else { ([datetime]$x).ToString('yyyy-MM-dd HH:mm') } }
function Mask([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return $s }
    $s = [regex]::Replace($s, '(?i)(ad\s*ve\s*soyad|adi\s*soyadi|soyadi|anne\s*adi|baba\s*adi)\s*:?\s*[^\n,;.]{2,60}', '$1: [KIMLIK-SILENDI]')
    $s = [regex]::Replace($s, '(?i)dogum\s*tarihi\s*:?\s*[\d./-]+', 'dogum tarihi: [KIMLIK-SILENDI]')
    $s = [regex]::Replace($s, '(?<!\d)\d{11}(?!\d)', '[MASKELI-11HANE]')
    $s = [regex]::Replace($s, '(?i)((?:sgk\s+)?(?:rapor\s+)?tak.p\s*nos?u?)\s*:?\s*(\d{5,15})', { param($m) ($m.Groups[1].Value + ': ' + (Pseudo $m.Groups[2].Value)) })
    return $s
}
function Esc([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
    return $s
}
function Clean([string]$s) { if ($null -eq $s) { '' } else { Esc (Mask $s) } }
function Flag($x) { if ($x -eq [System.DBNull]::Value -or $null -eq $x) { '' } else { if ($x) { 'evet' } else { 'hayir' } } }

# Pseudonimleme: ayni gercek numara -> her yerde ayni sahte deger (salted SHA256, on only)
$psalt = 'rapor-korpus|v3|b3a9f7e1d5c24a80'
$sha256 = [System.Security.Cryptography.SHA256]::Create()
function Pseudo([object]$x) {
    $s = [string]$x
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    $b = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($psalt + '|' + $s.Trim()))
    return 'PSD-' + (($b[0..5] | ForEach-Object { $_.ToString('x2') }) -join '')
}

function AddGroup($ht, $k, $v) {
    if (-not $ht.ContainsKey($k)) { $ht[$k] = New-Object System.Collections.Generic.List[string] }
    [void]$ht[$k].Add($v)
}

Write-Host 'Referans tablolar yukleniyor...'
$icdMap  = @{}; foreach ($r in (Fill 'SELECT ICDId, ICDKodu, ICDAciklamasi FROM ICD'))  { $icdMap[[int]$r.ICDId]  = @("$($r.ICDKodu)", "$($r.ICDAciklamasi)") }
$emMap   = @{}; foreach ($r in (Fill 'SELECT EtkinMaddeId, EtkinMaddeAdi FROM EtkinMadde')) { $emMap[[int]$r.EtkinMaddeId] = "$($r.EtkinMaddeAdi)" }
$kulMap  = @{}; foreach ($r in (Fill 'SELECT KullanimId, KullanimSekli FROM Kullanim')) { $kulMap[[int]$r.KullanimId] = "$($r.KullanimSekli)" }
$bransMap= @{}; foreach ($r in (Fill 'SELECT BransId, BransAdi FROM Brans')) { $bransMap[[int]$r.BransId] = "$($r.BransAdi)" }
$rtMap   = @{}; foreach ($r in (Fill 'SELECT RaporTuruId, RaporTuruAdi FROM RaporTuru')) { $rtMap[[int]$r.RaporTuruId] = "$($r.RaporTuruAdi)" }
$kurumMap= @{}; $gi=0; foreach ($r in (Fill 'SELECT KurumId, KurumAdi FROM Kurum ORDER BY KurumAdi')) { $gi++; $kurumMap[[int]$r.KurumId] = ('KURUM-' + $gi) }
$hasMap  = @{}; $hi=0; foreach ($r in (Fill 'SELECT HastaneId, HastaneAdi FROM Hastane ORDER BY HastaneAdi')) { $hi++; $hasMap[[int]$r.HastaneId] = ('HASTANE-' + $hi) }
$rkMap   = @{}; foreach ($r in (Fill 'SELECT RaporKodId, RaporKodu, RaporKodAciklama FROM RaporKodlari')) { $rkMap[[int]$r.RaporKodId] = @("$($r.RaporKodu)", "$($r.RaporKodAciklama)") }
$altMap  = @{}; foreach ($r in (Fill 'SELECT ReceteAltTuruId, ReceteAltTuruAdi FROM ReceteAltTuru')) { $altMap[[int]$r.ReceteAltTuruId] = "$($r.ReceteAltTuruAdi)" }
$renkMap = @{}; foreach ($r in (Fill 'SELECT ReceteRenkId, ReceteRenkAdi FROM ReceteRenk')) { $renkMap[[int]$r.ReceteRenkId] = "$($r.ReceteRenkAdi)" }
$prvMap  = @{}; foreach ($r in (Fill 'SELECT ProvizyonTipId, ProvizyonTipAdi FROM ProvizyonTipi')) { $prvMap[[int]$r.ProvizyonTipId] = "$($r.ProvizyonTipAdi)" }
$kgMap   = @{}; foreach ($r in (Fill 'SELECT KogusId, KogusAdi FROM Kogus')) { $kgMap[[int]$r.KogusId] = "$($r.KogusAdi)" }
$serMap  = @{}; foreach ($r in (Fill 'SELECT SertifikaId, SertifikaAdi FROM Sertifika')) { $serMap[[int]$r.SertifikaId] = "$($r.SertifikaAdi)" }
$eakMap  = @{}; foreach ($r in (Fill 'SELECT EReceteAciklamaId, EReceteAciklamaAdi FROM EReceteAciklama')) { $eakMap[[int]$r.EReceteAciklamaId] = "$($r.EReceteAciklamaAdi)" }
$eatMap  = @{}; foreach ($r in (Fill 'SELECT EReceteAciklamaTuruId, EReceteAciklamaTuruAdi FROM EReceteAciklamaTuru')) { $eatMap[[int]$r.EReceteAciklamaTuruId] = "$($r.EReceteAciklamaTuruAdi)" }
$umbMap  = @{}; foreach ($r in (Fill 'SELECT UMTBransId, UMTBransAdi FROM UMTBrans')) { $umbMap[[int]$r.UMTBransId] = "$($r.UMTBransAdi)" }
$umrMap  = @{}; foreach ($r in (Fill 'SELECT UMTRaporId, UMTRaporKod, UMTRaporAciklama FROM UMTRapor')) { $umrMap[[int]$r.UMTRaporId] = @("$($r.UMTRaporKod)", "$($r.UMTRaporAciklama)") }
$umtMap  = @{}; foreach ($r in (Fill 'SELECT UMTReceteTeshisId, UMTReceteTeshisKod, UMTReceteTeshisAciklama FROM UMTReceteTeshis')) { $umtMap[[int]$r.UMTReceteTeshisId] = @("$($r.UMTReceteTeshisKod)", "$($r.UMTReceteTeshisAciklama)") }
$utsMap  = @{}; foreach ($r in (Fill 'SELECT UMTTesisId, UMTTesisAdi FROM UMTTesis')) { $utsMap[[int]$r.UMTTesisId] = "$($r.UMTTesisAdi)" }
$urunMap = @{}
$urunBarkodMap = @{}   # UrunId -> barkod (Barkod + FIBBarkodUrunId tablolarindan)
$urunKamuMap = @{}     # UrunId -> kamu kodu (Urun.UrunSGKKodId -> SGKKodlari.SGKKodu)
foreach ($r in (Fill 'SELECT UrunId, UrunAdi FROM Urun')) { $urunMap[[int]$r.UrunId] = "$($r.UrunAdi)" }
# Barkod haritasi cok kaynakli: Barkod(aktif) -> FIB -> FaturaSatir(FSBarkodAdi, 8-14 rakam) -> Karekod(GTIN) -> Barkod(silinmis fallback)
foreach ($r in (Fill 'SELECT BarkodUrunId, MIN(BarkodAdi) AS BarkodAdi FROM Barkod WHERE ISNULL(BarkodSilme,0)=0 AND BarkodAdi IS NOT NULL GROUP BY BarkodUrunId')) {
    $k = [int]$r.BarkodUrunId; $v = ([string]$r.BarkodAdi).Trim()
    if ($v -and -not $urunBarkodMap.ContainsKey($k)) { $urunBarkodMap[$k] = $v }
}
foreach ($r in (Fill 'SELECT FBUrunId, MIN(FBBarkod) AS FBBarkod FROM FIBBarkodUrunId WHERE FBBarkod IS NOT NULL GROUP BY FBUrunId')) {
    $k = [int]$r.FBUrunId; $v = ([string]$r.FBBarkod).Trim()
    if ($v -and -not $urunBarkodMap.ContainsKey($k)) { $urunBarkodMap[$k] = $v }
}
foreach ($r in (Fill "SELECT FSUrunId, MIN(FSBarkodAdi) AS b FROM FaturaSatir WHERE FSBarkodAdi IS NOT NULL AND FSBarkodAdi NOT LIKE '%[^0-9]%' AND LEN(FSBarkodAdi) BETWEEN 8 AND 14 GROUP BY FSUrunId")) {
    $k = [int]$r.FSUrunId; $v = ([string]$r.b).Trim()
    if ($v -and -not $urunBarkodMap.ContainsKey($k)) { $urunBarkodMap[$k] = $v }
}
foreach ($r in (Fill 'SELECT KKUrunId, MIN(GTIN) AS g FROM Karekod WHERE GTIN IS NOT NULL GROUP BY KKUrunId')) {
    if ($null -eq $r.KKUrunId -or $r.KKUrunId -eq [System.DBNull]::Value) { continue }
    $k = [int]$r.KKUrunId; $v = ([string]$r.g).Trim().TrimStart('0')
    if ($v -and -not $urunBarkodMap.ContainsKey($k)) { $urunBarkodMap[$k] = $v }
}
foreach ($r in (Fill 'SELECT BarkodUrunId, MIN(BarkodAdi) AS BarkodAdi FROM Barkod WHERE ISNULL(BarkodSilme,0)=1 AND BarkodAdi IS NOT NULL GROUP BY BarkodUrunId')) {
    $k = [int]$r.BarkodUrunId; $v = ([string]$r.BarkodAdi).Trim()
    if ($v -and -not $urunBarkodMap.ContainsKey($k)) { $urunBarkodMap[$k] = $v }
}
foreach ($r in (Fill 'SELECT u.UrunId AS Uid, s.SGKKodu AS Kod FROM Urun u INNER JOIN SGKKodlari s ON s.SGKKodId=u.UrunSGKKodId')) {
    $v = ([string]$r.Kod).Trim()
    if ($v -and -not $urunKamuMap.ContainsKey([int]$r.Uid)) { $urunKamuMap[[int]$r.Uid] = $v }
}
$ukeyDolu = 0; $ukeyBos = 0
Write-Host ("V3 kesif: urunBarkod dolu=" + $urunBarkodMap.Count + " urunKamu dolu=" + $urunKamuMap.Count)

Write-Host 'Rapor alt tablolari yukleniyor...'
$repIcd   = @{}
foreach ($r in (Fill 'SELECT RRKIRaporAnaId, RRKIICDId, RRKIICDId2, RRKIICDId3, RRKIICDId4, RRKIICDId5, RRKIRaporKodId, RRKIBaslamaTarihi, RRKIBitisTarihi FROM RaporRaporKodlariICD WHERE ISNULL(RRKISilme,0)=0')) {
    $parts = @()
    foreach ($icdId in @($r.RRKIICDId, $r.RRKIICDId2, $r.RRKIICDId3, $r.RRKIICDId4, $r.RRKIICDId5)) {
        if ($icdId -is [int] -and $icdMap.ContainsKey($icdId)) { $m = $icdMap[$icdId]; $parts += ($m[0] + ' ' + $m[1]) }
    }
    if ($r.RRKIRaporKodId -is [int] -and $rkMap.ContainsKey($r.RRKIRaporKodId)) { $k = $rkMap[$r.RRKIRaporKodId]; $parts += ('Rapor Kodu: ' + $k[0] + ' (' + $k[1] + ')') }
    $bt = Dte $r.RRKIBaslamaTarihi; $ft = Dte $r.RRKIBitisTarihi
    if ($bt -or $ft) { $parts += ('Tarih: ' + $bt + ' .. ' + $ft) }
    if ($parts.Count -gt 0) { AddGroup $repIcd ([int]$r.RRKIRaporAnaId) ($parts -join ' | ') }
}

$repEm = @{}
foreach ($r in (Fill 'SELECT e.EtkinMaddeRaporAnaId AS rid, m.EtkinMaddeAdi AS ad, e.EtkinMaddeDoz AS doz, e.EtkinMaddeAralik AS ar, e.EtkinMaddePeriyotId AS per, e.EtkinMaddeTekrar AS tek, k.KullanimSekli AS kul, e.EtkinMaddeAdetMiktar AS adet, e.EtkinMaddeIcerikMiktar AS icerik, b.BirimAdi AS birim, e.EtkinMaddeEklemeZamani AS ekz FROM RaporEtkinMadde e LEFT JOIN EtkinMadde m ON m.EtkinMaddeId=e.EtkinMaddeId LEFT JOIN Kullanim k ON k.KullanimId=e.EtkinMaddeKullanimId LEFT JOIN Birim b ON b.BirimId=e.EtkinMaddeBirimId WHERE ISNULL(e.EtkinMaddeSilme,0)=0')) {
    $s = "$($r.ad) | Doz: $(Num $r.doz) | Aralik: $($r.ar) | Periyot kodu: $($r.per) | Tekrar: $($r.tek)"
    if ($r.kul) { $s += " | Kullanim: $($r.kul)" }
    if ($r.adet) { $s += " | Adet: $($r.adet)" }
    if ($r.icerik) { $s += " | Icerik: $($r.icerik)" }
    if ($r.birim) { $s += " | Birim: $($r.birim)" }
    if ($r.ekz) { $s += " | Ekleme: $(Dts $r.ekz)" }
    AddGroup $repEm ([int]$r.rid) $s
}

$repBrans = @{}
foreach ($r in (Fill 'SELECT d.RaporDoktorRaporAnaId AS rid, b.BransAdi AS ad, d.RaporDoktorDoktorId AS dr FROM RaporDoktor d LEFT JOIN Brans b ON b.BransId=d.RaporDoktorBransId WHERE ISNULL(d.RaporDoktorSilme,0)=0')) {
    $k = [int]$r.rid
    if (-not $repBrans.ContainsKey($k)) { $repBrans[$k] = New-Object System.Collections.Generic.List[string] }
    $v = "$($r.ad) (Doktor Ref: $($r.dr))"
    if (-not $repBrans[$k].Contains($v)) { [void]$repBrans[$k].Add($v) }
}

$repEk = @{}
foreach ($r in (Fill 'SELECT REBRaporAnaId AS rid, REBTuru AS tur, REBDeger AS deg, REBAciklama AS aci, REBEklemeTarihi AS ekt FROM RaporEkBilgi')) {
    $s = "$($r.tur): $($r.deg)"
    if ($r.aci) { $s += " ($($r.aci))" }
    if ($r.ekt) { $s += " [ekleme: $(Dts $r.ekt)]" }
    AddGroup $repEk ([int]$r.rid) $s
}

Write-Host 'SGK recete alt tablolari yukleniyor...'
$rxIcd = @{}
foreach ($r in (Fill 'SELECT r.ReceteICDRxId AS rid, i.ICDKodu AS kod, i.ICDAciklamasi AS ad, r.ReceteICDICDAciklama AS not_ FROM ReceteICD r LEFT JOIN ICD i ON i.ICDId=r.ReceteICDICDId WHERE ISNULL(r.ReceteICDSilme,0)=0')) {
    $s = "$($r.kod) $($r.ad)"
    if ($r.not_ -and "$($r.not_)" -ne "$($r.ad)") { $s += " | Not: $(Clean $r.not_)" }
    AddGroup $rxIcd ([int]$r.rid) $s
}

$rxIlac = @{}
foreach ($r in (Fill 'SELECT r.RIRxId AS rid, r.RIUrunId AS urunId, u.UrunAdi AS ad, r.RIAdet AS adet, r.RIKKAdet AS kk, r.RIDoz AS doz, r.RIAralik AS ar, r.RIPeriyotId AS per, r.RITekrar AS tek, rk.RaporKodu AS rkod, r.RIBitisTarihi AS bit, r.RIOzellik AS oz, r.RIRaporNo AS rno, r.RIIade AS iade, r.RITekli AS tekli FROM ReceteIlaclari r LEFT JOIN Urun u ON u.UrunId=r.RIUrunId LEFT JOIN RaporKodlari rk ON rk.RaporKodId=r.RIRaporKodId WHERE ISNULL(r.RISilme,0)=0')) {
    $s = "$($r.ad) | Adet: $($r.adet)"
    if ($r.kk -is [int] -and $r.kk -gt 0) { $s += " | KK: $($r.kk)" }
    if ($r.doz -is [double] -and $r.doz -gt 0) { $s += " | Doz: $(Num $r.doz)/Aralik $($r.ar)/Periyot kodu $($r.per)" }
    if ($r.tek -is [int] -and $r.tek -gt 1) { $s += " | Tekrar: $($r.tek)" }
    if ($r.rkod) { $s += " | Rapor Kodu: $($r.rkod)" }
    if ($r.rno) { $s += " | Ilac Rapor No: $(Pseudo $r.rno)" }
    if ($r.bit) { $s += " | Bitis: $(Dte $r.bit)" }
    if ($r.iade -is [bool] -and $r.iade) { $s += ' | Iade' }
    if ($r.tekli -is [bool] -and $r.tekli) { $s += ' | Tekli' }
    if ($r.oz) { $s += " | $(Clean $r.oz)" }
    $uKey = $null
    $uId = $null
    if ([int]::TryParse([string]$r.urunId, [ref]$uId)) {
        if ($urunBarkodMap.ContainsKey($uId)) { $uKey = 'BARKOD:' + $urunBarkodMap[$uId] }
        elseif ($urunKamuMap.ContainsKey($uId)) { $uKey = 'KAMU:' + $urunKamuMap[$uId] }
    }
    if ($uKey) { $s += ' | UrunKey: ' + $uKey; $ukeyDolu++ } else { $ukeyBos++ }
    AddGroup $rxIlac ([int]$r.rid) $s
}

$rxTeshis = @{}
foreach ($r in (Fill 'SELECT RTRxId AS rid, RTTeshisKodu AS kod, RTSecilenIlacAdi AS ilac FROM ReceteTeshis')) {
    AddGroup $rxTeshis ([int]$r.rid) ("$($r.kod) -> $($r.ilac)")
}

$rxOzel = @{}
foreach ($r in (Fill 'SELECT RORxId AS rid, ROSiraliDagitim AS sirali, ROEmekli AS emekli, ROReceteGrubu AS grup, RORenklideVarmi AS renkli, ROSonlandirildi AS sonlandi, ROCetasFaturaId AS cfid, ROCetasIcmalId AS cid FROM ReceteOzel')) {
    $parts = @()
    if ($r.sirali -ne [System.DBNull]::Value) { $parts += "Sirali Dagitim: $($r.sirali)" }
    if ($r.emekli -is [bool]) { $parts += "Emekli: $(Flag $r.emekli)" }
    if ($r.grup -ne [System.DBNull]::Value) { $parts += "Recete Grubu: $($r.grup)" }
    if ($r.renkli -ne [System.DBNull]::Value) { $parts += "Renkli: $($r.renkli)" }
    if ($r.sonlandi -ne [System.DBNull]::Value) { $parts += "Sonlandirildi: $($r.sonlandi)" }
    if ($r.cfid) { $parts += "Cetas Fatura Ref: $($r.cfid)" }
    if ($r.cid) { $parts += "Cetas Icmal Ref: $($r.cid)" }
    if ($parts.Count -gt 0) { AddGroup $rxOzel ([int]$r.rid) ($parts -join ' | ') }
}

Write-Host 'E-recete alt tablolari yukleniyor...'
$erIcd = @{}
foreach ($r in (Fill 'SELECT r.EReceteICDEReceteId AS rid, i.ICDKodu AS kod, i.ICDAciklamasi AS ad FROM EReceteICD r LEFT JOIN ICD i ON i.ICDId=r.EReceteICDICDId WHERE ISNULL(r.EReceteICDSilme,0)=0')) {
    AddGroup $erIcd ([int]$r.rid) ("$($r.kod) $($r.ad)")
}

$erIlac = @{}
foreach ($r in (Fill 'SELECT r.ERIEReceteId AS rid, r.ERIUrunId AS urunId, r.ERIIlacAdi AS ad, r.ERIAdet AS adet, r.ERITekrar AS tek, r.ERIDoz AS doz, r.ERIAralik AS ar, r.ERIPeriyotId AS per, k.KullanimSekli AS kul, u.UrunAdi AS uad, r.ERIOdenme AS od FROM EReceteIlac r LEFT JOIN Kullanim k ON k.KullanimId=r.ERIKullanimId LEFT JOIN Urun u ON u.UrunId=r.ERIUrunId WHERE ISNULL(r.ERISilme,0)=0')) {
    $ad = $r.ad; if (-not $ad) { $ad = $r.uad }
    $s = "$ad | Adet: $($r.adet) | Doz: $(Num $r.doz)/Aralik $($r.ar)/Periyot kodu $($r.per) | Tekrar: $($r.tek)"
    if ($r.kul) { $s += " | Kullanim: $($r.kul)" }
    if ($r.od -ne [System.DBNull]::Value) { $s += " | Odeme Durumu: $($r.od)" }
    $uKey = $null
    $uId = $null
    if ([int]::TryParse([string]$r.urunId, [ref]$uId)) {
        if ($urunBarkodMap.ContainsKey($uId)) { $uKey = 'BARKOD:' + $urunBarkodMap[$uId] }
        elseif ($urunKamuMap.ContainsKey($uId)) { $uKey = 'KAMU:' + $urunKamuMap[$uId] }
    }
    if ($uKey) { $s += ' | UrunKey: ' + $uKey; $ukeyDolu++ } else { $ukeyBos++ }
    AddGroup $erIlac ([int]$r.rid) $s
}

$erNot = @{}
foreach ($r in (Fill 'SELECT a.ERAEReceteId AS rid, a.ERAEReceteAciklamaTuruId AS tur, a.ERAEReceteAciklamaId AS ack FROM EReceteAciklamalari a')) {
    $t = ''; if ($r.tur -is [int] -and $eatMap.ContainsKey($r.tur)) { $t = $eatMap[$r.tur] }
    $n = ''; if ($r.ack -is [int] -and $eakMap.ContainsKey($r.ack)) { $n = $eakMap[$r.ack] }
    AddGroup $erNot ([int]$r.rid) ("$t : $n")
}

Write-Host 'Elden ilaclari yukleniyor...'
$elIlac = @{}
foreach ($r in (Fill 'SELECT r.RIRxId AS rid, u.UrunAdi AS ad, r.RIAdet AS adet, r.RIOzellik AS oz, r.RIUrunId AS urunId FROM EldenIlaclari r LEFT JOIN Urun u ON u.UrunId=r.RIUrunId WHERE ISNULL(r.RISilme,0)=0')) {
    $s = "$($r.ad) | Adet: $($r.adet)"
    if ($r.oz) { $s += " | $(Clean $r.oz)" }
    $uKey = $null
    $uId = $null
    if ([int]::TryParse([string]$r.urunId, [ref]$uId)) {
        if ($urunBarkodMap.ContainsKey($uId)) { $uKey = 'BARKOD:' + $urunBarkodMap[$uId] }
        elseif ($urunKamuMap.ContainsKey($uId)) { $uKey = 'KAMU:' + $urunKamuMap[$uId] }
    }
    if ($uKey) { $s += ' | UrunKey: ' + $uKey; $ukeyDolu++ } else { $ukeyBos++ }
    AddGroup $elIlac ([int]$r.rid) $s
}

Write-Host 'SUT kural tablolari yukleniyor...'
$sutBrans = @{}
foreach ($r in (Fill 'SELECT UMTSUTBransUMTSUTId AS rid, UMTSUTBransUMTBransId AS bid, UMTSUTBransTipi AS tip FROM UMTSUTBrans')) {
    $n = ''; if ($r.bid -is [int] -and $umbMap.ContainsKey($r.bid)) { $n = $umbMap[$r.bid] }
    AddGroup $sutBrans ([int]$r.Uid) ("$n (#$($r.bid)) Tip: $($r.tip)")
}
$sutRaporIcd = @{}
foreach ($r in (Fill 'SELECT UMTSUTRICDUMTSUTRaporId AS rid, UMTSUTRICDler AS icdler FROM UMTSUTRaporICD')) {
    AddGroup $sutRaporIcd ([int]$r.rid) (Clean $r.icdler)
}
$sutRaporOdenmez = @{}
foreach ($r in (Fill 'SELECT UMTSUTRICDOUMTSUTRaporId AS rid, UMTSUTRICDOler AS ler FROM UMTSUTRaporICDOdenmez')) {
    AddGroup $sutRaporOdenmez ([int]$r.rid) (Clean $r.ler)
}
$sutRapor = @{}
foreach ($r in (Fill 'SELECT r.UMTSUTRaporUMTSUTId AS rid, r.UMTSUTRaporUMTRaporId AS rid2, r.UMTSUTRaporId AS ridr FROM UMTSUTRapor r')) {
    $kod = ''; $acd = ''
    if ($r.rid2 -is [int] -and $umrMap.ContainsKey($r.rid2)) { $k = $umrMap[$r.rid2]; $kod = $k[0]; $acd = $k[1] }
    $s = "UMTRapor: $kod - $acd (rule ref: $($r.ridr))"
    $rkey = [int]$r.ridr
    if ($sutRaporIcd.ContainsKey($rkey)) { $s += ' ; ICD: ' + (($sutRaporIcd[$rkey]) -join ' ; ') }
    if ($sutRaporOdenmez.ContainsKey($rkey)) { $s += ' ; Odenmeyen ICD: ' + (($sutRaporOdenmez[$rkey]) -join ' ; ') }
    AddGroup $sutRapor ([int]$r.rid) $s
}
$sutTesis = @{}
foreach ($r in (Fill 'SELECT t.UMTSUTTesisUMTSUTId AS rid, t.UMTSUTTesisUMTTesisId AS tid, t.UMTSUTTesisTip AS tip FROM UMTSUTTesis t')) {
    $n = ''; if ($r.tid -is [int] -and $utsMap.ContainsKey($r.tid)) { $n = $utsMap[$r.tid] }
    AddGroup $sutTesis ([int]$r.rid) ("$n (ref: $($r.tid)) Tip: $($r.tip)")
}
$sutTeshis = @{}
foreach ($r in (Fill 'SELECT UMTSUTRTUMTSUTId AS rid, UMTSUTRTUMTRTId AS tid FROM UMTSUTReceteTeshis')) {
    $n = ''; if ($r.tid -is [int] -and $umtMap.ContainsKey($r.tid)) { $k = $umtMap[$r.tid]; $n = "$($k[0]) $($k[1])" }
    AddGroup $sutTeshis ([int]$r.rid) $n
}
$sutOde = @{}
foreach ($r in (Fill 'SELECT UMTSUTOTUMTSUTId AS rid, UMTSUTOdenenTani AS od, UMTSUTOdenmeyenTani AS onm FROM UMTSUTOdeTani')) {
    AddGroup $sutOde ([int]$r.rid) ("Odenen: $(Clean $r.od) | Odenmeyen: $(Clean $r.onm)")
}
$sutKiyas = @{}
foreach ($r in (Fill 'SELECT UMTSUTKTSUTId AS rid, UMTSUTKTTanilar AS t, UMTSUTKTvarvar AS vv, UMTSUTKTvaryok AS vy, UMTSUTKTyokvar AS yv, UMTSUTKTyokyok AS yy, UMTSUTKTYasakKullanim AS yk, UMTSUTKTMecburKullanim AS mk, UMTSUTKTOzelNot AS not_ FROM UMTSUTKiyasTani')) {
    $s = "Tanilar: $(Clean $r.t) | Var+Var: $(Clean $r.vv) | Var+Yok: $(Clean $r.vy) | Yok+Var: $(Clean $r.yv) | Yok+Yok: $(Clean $r.yy) | Yasak: $(Clean $r.yk) | Mecbur: $(Clean $r.mk) | Ozel Not: $(Clean $r.not_)"
    AddGroup $sutKiyas ([int]$r.rid) $s
}
$sutEkA = @{}
foreach ($r in (Fill 'SELECT USEAUMTSUTId AS rid, USEAMaxIlacAdet AS adet, USEAGunlukKalori AS kal FROM UMTSUTEkA')) {
    AddGroup $sutEkA ([int]$r.rid) ("Max Ilac Adet: $($r.adet) | Gunluk Kalori: $(Num $r.kal)")
}

$enc = New-Object System.Text.UTF8Encoding($false)
$sw = New-Object System.IO.StreamWriter($outPath, $false, $enc, 8388608)
$cnt = @{ rapor = 0; sgk = 0; erx = 0; elden = 0; sut = 0; umt = 0 }

Write-Host 'HTML yaziliyor...'

$sw.WriteLine('<!DOCTYPE html><html lang="tr"><head><meta charset="utf-8"><title>Reçete, Rapor ve SUT Kurallari Korpusu</title>')
$sw.WriteLine('<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:1200px;margin:1rem auto;padding:0 1rem}section{border:1px solid #bbb;border-radius:6px;padding:.5rem .8rem;margin:.6rem 0;font-size:.85rem}.t{font-weight:bold}.b{color:#333;margin:.15rem 0}.f{color:#444;background:#f7f7f7;padding:.3rem;border-left:3px solid #999;white-space:pre-wrap}h2{background:#eee;padding:.3rem .6rem;border-radius:4px}h3{margin:.4rem 0}.meta{color:#666;font-size:.8rem}</style></head><body>')
$sw.WriteLine('<h1>Reçete, Rapor ve SUT Kurallari Korpusu</h1>')
$sw.WriteLine('<p class="meta">Üretim: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm') + '. V3: hasta yasi Musteri.MusteriDogumTarihi + recete/rapor tarihinden TURETILMIS tek sayi olarak yazildi (dogum tarihi ciktiya asla yazilmaz); ilac satirlarina urun katalog anahtari (BARKOD:/KAMU:) eklendi. Kimlik temizligi: gercek Rapor/Takip/Protokol No, SGK Islem No, E-Recete No, Renkli No ve ilac satiri rapor no alanlari salted-hash ile TUTARLI sahte kimliklere (PSD-...) cevrildi (ayni gercek numara -> ayni sahte deger; rapor-recete eslesmesi korunur); kurum/hastane adlari KURUM-n / HASTANE-n etiketlerine cevrildi. Kisisel kimlik verileri yoktur: hasta ad-soyad/TC/adres/telefon/dogum tarihi, doktor ad-soyad yazilmaz. Serbest metinlerdeki 11 haneli sayilar [MASKELI-11HANE], kimlik etiketli ifadeler [KIMLIK-SILENDI] olarak maskelenmistir. Periyot/Aralik/Tip alanlari kaynak sistem kodudur.</p>')

$ageStats = @{ raporDolu = 0; raporBos = 0; sgkDolu = 0; sgkBos = 0; erxDolu = 0; erxBos = 0 }
$dtR = Fill 'SELECT RaporAnaId, RaporAnaRaporTarihi, RaporAnaRaporNo, RaporAnaRaporTakipNo, RaporAnaProtokolNo, RaporAnaRaporTuruId, RaporAnaHastaneId, RaporAnaMusteriId, RaporAnaAciklamalar FROM RaporAna WHERE ISNULL(RaporAnaSilme,0)=0 ORDER BY RaporAnaId'
$sw.WriteLine('<h2>1) Raporlar (' + $dtR.Rows.Count + ' kayit)</h2>')
foreach ($r in $dtR.Rows) {
    $id = [int]$r.RaporAnaId
    $tur = ''; if ($r.RaporAnaRaporTuruId -is [int] -and $rtMap.ContainsKey($r.RaporAnaRaporTuruId)) { $tur = $rtMap[$r.RaporAnaRaporTuruId] }
    $has = ''; if ($r.RaporAnaHastaneId -is [int] -and $hasMap.ContainsKey($r.RaporAnaHastaneId)) { $has = $hasMap[$r.RaporAnaHastaneId] }
    $yas = ''
    $hRef = 0
    if ([int]::TryParse([string]$r.RaporAnaMusteriId, [ref]$hRef) -and $hastaDogum.ContainsKey($hRef)) { $yas = YasYaz $hastaDogum[$hRef] $r.RaporAnaRaporTarihi }
    if ($yas) { $ageStats.raporDolu++ } else { $ageStats.raporBos++ }
    $sw.WriteLine('<section><div class="t">RAPOR R#' + $id + ' [' + (Esc $tur) + '] — Rapor Tarihi: ' + (Dte $r.RaporAnaRaporTarihi) + ' | Rapor No: ' + (Pseudo $r.RaporAnaRaporNo) + ' | Takip No: ' + (Pseudo $r.RaporAnaRaporTakipNo) + ' | Protokol No: ' + (Pseudo $r.RaporAnaProtokolNo) + ' | Kurum: ' + (Esc $has) + ' | Yas: ' + $yas + '</div>')
    if ($repIcd.ContainsKey($id))  { $sw.WriteLine('<div class="b"><b>Tani/Rapor Kodlari:</b> ' + (($repIcd[$id] | ForEach-Object { Esc $_ }) -join ' ; ') + '</div>') }
    if ($repEm.ContainsKey($id))   { $sw.WriteLine('<div class="b"><b>Etken Maddeler:</b> ' + (($repEm[$id] | ForEach-Object { Esc $_ }) -join ' ; ') + '</div>') }
    if ($repBrans.ContainsKey($id)){ $sw.WriteLine('<div class="b"><b>Kurul Hekim Branslari:</b> ' + (($repBrans[$id] | ForEach-Object { Esc $_ }) -join ', ') + '</div>') }
    if ($repEk.ContainsKey($id))   { $sw.WriteLine('<div class="b"><b>Ek Bilgiler (Lab vb.):</b> ' + (($repEk[$id] | ForEach-Object { Clean $_ }) -join ' ; ') + '</div>') }
    if ($r.RaporAnaAciklamalar)    { $sw.WriteLine('<div class="f"><b>Kanaat/Notlar:</b>' + [Environment]::NewLine + (Clean ([string]$r.RaporAnaAciklamalar)) + '</div>') }
    $sw.WriteLine('</section>')
    $cnt.rapor++
}

$dtS = Fill 'SELECT RxId, RxReceteTarihi, RxKayitTarihi, RxIslemTarihi, RxSeriNo, RxProtokolNo, RxSgkIslemNo, RxEReceteNo, RxKurumId, RxHastaneId, RxReceteRenkId, RxReceteAltTuruId, RxProvizyonTipId, RxBransId, RxReceteTipi, RxDoktorId, RxCevapKodu, RxITSBildirim, RxUtsDurum, RxMusteriId FROM ReceteAna WHERE ISNULL(RxSilme,0)=0 ORDER BY RxId'
$sw.WriteLine('<h2>2) SGK Receteleri (' + $dtS.Rows.Count + ' kayit)</h2>')
foreach ($r in $dtS.Rows) {
    $id = [int]$r.RxId
    $kur = ''; if ($r.RxKurumId -is [int] -and $kurumMap.ContainsKey($r.RxKurumId)) { $kur = $kurumMap[$r.RxKurumId] }
    $has = ''; if ($r.RxHastaneId -is [int] -and $hasMap.ContainsKey($r.RxHastaneId)) { $has = $hasMap[$r.RxHastaneId] }
    $rn = ''; if ($r.RxReceteRenkId -is [int] -and $renkMap.ContainsKey($r.RxReceteRenkId)) { $rn = $renkMap[$r.RxReceteRenkId] }
    $at = ''; if ($r.RxReceteAltTuruId -is [int] -and $altMap.ContainsKey($r.RxReceteAltTuruId)) { $at = $altMap[$r.RxReceteAltTuruId] }
    $pv = ''; if ($r.RxProvizyonTipId -is [int] -and $prvMap.ContainsKey($r.RxProvizyonTipId)) { $pv = $prvMap[$r.RxProvizyonTipId] }
    $br = ''; if ($r.RxBransId -is [int] -and $bransMap.ContainsKey($r.RxBransId)) { $br = $bransMap[$r.RxBransId] }
    $yas = ''
    $hRef = 0
    if ([int]::TryParse([string]$r.RxMusteriId, [ref]$hRef) -and $hastaDogum.ContainsKey($hRef)) { $yas = YasYaz $hastaDogum[$hRef] $r.RxReceteTarihi }
    if ($yas) { $ageStats.sgkDolu++ } else { $ageStats.sgkBos++ }
    $sw.WriteLine('<section><div class="t">SGK RX #' + $id + ' — Recete Tarihi: ' + (Dte $r.RxReceteTarihi) + ' | Kayit: ' + (Dte $r.RxKayitTarihi) + ' | Islem: ' + (Dts $r.RxIslemTarihi) + ' | Seri No: ' + (Pseudo $r.RxSeriNo) + ' | Protokol: ' + (Pseudo $r.RxProtokolNo) + ' | SGK Islem No: ' + (Pseudo $r.RxSgkIslemNo) + ' | E-Recete No: ' + (Pseudo $r.RxEReceteNo) + ' | Yas: ' + $yas + '</div>')
    $sw.WriteLine('<div class="b">Kurum: ' + (Esc $kur) + ' | Hastane: ' + (Esc $has) + ' | Renk: ' + (Esc $rn) + ' | Alt Tur: ' + (Esc $at) + ' | Provizyon Tipi: ' + (Esc $pv) + ' | Brans: ' + (Esc $br) + ' | Recete Tipi: ' + (Esc $r.RxReceteTipi) + ' | UTS Durum: ' + (Esc $r.RxUtsDurum) + ' | Cevap Kodu: ' + (Flag $r.RxCevapKodu) + ' | ITS Bildirim: ' + (Flag $r.RxITSBildirim) + ' | Doktor Ref: ' + (Esc $r.RxDoktorId) + '</div>')
    if ($rxIcd.ContainsKey($id))    { $sw.WriteLine('<div class="b"><b>Tanilar:</b> ' + (($rxIcd[$id] | ForEach-Object { Esc $_ }) -join ' ; ') + '</div>') }
    if ($rxIlac.ContainsKey($id))   { $sw.WriteLine('<div class="b"><b>İlaçlar:</b><ul>' + (($rxIlac[$id] | ForEach-Object { '<li>' + (Esc $_) + '</li>' }) -join '') + '</ul></div>') }
    if ($rxTeshis.ContainsKey($id)) { $sw.WriteLine('<div class="b"><b>Recete Teshis Secimi:</b> ' + (($rxTeshis[$id] | ForEach-Object { Esc $_ }) -join ' ; ') + '</div>') }
    if ($rxOzel.ContainsKey($id))   { $sw.WriteLine('<div class="b"><b>Ozellikler:</b> ' + (($rxOzel[$id] | ForEach-Object { Esc $_ }) -join ' ; ') + '</div>') }
    $sw.WriteLine('</section>')
    $cnt.sgk++
}

$dtE = Fill 'SELECT EReceteId, EReceteNo, EReceteTakipNo, EReceteTarihi, EReceteProtokolNo, EReceteReceteAltTuruId, EReceteReceteRenkId, EReceteProvizyonTipId, EReceteBransId, EReceteHastaneId, EReceteDoktorId, EReceteImzaTuru, EReceteRenkliNo, EReceteSertifikaId, EReceteMusteriId FROM ERecete WHERE ISNULL(EReceteSilme,0)=0 ORDER BY EReceteId'
$sw.WriteLine('<h2>3) E-Receteler (' + $dtE.Rows.Count + ' kayit)</h2>')
foreach ($r in $dtE.Rows) {
    $id = [int]$r.EReceteId
    $at = ''; if ($r.EReceteReceteAltTuruId -is [int] -and $altMap.ContainsKey($r.EReceteReceteAltTuruId)) { $at = $altMap[$r.EReceteReceteAltTuruId] }
    $rn = ''; if ($r.EReceteReceteRenkId -is [int] -and $renkMap.ContainsKey($r.EReceteReceteRenkId)) { $rn = $renkMap[$r.EReceteReceteRenkId] }
    $pv = ''; if ($r.EReceteProvizyonTipId -is [int] -and $prvMap.ContainsKey($r.EReceteProvizyonTipId)) { $pv = $prvMap[$r.EReceteProvizyonTipId] }
    $br = ''; if ($r.EReceteBransId -is [int] -and $bransMap.ContainsKey($r.EReceteBransId)) { $br = $bransMap[$r.EReceteBransId] }
    $has = ''; if ($r.EReceteHastaneId -is [int] -and $hasMap.ContainsKey($r.EReceteHastaneId)) { $has = $hasMap[$r.EReceteHastaneId] }
    $ser = ''; if ($r.EReceteSertifikaId -is [int] -and $serMap.ContainsKey($r.EReceteSertifikaId)) { $ser = $serMap[$r.EReceteSertifikaId] }
    $yas = ''
    $hRef = 0
    if ([int]::TryParse([string]$r.EReceteMusteriId, [ref]$hRef) -and $hastaDogum.ContainsKey($hRef)) { $yas = YasYaz $hastaDogum[$hRef] $r.EReceteTarihi }
    if ($yas) { $ageStats.erxDolu++ } else { $ageStats.erxBos++ }
    $sw.WriteLine('<section><div class="t">E-RX #' + $id + ' — Recete Tarihi: ' + (Dte $r.EReceteTarihi) + ' | E-Recete No: ' + (Pseudo $r.EReceteNo) + ' | Takip No: ' + (Pseudo $r.EReceteTakipNo) + ' | Protokol: ' + (Pseudo $r.EReceteProtokolNo) + ' | Renkli No: ' + (Pseudo $r.EReceteRenkliNo) + ' | Yas: ' + $yas + '</div>')
    $sw.WriteLine('<div class="b">Alt Tur: ' + (Esc $at) + ' | Renk: ' + (Esc $rn) + ' | Provizyon: ' + (Esc $pv) + ' | Brans: ' + (Esc $br) + ' | Hastane: ' + (Esc $has) + ' | Imza Turu: ' + (Esc $r.EReceteImzaTuru) + ' | Sertifika: ' + (Esc $ser) + ' | Doktor Ref: ' + (Esc $r.EReceteDoktorId) + '</div>')
    if ($erIcd.ContainsKey($id))  { $sw.WriteLine('<div class="b"><b>Tanilar:</b> ' + (($erIcd[$id] | ForEach-Object { Esc $_ }) -join ' ; ') + '</div>') }
    if ($erIlac.ContainsKey($id)) { $sw.WriteLine('<div class="b"><b>İlaçlar:</b><ul>' + (($erIlac[$id] | ForEach-Object { '<li>' + (Esc $_) + '</li>' }) -join '') + '</ul></div>') }
    if ($erNot.ContainsKey($id))  { $sw.WriteLine('<div class="b"><b>Aciklamalar:</b> ' + (($erNot[$id] | ForEach-Object { Esc $_ }) -join ' ; ') + '</div>') }
    $sw.WriteLine('</section>')
    $cnt.erx++
}

$dtL = Fill 'SELECT RxId, RxReceteTarihi, RxProtokolNo, RxBelgeNo, RxHastaneId FROM EldenAna WHERE ISNULL(RxSilme,0)=0 ORDER BY RxId'
$sw.WriteLine('<h2>4) Elden Receteler (' + $dtL.Rows.Count + ' kayit)</h2>')
foreach ($r in $dtL.Rows) {
    $id = [int]$r.RxId
    $has = ''; if ($r.RxHastaneId -is [int] -and $hasMap.ContainsKey($r.RxHastaneId)) { $has = $hasMap[$r.RxHastaneId] }
    $sw.WriteLine('<section><div class="t">ELDEN #' + $id + ' — Recete Tarihi: ' + (Dte $r.RxReceteTarihi) + ' | Protokol: ' + (Pseudo $r.RxProtokolNo) + ' | Belge No: ' + (Pseudo $r.RxBelgeNo) + ' | Hastane: ' + (Esc $has) + '</div>')
    if ($elIlac.ContainsKey($id)) { $sw.WriteLine('<div class="b"><b>İlaçlar:</b><ul>' + (($elIlac[$id] | ForEach-Object { '<li>' + (Esc $_) + '</li>' }) -join '') + '</ul></div>') }
    $sw.WriteLine('</section>')
    $cnt.elden++
}

$sw.WriteLine('<h2>5) SUT Kurallari (referans tablolar)</h2>')
$sw.WriteLine('<p class="meta">SUT = Saglik Uygulama Tebligi: SGK kapsaminda tani/tedavi amaciyla ilac raporlarinin ve recetelerin odeme kurallarini duzenler (rapor kodlari, ICD eslesmeleri, kullanim suresi/doz/tekrar sinirlari, yas-cinsiyet sartlari, brans-tesis kurallari, takip yontemleri). Asagida Botanik sisteminin kullandigi SUT kural setinin tam kopyasi bulunur.</p>')

$dtSut = Fill 'SELECT SUTId, SUTUstId, SUTKodu, SUTBaslik, SUTYururlukTarihi, SUTGecerlilikTarihi, SUTIcerik, SUTUyarma FROM SUT WHERE ISNULL(SUTSilme,0)=0 ORDER BY SUTId'
$sw.WriteLine('<h3>5.1) SUT Tablosu (' + $dtSut.Rows.Count + ')</h3>')
foreach ($r in $dtSut.Rows) {
    $sw.WriteLine('<section><div class="t">SUT [' + (Esc $r.SUTKodu) + '] ' + (Clean $r.SUTBaslik) + '</div>')
    $sw.WriteLine('<div class="b">Yururluk: ' + (Dte $r.SUTYururlukTarihi) + ' | Gecerlilik: ' + (Dte $r.SUTGecerlilikTarihi) + ' | Ust Id: ' + (Esc $r.SUTUstId) + ' | Uyarma: ' + (Flag $r.SUTUyarma) + '</div>')
    if ($r.SUTIcerik) { $sw.WriteLine('<div class="f">' + (Clean ([string]$r.SUTIcerik)) + '</div>') }
    $sw.WriteLine('</section>')
}

$dtUmt = Fill 'SELECT UMTId, UMTUrunId, UMTUrunAdi, UMTAmbalajMiktar, UMTTekDoz, UMTKullanimSuresi, UMTKullanimSuresiPeriyot, UMTCinsiyet, UMTYasMin1, UMTYasMax1, UMTYasMin2, UMTYasMax2, UMTReceteTur1, UMTReceteTur2, UMTKamuKodu, UMTTedaviSemasi, UMTAyAralik1, UMTAyAralik2, UMTAyakAralik, UMTAyakPeriyot, UMTAyakTekrar, UMTAyakDoz, UMTYatanAralik, UMTYatanPeriyot, UMTYatanTekrar, UMTYatanDoz, UMTRaporluAralik, UMTRaporluPeriyot, UMTRaporluTekrar, UMTRaporluDoz, UMTTISAralik, UMTTISPeriyot, UMTTISAdet, UMTKutuKalori FROM UMT ORDER BY UMTId'
$sw.WriteLine('<h3>5.2) UMT (Medula ilac kurallari) (' + $dtUmt.Rows.Count + ')</h3>')
foreach ($r in $dtUmt.Rows) {
    $sw.WriteLine('<section><div class="t">UMT U#' + $r.UMTId + ' ' + (Clean $r.UMTUrunAdi) + '</div>')
    $sw.WriteLine('<div class="b">Kamu Kodu: ' + (Clean $r.UMTKamuKodu) + ' | Ambalaj: ' + (Num $r.UMTAmbalajMiktar) + ' | Tek Doz: ' + (Num $r.UMTTekDoz) + ' | Kullanim Suresi: ' + (Num $r.UMTKullanimSuresi) + ' (' + (Esc $r.UMTKullanimSuresiPeriyot) + ') | Cinsiyet: ' + (Esc $r.UMTCinsiyet) + ' | Yas: ' + (Esc $r.UMTYasMin1) + '-' + (Esc $r.UMTYasMax1) + ' / ' + (Esc $r.UMTYasMin2) + '-' + (Esc $r.UMTYasMax2) + ' | Recete Tur: ' + (Clean $r.UMTReceteTur1) + ' / ' + (Clean $r.UMTReceteTur2) + ' | Tedavi Semasi: ' + (Esc $r.UMTTedaviSemasi) + '</div>')
    $sw.WriteLine('<div class="b">Ayaktan: Doz ' + (Num $r.UMTAyakDoz) + '/Aralik ' + (Esc $r.UMTAyakAralik) + '/Periyot ' + (Esc $r.UMTAyakPeriyot) + '/Tekrar ' + (Esc $r.UMTAyakTekrar) + ' | Yatan: Doz ' + (Num $r.UMTYatanDoz) + '/Aralik ' + (Esc $r.UMTYatanAralik) + '/Periyot ' + (Esc $r.UMTYatanPeriyot) + '/Tekrar ' + (Esc $r.UMTYatanTekrar) + ' | Raporlu: Doz ' + (Num $r.UMTRaporluDoz) + '/Aralik ' + (Esc $r.UMTRaporluAralik) + '/Periyot ' + (Esc $r.UMTRaporluPeriyot) + '/Tekrar ' + (Esc $r.UMTRaporluTekrar) + '</div>')
    $sw.WriteLine('<div class="b">TIS: Aralik ' + (Esc $r.UMTTISAralik) + '/Periyot ' + (Esc $r.UMTTISPeriyot) + '/Adet ' + (Esc $r.UMTTISAdet) + ' | Ay Aralik: ' + (Esc $r.UMTAyAralik1) + '/' + (Esc $r.UMTAyAralik2) + ' | Kutu Kalori: ' + (Num $r.UMTKutuKalori) + '</div>')
    $sw.WriteLine('</section>')
    $cnt.umt++
}

$dtUsut = Fill 'SELECT UMTSUTId, UMTSUTUrunId, UMTSUTKod, UMTSUTTedaviTip, UMTSUTRpTip, UMTSUTCinsiyet, UMTSUTYasMin, UMTSUTYasMax, UMTSUTTakipYontemiId, UMTSUTTakipAralik, UMTSUTTakipPeriyot, UMTSUTTakipAdet, UMTSUTTakipTur, UMTSUTEtkMad1, UMTSUTEtkMad2, UMTSUTEtkMad3, UMTSUTEtkMad4, UMTSUTEtkMad5, UMTSUTRxDrTipi, UMTSUTRxSrtTipi, UMTSUTRpTipi, UMTSUTRpSrtTipi, UMTSUTRpDrTipi, UMTSUTRpVeveya, UMTSUTRpMaxTarih, UMTSUTRpMaxSure, UMTSUTRpMaxSureTip, UMTSUTMaxOdSure, UMTSUTMaxOdSureTip, UMTSUTMaxKullAralik, UMTSUTMaxKullPeriyot, UMTSUTMaxKullTekrar, UMTSUTMaxKullDoz FROM UMTSUT ORDER BY UMTSUTId'
$sw.WriteLine('<h3>5.3) UMTSUT (urun bazli SUT kurallari) (' + $dtUsut.Rows.Count + ')</h3>')
foreach ($r in $dtUsut.Rows) {
    $id = [int]$r.UMTSUTId
    $uad = ''; if ($r.UMTSUTUrunId -is [int] -and $urunMap.ContainsKey($r.UMTSUTUrunId)) { $uad = $urunMap[$r.UMTSUTUrunId] }
    $etks = @(); foreach ($x in @($r.UMTSUTEtkMad1, $r.UMTSUTEtkMad2, $r.UMTSUTEtkMad3, $r.UMTSUTEtkMad4, $r.UMTSUTEtkMad5)) { if ($x -is [int] -and $emMap.ContainsKey($x)) { $etks += $emMap[$x] } }
    $sw.WriteLine('<section><div class="t">SUT KURAL K#' + $id + ' — Urun: ' + (Esc $uad) + ' (ref: ' + (Esc $r.UMTSUTUrunId) + ') | SUT Kod: ' + (Clean $r.UMTSUTKod) + '</div>')
    $sw.WriteLine('<div class="b">Tedavi Tip: ' + (Esc $r.UMTSUTTedaviTip) + ' | Rp Tip: ' + (Esc $r.UMTSUTRpTip) + '/' + (Esc $r.UMTSUTRpTipi) + ' | Cinsiyet: ' + (Esc $r.UMTSUTCinsiyet) + ' | Yas: ' + (Esc $r.UMTSUTYasMin) + '-' + (Esc $r.UMTSUTYasMax) + ' | Takip: Yontem ' + (Esc $r.UMTSUTTakipYontemiId) + ', Aralik ' + (Esc $r.UMTSUTTakipAralik) + ', Periyot ' + (Esc $r.UMTSUTTakipPeriyot) + ', Adet ' + (Esc $r.UMTSUTTakipAdet) + ', Tur ' + (Esc $r.UMTSUTTakipTur) + '</div>')
    $sw.WriteLine('<div class="b">Rx Dr Tip: ' + (Esc $r.UMTSUTRxDrTipi) + ' | Rx Srt Tip: ' + (Esc $r.UMTSUTRxSrtTipi) + ' | Rp Srt: ' + (Esc $r.UMTSUTRpSrtTipi) + ' | Rp Dr: ' + (Esc $r.UMTSUTRpDrTipi) + ' | Rp Ve/Ya: ' + (Esc $r.UMTSUTRpVeveya) + ' | Rp Max Tarih: ' + (Dte $r.UMTSUTRpMaxTarih) + ' | Rp Max Sure: ' + (Esc $r.UMTSUTRpMaxSure) + ' (' + (Esc $r.UMTSUTRpMaxSureTip) + ') | Max Od Sure: ' + (Esc $r.UMTSUTMaxOdSure) + ' (' + (Esc $r.UMTSUTMaxOdSureTip) + ')</div>')
    $sw.WriteLine('<div class="b">Max Kullanim: Doz ' + (Num $r.UMTSUTMaxKullDoz) + '/Aralik ' + (Esc $r.UMTSUTMaxKullAralik) + '/Periyot ' + (Esc $r.UMTSUTMaxKullPeriyot) + '/Tekrar ' + (Esc $r.UMTSUTMaxKullTekrar) + '</div>')
    if ($etks.Count -gt 0) { $sw.WriteLine('<div class="b"><b>Etken Maddeler:</b> ' + (($etks | ForEach-Object { Esc $_ }) -join ', ') + '</div>') }
    if ($sutBrans.ContainsKey($id))  { $sw.WriteLine('<div class="b"><b>Brans Kurallari:</b> ' + (($sutBrans[$id] | ForEach-Object { Esc $_ }) -join ' ; ') + '</div>') }
    if ($sutRapor.ContainsKey($id))  { $sw.WriteLine('<div class="b"><b>Rapor Kurallari:</b> ' + (($sutRapor[$id] | ForEach-Object { Esc $_ }) -join ' ; ') + '</div>') }
    $sw.WriteLine('</section>')
    $cnt.sut++
}

$sw.WriteLine('<p class="meta">Toplam: ' + $cnt.rapor + ' rapor, ' + $cnt.sgk + ' SGK recete, ' + $cnt.erx + ' e-recete, ' + $cnt.elden + ' elden recete, ' + $cnt.umt + ' UMT kurali, ' + $cnt.sut + ' UMTSUT kurali.</p>')
$sw.WriteLine('</body></html>')
$sw.Flush()
$sw.Dispose()
$conn.Close()

Write-Host ("TAMAM. Rapor: $($cnt.rapor) | SGK: $($cnt.sgk) | ERx: $($cnt.erx) | Elden: $($cnt.elden) | UMT: $($cnt.umt) | UMTSUT: $($cnt.sut)")
Write-Host ("YAS dolu/bos -> rapor: $($ageStats.raporDolu)/$($ageStats.raporBos) | sgk: $($ageStats.sgkDolu)/$($ageStats.sgkBos) | erx: $($ageStats.erxDolu)/$($ageStats.erxBos)")
Write-Host ("URUNKEY dolu/bos: $($ukeyDolu)/$($ukeyBos)")
Write-Host ("Dosya: $outPath | Boyut: " + (Get-Item $outPath).Length)

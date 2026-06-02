# -*- coding: utf-8 -*-
"""
Turizm, mehmonxona va restoran biznesida axborot texnologiyalari - Konspekt generatori.
Standart Python kutubxonalari bilan .docx fayl yaratadi.
"""
import zipfile
import datetime

TOPICS = []

def esc(text):
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))

def run(text, bold=False, italic=False, size=28, font="Times New Roman"):
    rpr = '<w:rPr>'
    rpr += '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (font, font, font)
    if bold:
        rpr += '<w:b/><w:bCs/>'
    if italic:
        rpr += '<w:i/><w:iCs/>'
    rpr += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size, size)
    rpr += '</w:rPr>'
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, esc(text))

def para(text, bold=False, italic=False, size=28, align="both",
         indent_first=709, space_after=120, line=360, keep_next=False):
    ppr = '<w:pPr>'
    if keep_next:
        ppr += '<w:keepNext/>'
    ppr += '<w:spacing w:after="%d" w:line="%d" w:lineRule="auto"/>' % (space_after, line)
    if indent_first:
        ppr += '<w:ind w:firstLine="%d"/>' % indent_first
    ppr += '<w:jc w:val="%s"/>' % align
    ppr += '</w:pPr>'
    return '<w:p>%s%s</w:p>' % (ppr, run(text, bold=bold, italic=italic, size=size))

def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def empty_para():
    return '<w:p><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/></w:pPr></w:p>'

def build_topic(t, is_first):
    parts = []
    if not is_first:
        parts.append(page_break())
    parts.append(para("%d-MAVZU. %s" % (t["num"], t["title"].upper()),
                      bold=True, size=30, align="center",
                      indent_first=0, space_after=180, keep_next=True))
    parts.append(para("Reja:", bold=True, size=28, align="left",
                      indent_first=0, space_after=60, keep_next=True))
    for i, r in enumerate(t["reja"], 1):
        parts.append(para("%d. %s" % (i, r), size=28, align="left",
                          indent_first=360, space_after=40))
    parts.append(empty_para())
    for idx, (head, paras) in enumerate(t["sections"], 1):
        parts.append(para("%d. %s" % (idx, head), bold=True, size=28,
                          align="left", indent_first=0, space_after=80, keep_next=True))
        for p in paras:
            parts.append(para(p, size=28, align="both", indent_first=709, space_after=120))
    return "".join(parts)

def build_document():
    body = []
    for i, t in enumerate(TOPICS):
        body.append(build_topic(t, is_first=(i == 0)))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" '
            'w:header="708" w:footer="708" w:gutter="0"/>'
            '<w:pgNumType w:start="1"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:body>' + "".join(body) + sect + '</w:body></w:document>')

CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '</Relationships>')

DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>')

STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
    '<w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="uz-Latn-UZ"/>'
    '</w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    '</w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/><w:rPr>'
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
    '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>'
    '</w:styles>')

def core_props():
    now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Turizm, mehmonxona va restoran biznesida AT - Konspekt</dc:title>'
        '<dc:creator>Konspekt generatori</dc:creator>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
        '</cp:coreProperties>' % (now, now))

APP_PROPS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    '<Application>Konspekt generatori</Application></Properties>')

def write_docx(path="Turizm_Konspekt_30_mavzu.docx"):
    document = build_document()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("docProps/core.xml", core_props())
        z.writestr("docProps/app.xml", APP_PROPS)
    print("Yaratildi:", path)
    print("Mavzular soni:", len(TOPICS))

# ===================== MAVZULAR =====================

TOPICS.append({
    "num": 1,
    "title": "Fanga kirish. Turizm, mehmonxona va restoran biznesida axborot texnologiyalari",
    "reja": [
        "Fanning maqsadi, vazifalari va o'rganish obyekti.",
        "Turizm sohasida axborot texnologiyalarining ahamiyati.",
        "Mehmonxona va restoran biznesida AT ning tutgan o'rni.",
        "Fanning boshqa fanlar bilan aloqasi va istiqbollari.",
    ],
    "sections": [
        ("Fanning maqsadi, vazifalari va o'rganish obyekti", [
            "Ushbu fanning asosiy maqsadi — talabalarda turizm, mehmonxona va restoran biznesi sohasida axborot texnologiyalaridan samarali foydalanish ko'nikma va malakalarini shakllantirishdir. Fan zamonaviy AT vositalarini o'rganish orqali bo'lajak mutaxassislarni raqobatbardosh kadrlar sifatida tayyorlashga qaratilgan.",
            "Fanning vazifalari qatoriga quyidagilar kiradi: turizmda qo'llaniladigan dasturiy ta'minotni o'zlashtirish, bronlash tizimlarini o'rganish, elektron tijorat va reklama vositalarini bilish, ofis dasturlari yordamida hujjatlarni rasmiylashtirish va multimediya vositalaridan foydalanishni egallash.",
            "Fanning o'rganish obyekti — turizm, mehmonxona va restoran korxonalari faoliyatida qo'llaniladigan zamonaviy axborot texnologiyalari, dasturiy vositalar, kommunikatsiya tizimlari va ularni amaliyotda tatbiq etish usullaridir.",
        ]),
        ("Turizm sohasida axborot texnologiyalarining ahamiyati", [
            "Zamonaviy turizm sanoati axborot texnologiyalarisiz tasavvur qilib bo'lmaydi. AT turistik mahsulotlarni yaratish, sotish, reklama qilish va mijozlarga xizmat ko'rsatishning barcha bosqichlarida keng qo'llaniladi. Onlayn bronlash tizimlari, elektron sayohat agentliklari va mobil ilovalar turizm xizmatlarini yangi darajaga ko'tardi.",
            "Axborot texnologiyalari turistik korxonalarga global bozorga chiqish, xalqaro mijozlar bilan aloqa o'rnatish va xizmatlarni dunyo bo'ylab taklif qilish imkonini beradi. Bu esa korxonaning daromadini oshiradi va raqobatbardoshligini ta'minlaydi.",
            "AT yordamida turistik ma'lumotlarni to'plash, tahlil qilish va qaror qabul qilish jarayonlari avtomatlashtiriladi. Bu mijozlar ehtiyojlarini yaxshiroq tushunish, xizmat sifatini oshirish va biznes jarayonlarini optimallashtirish imkonini beradi.",
        ]),
        ("Mehmonxona va restoran biznesida AT ning tutgan o'rni", [
            "Mehmonxona biznesida AT xonalarni bron qilish, mijozlarni ro'yxatga olish, hisob-kitob qilish va xizmat ko'rsatish jarayonlarini boshqarish uchun ishlatiladi. Property Management System (PMS) tizimlari mehmonxona faoliyatini yaxlit boshqaradi.",
            "Restoran biznesida AT buyurtmalarni qabul qilish, oshxona bilan aloqa, hisobni tayyorlash, inventarni boshqarish va mijozlar bazasini yuritish uchun qo'llaniladi. POS (Point of Sale) tizimlari restoran ishini avtomatlashtiradi va tezlashtiradi.",
            "Ikkala sohada ham AT mijozlar tajribasini yaxshilash, xodimlar samaradorligini oshirish, xarajatlarni kamaytirish va boshqaruv qarorlarini ma'lumotlarga asoslab qabul qilish imkonini beradi.",
        ]),
        ("Fanning boshqa fanlar bilan aloqasi va istiqbollari", [
            "Fan informatika, menejment, marketing, iqtisodiyot va turizm asoslari fanlari bilan chambarchas bog'liq. U bir tomondan texnik bilimlarni, ikkinchi tomondan turizm sohasining o'ziga xos xususiyatlarini birlashtiradi.",
            "Kelajakda sun'iy intellekt, katta ma'lumotlar tahlili (Big Data), virtual reallik (VR), narsalar interneti (IoT) va bulutli texnologiyalar turizm sohasida yanada keng qo'llanilishi kutilmoqda. Bu texnologiyalar personallashtirilgan xizmatlar va aqlli mehmonxonalar yaratishga asos bo'ladi.",
            "Fanni o'zlashtirgan mutaxassis zamonaviy texnologiyalarni biladi, ularni turizm biznesiga tatbiq eta oladi va raqamli transformatsiya sharoitida samarali ishlash qobiliyatiga ega bo'ladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 2,
    "title": "Zamonaviy axborot texnologiyalar. Ularga oid ma'lumotlarni to'plash va tahlili",
    "reja": [
        "Zamonaviy axborot texnologiyalari tushunchasi va turlari.",
        "Ma'lumotlarni to'plash usullari va vositalari.",
        "Ma'lumotlarni tahlil qilish texnologiyalari.",
        "Turizm sohasida ma'lumotlar tahlilining amaliy qo'llanilishi.",
    ],
    "sections": [
        ("Zamonaviy axborot texnologiyalari tushunchasi va turlari", [
            "Zamonaviy axborot texnologiyalari — bu axborotni yaratish, qayta ishlash, saqlash, uzatish va himoyalashni ta'minlovchi usul va vositalarning yaxlit tizimi. Ular bulutli hisoblash, mobil texnologiyalar, sun'iy intellekt, katta ma'lumotlar va narsalar internetini o'z ichiga oladi.",
            "AT lar infratuzilma texnologiyalari (serverlar, tarmoqlar), platforma texnologiyalari (operatsion tizimlar, ma'lumotlar bazalari) va amaliy texnologiyalarga (dasturiy ilovalar, veb-xizmatlar) bo'linadi. Har bir qatlam o'z vazifasini bajaradi.",
            "Turizm sohasida zamonaviy AT lar onlayn bronlash platformalari (Booking.com, Airbnb), mobil ilovalar, GPS navigatsiya, QR-kodlar va raqamli to'lov tizimlari ko'rinishida namoyon bo'ladi. Ular turistlarning sayohat tajribasini tubdan o'zgartirdi.",
        ]),
        ("Ma'lumotlarni to'plash usullari va vositalari", [
            "Turizm sohasida ma'lumotlar turli manbalardan to'planadi: veb-saytlar, ijtimoiy tarmoqlar, bronlash tizimlari, anketa so'rovnomalar, mijozlar fikr-mulohazalari va sensorlar. Har bir manba o'ziga xos ma'lumot turini taqdim etadi.",
            "Ma'lumotlarni to'plash vositalari: veb-skraping dasturlari, CRM (Customer Relationship Management) tizimlari, Google Analytics kabi tahlil platformalari, onlayn so'rovnomalar (Google Forms, SurveyMonkey) va POS tizimlarining hisobotlari.",
            "To'plangan ma'lumotlar tarkibli (jadvallar, bazalar) va tarkibsiz (matnlar, rasmlar, video) bo'lishi mumkin. Zamonaviy AT barcha turdagi ma'lumotlarni qayta ishlash va foydali axborot ajratib olish imkonini beradi.",
        ]),
        ("Ma'lumotlarni tahlil qilish texnologiyalari", [
            "Ma'lumotlarni tahlil qilish — to'plangan xom ma'lumotlardan foydali xulosalar chiqarish jarayoni. Buning uchun statistik tahlil, vizualizatsiya, klasterlash va bashorat qilish usullaridan foydalaniladi.",
            "Tahlil vositalari: Microsoft Excel (oddiy tahlil va diagrammalar), Google Data Studio (vizualizatsiya), Python va R (ilg'or tahlil), Power BI va Tableau (biznes-intellekt). Bu vositalar ma'lumotlarni tushunarli shaklga keltiradi.",
            "Big Data texnologiyalari katta hajmdagi ma'lumotlarni tez qayta ishlash imkonini beradi. Turizm sohasida bu texnologiyalar yo'nalishlar trendini aniqlash, narxlarni optimallashtirish va mijozlar xulq-atvorini bashorat qilish uchun ishlatiladi.",
        ]),
        ("Turizm sohasida ma'lumotlar tahlilining amaliy qo'llanilishi", [
            "Revenue Management — mehmonxonalarda talab va taklif tahlili asosida xona narxlarini dinamik boshqarish. Ma'lumotlar tahlili orqali yuqori va past mavsumlar aniqlanadi, narxlar shunga mos optimallanadi.",
            "Mijozlar segmentatsiyasi — ma'lumotlar tahlili orqali mijozlarni guruhlarga bo'lish (yoshiga, daromadiga, qiziqishlariga ko'ra) va har bir guruhga moslashtirilgan takliflar tayyorlash. Bu marketing samaradorligini oshiradi.",
            "Raqobatchilar tahlili — internet orqali raqobatchilarning narxlari, xizmatlari va mijozlar fikrlarini kuzatish. Bu ma'lumotlar o'z korxonasining strategiyasini ishlab chiqishda muhim asos bo'ladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 3,
    "title": "Axborotlarni saqlash va ko'chirishda zamonaviy vositalardan foydalanish",
    "reja": [
        "Axborotni saqlash vositalari va ularning tasnifi.",
        "Bulutli saqlash xizmatlari va ularning afzalliklari.",
        "Ma'lumotlarni ko'chirish va uzatish texnologiyalari.",
        "Axborot xavfsizligi va zaxira nusxalash.",
    ],
    "sections": [
        ("Axborotni saqlash vositalari va ularning tasnifi", [
            "Axborotni saqlash vositalari mahalliy (lokal) va masofaviy (remote) turlarga bo'linadi. Mahalliy vositalar: qattiq disk (HDD), SSD, USB flesh-disk, SD-kartalar. Masofaviy vositalar: serverlar, bulutli xizmatlar va tarmoq xotira qurilmalari (NAS).",
            "SSD (Solid State Drive) HDD ga nisbatan tezroq ishlaydi, shovqinsiz va zarbalarga chidamli. USB flesh-disklar ma'lumotlarni ko'chirib o'tkazishda qulay. Tashqi qattiq disklar katta hajmdagi ma'lumotlarni saqlash va zaxiralash uchun ishlatiladi.",
            "Turizm korxonalarida ma'lumotlarni saqlash muhim: mijozlar bazasi, bronlash tarixi, moliyaviy hisobotlar, foto va video materiallar. To'g'ri saqlash vositasini tanlash ma'lumotlar xavfsizligi va ularga tez kirish imkonini ta'minlaydi.",
        ]),
        ("Bulutli saqlash xizmatlari va ularning afzalliklari", [
            "Bulutli saqlash — ma'lumotlarni masofaviy serverlarda internet orqali saqlash xizmati. Mashhur xizmatlar: Google Drive, Dropbox, OneDrive, iCloud, Yandex.Disk. Ular istalgan qurilmadan, istalgan joydan ma'lumotlarga kirish imkonini beradi.",
            "Bulutli saqlashning afzalliklari: avtomatik zaxiralash, bir nechta qurilma bilan sinxronlash, hamkorlik imkoniyati (bir faylda bir nechta kishi ishlashi), qo'shimcha apparat sotib olish zaruratining yo'qligi va ma'lumotlarni yo'qotish xavfining kamayishi.",
            "Turizm korxonalari bulutli xizmatlardan hujjatlarni saqlash, xodimlar bilan ulashish, mijozlarga material yuborish va filiallar o'rtasida ma'lumot almashinuvi uchun foydalanadi. Bu ish samaradorligini oshiradi va xarajatlarni kamaytiradi.",
        ]),
        ("Ma'lumotlarni ko'chirish va uzatish texnologiyalari", [
            "Ma'lumotlarni ko'chirish usullari: USB orqali (flesh-disk, kabel), tarmoq orqali (Wi-Fi, Ethernet), bluetooth orqali, elektron pochta va bulutli xizmatlar orqali. Har bir usulning tezligi va qulayligi turlicha.",
            "Katta hajmdagi fayllarni uzatish uchun FTP (File Transfer Protocol), WeTransfer, Google Drive yoki maxsus dasturlar (SHAREit, AirDrop) ishlatiladi. Tarmoq ichida fayllarni ulashish uchun papkaga umumiy kirish (shared folder) sozlanadi.",
            "Zamonaviy usullar: NFC (Near Field Communication) orqali tez ma'lumot almashish, QR-kod orqali havola uzatish va AirDrop (Apple qurilmalari uchun). Turizm sohasida bu usullar mijozlarga ma'lumot berish va xodimlar o'rtasida operativ aloqa uchun foydali.",
        ]),
        ("Axborot xavfsizligi va zaxira nusxalash", [
            "Axborot xavfsizligi — ma'lumotlarni ruxsatsiz kirish, o'zgartirish, o'chirish va tarqalishdan himoyalash. Turizm korxonalarida mijozlarning shaxsiy ma'lumotlari (pasport, bank karta), bronlash ma'lumotlari va moliyaviy hisoblar muhofaza ostida bo'lishi kerak.",
            "Himoyalash vositalari: parollar (murakkab, muntazam o'zgartiriladigan), shifrlaash (encryption), antivirus dasturlari, firewall (tarmoq ekrani) va kirish huquqlarini boshqarish. Har bir xodim faqat o'ziga kerakli ma'lumotlarga kirish imkoniyatiga ega bo'lishi kerak.",
            "Zaxira nusxalash (backup) — ma'lumotlarning nusxasini yaratib, alohida joyda saqlash. Bu kompyuter buzilganda, virus hujumida yoki tasodifiy o'chirishda ma'lumotlarni tiklash imkonini beradi. 3-2-1 qoidasi: 3 nusxa, 2 ta turli vosita, 1 tasi masofaviy.",
        ]),
    ],
})



TOPICS.append({
    "num": 4,
    "title": "Word, Excel ko'rinishida kiritish. Rasmlarni Word hujjatlariga jo'natish",
    "reja": [
        "Microsoft Word da ma'lumotlarni kiritish va formatlash.",
        "Microsoft Excel da jadvallar yaratish va formatlash.",
        "Word hujjatlariga rasm va grafik obyektlarni qo'shish.",
        "Word va Excel integratsiyasi va amaliy qo'llanilishi.",
    ],
    "sections": [
        ("Microsoft Word da ma'lumotlarni kiritish va formatlash", [
            "Microsoft Word turizm sohasida arizalar, shartnomalar, yo'riqnomalar, hisobotlar va boshqa rasmiy hujjatlarni tayyorlash uchun keng ishlatiladi. Matn kiritish klaviatura orqali amalga oshiriladi, formatlash esa shrift, o'lcham, rang va paragraf sozlamalari orqali bajariladi.",
            "Turizm korxonasida Word yordamida turistik marshrutlar tavsifi, mehmonxona xizmatlari ro'yxati, restoran menyusi va reklama materiallarini yaratish mumkin. Uslublar (Styles) qo'llash hujjatga professional ko'rinish beradi.",
            "Hujjatni rasmiylashtirish standartlari: A4 format, Times New Roman 14pt, 1.5 interval, chapdan 3 sm, o'ngdan 1.5 sm maydon. Bu parametrlar rasmiy hujjatlar uchun qabul qilingan talablardir.",
        ]),
        ("Microsoft Excel da jadvallar yaratish va formatlash", [
            "Microsoft Excel turizm sohasida moliyaviy hisobotlar, statistik jadvallar, narxnomalar, xodimlar jadvali va bronlash ro'yxatlarini yaratish uchun ishlatiladi. Ma'lumotlar kataklarga kiritiladi va turli formatlash qo'llaniladi.",
            "Excel da formulalar (SUM, AVERAGE, COUNT, IF) yordamida avtomatik hisob-kitoblar bajariladi. Masalan, mehmonxona daromadini oylar bo'yicha hisoblash, xonalar bandligini foizda ko'rsatish yoki restoran xarajatlarini tahlil qilish.",
            "Diagrammalar (ustunli, chiziqli, doiraviy) ma'lumotlarni vizual ko'rinishda taqdim etadi. Turizm statistikasi, mehmonlar soni dinamikasi yoki daromad taqsimoti diagrammalarda yaqqol ko'rinadi.",
        ]),
        ("Word hujjatlariga rasm va grafik obyektlarni qo'shish", [
            "Word hujjatiga rasm qo'shish 'Qo'yish' (Insert) yorlig'idagi 'Rasm' (Picture) buyrug'i orqali amalga oshiriladi. Kompyuterdan, internetdan yoki kameradan rasmni tanlash mumkin. Rasm o'lchami, joylashuvi va matn bilan munosabati sozlanadi.",
            "Turizm hujjatlarida rasmlar muhim o'rin tutadi: mehmonxona fotosuratlari, turistik joylar rasmlari, restoran interyeri va xizmatlar tasviri. Rasm hujjatni ko'rgazmali va jozibali qiladi, bu esa reklama materiallarida ayniqsa muhim.",
            "Word ga shuningdek SmartArt (tuzilmaviy diagrammalar), Shapes (shakllar), WordArt (badiiy matn) va Screenshot (ekran tasviri) qo'shish mumkin. Bu vositalar korxona tuzilmasi sxemasi, xizmatlar tasnifi yoki yo'nalishlar xaritasini yaratishda foydali.",
        ]),
        ("Word va Excel integratsiyasi va amaliy qo'llanilishi", [
            "Excel jadvali yoki diagrammasini Word hujjatiga qo'yish mumkin: nusxalab joylashtirish (Paste) yoki bog'lab joylashtirish (Paste Link). Bog'lab joylashtirilganda Excel dagi o'zgarishlar Word da avtomatik yangilanadi.",
            "Xat birlashtirish (Mail Merge) — Excel dagi mijozlar ro'yxatidan foydalanib, Word da har bir mijozga personallashtirilgan xat, taklifnoma yoki sertifikat yaratish. Bu turizm korxonalarida ommaviy xatlar yuborishda keng qo'llaniladi.",
            "Amaliy misollar: mehmonxona hisobotini Word da tayyorlab, ichiga Excel jadvallarini joylashtirish; turistik paket narxlarini Excel da hisoblash va natijani Word taklif xatiga ko'chirish; restoran oylik hisobotini ikki dastur integratsiyasi orqali yaratish.",
        ]),
    ],
})

TOPICS.append({
    "num": 5,
    "title": "Prezentatsiyalarni ishlab chiqish va multimediyadan foydalangan holda namoyish etish",
    "reja": [
        "PowerPoint dasturi va prezentatsiya yaratish asoslari.",
        "Multimediya elementlarini prezentatsiyaga qo'shish.",
        "Prezentatsiyani professional darajada rasmiylashtirish.",
        "Turizm sohasida prezentatsiyalarning amaliy qo'llanilishi.",
    ],
    "sections": [
        ("PowerPoint dasturi va prezentatsiya yaratish asoslari", [
            "Microsoft PowerPoint — slaydlar asosida vizual prezentatsiyalar yaratish dasturi. Turizm sohasida u turistik mahsulotlarni taqdim etish, yangi loyihalarni tushuntirish, xodimlarni o'qitish va konferensiyalarda ma'ruza qilish uchun ishlatiladi.",
            "Prezentatsiya yaratish bosqichlari: maqsadni aniqlash, mazmunni rejalashtirish, slaydlar dizaynini tanlash, matn va media kiritish, animatsiya qo'shish va sinovdan o'tkazish. Har bir slaydda bitta asosiy fikr bo'lishi tavsiya etiladi.",
            "Slayd maketi (layout) turlarini to'g'ri tanlash muhim: sarlavha slaydi, mazmun slaydi, rasm slaydi, taqqoslash slaydi va bo'sh slayd. Yagona dizayn shablon (theme) qo'llash prezentatsiyaga professional ko'rinish beradi.",
        ]),
        ("Multimediya elementlarini prezentatsiyaga qo'shish", [
            "Multimediya — bu matn, rasm, audio, video va animatsiyaning birlashmasidan iborat kontentdir. PowerPoint ga turli multimediya elementlarini qo'shish mumkin: fotosuratlar, video roliklar, audio fayllar, GIF animatsiyalar va 3D obyektlar.",
            "Video qo'shish: kompyuterdan video fayl kiritish yoki YouTube havolasini joylashtirish. Audio: fon musiqasi yoki ovozli izoh qo'shish. Bu elementlar prezentatsiyani jonli va qiziqarli qiladi, auditoriya e'tiborini ushlab turadi.",
            "Turizm prezentatsiyalarida multimediya ayniqsa muhim: turistik joylarning video ko'rinishi, mehmonxona virtual turi, restoran atmosferasining audio-video tasviri mijozlarga xizmatni chuqurroq his qilish imkonini beradi.",
        ]),
        ("Prezentatsiyani professional darajada rasmiylashtirish", [
            "Professional prezentatsiyaning xususiyatlari: yagona rang sxemasi (3-4 ta asosiy rang), o'qishga qulay shrift (sarlavha uchun 28-36pt, matn uchun 18-24pt), minimal matn va ko'p vizual material, sifatli rasmlar va aniq diagrammalar.",
            "Dizayn tamoyillari: oddiylik (har bir slaydda bitta fikr), muvozanat (elementlar bir tekis joylashgan), kontrast (muhim narsalar ajralib turadi) va takroriylik (yagona uslub barcha slaydlarda). Bu tamoyillar auditoriya idrokini osonlashtiradi.",
            "Prezentatsiyada ko'p tarqalgan xatolar: juda ko'p matn, mayda shrift, sifatsiz rasmlar, ortiqcha animatsiya va rang uyg'unsizligi. Professional prezentatsiya minimalistik, aniq va maqsadga yo'naltirilgan bo'ladi.",
        ]),
        ("Turizm sohasida prezentatsiyalarning amaliy qo'llanilishi", [
            "Turistik tur paketini taqdim etish: marshrutning fotosuratli ko'rinishi, narxlar jadvali, xizmatlar ro'yxati va aloqa ma'lumotlari slaydlarda joylashtiriladi. Bu mijozlarni jalb qilish va sotuvni oshirish uchun samarali vosita.",
            "Mehmonxona va restoran prezentatsiyasi: xonalar toifalari, interer fotosuratlari, xizmatlar, narxlar va imkoniyatlar ko'rgazmali tarzda taqdim etiladi. Investorlar yoki hamkorlar uchun biznes-plan prezentatsiyasi ham tayyorlanadi.",
            "O'quv va trening prezentatsiyalari: yangi xodimlarni o'qitish, xizmat ko'rsatish standartlarini tushuntirish va xavfsizlik qoidalarini o'rgatish uchun prezentatsiyalar keng ishlatiladi. Interaktiv elementlar (savol-javob, test) samaradorlikni oshiradi.",
        ]),
    ],
})

TOPICS.append({
    "num": 6,
    "title": "Ma'lumotlar asosida slaydlar tayyorlash, animatsiya effektlarini qo'yish",
    "reja": [
        "Ma'lumotlarni slaydlarda vizual ifodalash usullari.",
        "Animatsiya effektlari va ularning turlari.",
        "Slaydlar o'tish effektlari (Transitions).",
        "Animatsiya va effektlarni maqsadga muvofiq qo'llash.",
    ],
    "sections": [
        ("Ma'lumotlarni slaydlarda vizual ifodalash usullari", [
            "Ma'lumotlarni slaydlarda samarali ko'rsatish uchun turli vizualizatsiya usullaridan foydalaniladi: jadvallar, diagrammalar (ustunli, doiraviy, chiziqli), infografikalar, SmartArt sxemalar va ikonkalar. To'g'ri tanlangan vizualizatsiya murakkab ma'lumotni oddiy va tushunarli qiladi.",
            "Turizm statistikasini ko'rsatish: mehmonlar soni dinamikasi chiziqli diagrammada, daromad taqsimoti doiraviy diagrammada, xonalar bandaligi jadvalda, turistik yo'nalishlar xaritada ifodalanadi. Har bir ma'lumot turiga mos vizualizatsiya tanlash muhim.",
            "SmartArt yordamida jarayonlar (bronlash bosqichlari), ierarxiya (korxona tuzilmasi), munosabatlar (xizmatlar o'rtasidagi bog'liqlik) va ro'yxatlar (afzalliklar) ko'rgazmali tarzda taqdim etiladi.",
        ]),
        ("Animatsiya effektlari va ularning turlari", [
            "Animatsiya effektlari — slayd elementlarining harakati bo'lib, ular to'rt turga bo'linadi: kirish (Entrance — element paydo bo'lishi), ta'kidlash (Emphasis — diqqatni jalb qilish), chiqish (Exit — element yo'qolishi) va harakat yo'li (Motion Path — element harakati).",
            "Kirish animatsiyalari: Fade (xiralashib paydo bo'lish), Fly In (uchib kelish), Zoom (kattalashib paydo bo'lish), Bounce (sakrab paydo bo'lish). Ta'kidlash: Pulse (pulsatsiya), Spin (aylanish), Grow/Shrink (kattalashish/kichrayish).",
            "Har bir animatsiyaning parametrlari sozlanadi: davomiyligi (tezligi), kechikish (boshqa animatsiyadan keyin necha soniya), boshlash sharti (sichqoncha bosish, oldingi bilan birga yoki keyin) va takrorlanish soni. Bu parametrlar animatsiyani nazorat qilishni ta'minlaydi.",
        ]),
        ("Slaydlar o'tish effektlari (Transitions)", [
            "Slaydlar o'rtasidagi o'tish (Transition) — bir slayddan ikkinchisiga o'tish paytidagi vizual effektdir. Turlar: oddiy (Cut, Fade), dinamik (Push, Wipe, Split) va murakkab (Morph, 3D effektlar). Yagona o'tish effektini barcha slaydlarga qo'llash tavsiya etiladi.",
            "Morph o'tishi (PowerPoint 2016+) — ikki slayd o'rtasidagi elementlarning silliq transformatsiyasi. Bu xaritada joylashuvni ko'rsatish, fotografiyalar orasida o'tish yoki tuzilmaning o'zgarishini animatsiya qilish uchun samarali vosita.",
            "O'tish effektlarini sozlash: davomiylik (0.5-2 soniya optimal), ovoz effekti (odatda o'chirish tavsiya etiladi) va avtomatik o'tish vaqti (taymer). Prezentatsiya avtomatik rejimda ham namoyish etilishi mumkin (kiosk rejimi).",
        ]),
        ("Animatsiya va effektlarni maqsadga muvofiq qo'llash", [
            "Animatsiyaning asosiy maqsadi — auditoriya e'tiborini muhim nuqtalarga yo'naltirish va ma'lumotni bosqichma-bosqich taqdim etish. Masalan, ro'yxat elementlarini birma-bir ko'rsatish auditoriyani keyingi fikrga tayyorlaydi.",
            "Ortiqcha animatsiya prezentatsiyani bezamasdan, aksincha chalg'itadi va professional ko'rinishini buzadi. Qoida: animatsiya mazmunni qo'llab-quvvatlashi kerak, uning o'zi maqsadga aylanmasligi lozim. Oddiylik va funksionallik ustuvor.",
            "Turizm prezentatsiyalarida samarali animatsiya misollari: xaritada marshrut yo'nalishini bosqichma-bosqich ko'rsatish, mehmonxona xonalarining panoramik ko'rinishi, restoran menyusining bo'limlarini ketma-ket taqdim etish.",
        ]),
    ],
})

TOPICS.append({
    "num": 7,
    "title": "Turistik xizmatlarni sotishda animatsion dasturlardan foydalanish",
    "reja": [
        "Animatsion dasturlar tushunchasi va turlari.",
        "Turizm sohasida animatsiyaning roli.",
        "Turistik xizmatlarni reklama qilishda animatsiya.",
        "Animatsion kontent yaratish vositalari.",
    ],
    "sections": [
        ("Animatsion dasturlar tushunchasi va turlari", [
            "Animatsion dasturlar — bu harakatlanuvchi grafik kontent (animatsiya, video effektlar, interaktiv vizualizatsiya) yaratish uchun mo'ljallangan dasturiy vositalardir. Ular 2D animatsiya (After Effects, Animate), 3D animatsiya (Blender, Maya) va veb-animatsiya (HTML5, CSS, Lottie) turlariga bo'linadi.",
            "Turizm sohasida animatsion dasturlar: reklama bannerlar yaratish, interaktiv xaritalar, virtual turlar, video roliklar va ijtimoiy tarmoqlarga mos animatsion kontent ishlab chiqish uchun ishlatiladi.",
            "Zamonaviy animatsion dasturlarning afzalligi — ular foydalanuvchiga chuqur texnik bilim talab qilmasdan, tayyor shablonlar asosida professional kontent yaratish imkonini beradi. Canva, Powtoon, Vyond kabi platformalar bunga misol.",
        ]),
        ("Turizm sohasida animatsiyaning roli", [
            "Animatsiya turizm sohasida e'tiborni jalb qilish, murakkab ma'lumotni sodda tushuntirish va hissiy ta'sir ko'rsatish uchun ishlatiladi. Harakatlanuvchi kontent statik rasmga qaraganda 5-10 barobar ko'proq e'tibor tortadi.",
            "Turistik marshrutni animatsion xaritada ko'rsatish, mehmonxona xizmatlarini qisqa animatsion video bilan taqdim etish, restoran taomlarini appetitli ko'rinishda tasvirlash — bular animatsiyaning turizmdagi amaliy qo'llanilishi.",
            "Interaktiv animatsiya — foydalanuvchi harakatiga javob beruvchi kontent. Masalan, mehmonxona saytidagi interaktiv xarita, 360-darajali xona ko'rinishi yoki restoran menyusidagi taomni aylantirib ko'rish imkoniyati.",
        ]),
        ("Turistik xizmatlarni reklama qilishda animatsiya", [
            "Ijtimoiy tarmoqlarda animatsion reklama: Instagram Stories, Facebook Ads va TikTok uchun qisqa (15-30 soniya) animatsion videolar yaratish. Ular statik rasmga qaraganda ko'proq o'zaro ta'sir (engagement) oladi.",
            "Veb-saytlarda animatsiya: bosh sahifadagi banner animatsiya, hover effektlari (sichqoncha olib borilganda), yuklanish animatsiyasi va parallaks effektlar. Bu saytni jonli va zamonaviy ko'rsatadi, foydalanuvchi tajribasini yaxshilaydi.",
            "Email marketingda animatsion GIF rasmlar: maxsus takliflar, chegirmalar va yangiliklar animatsion formatda yuborilganda ochilish va bosish ko'rsatkichlari 20-30% ga oshadi. Bu turizm korxonalari uchun samarali marketing vositasi.",
        ]),
        ("Animatsion kontent yaratish vositalari", [
            "Boshlang'ich daraja uchun: Canva (tayyor shablonli animatsiyalar), Powtoon (animatsion prezentatsiyalar), Animoto (foto-video slaydshowlar). Bu dasturlar bepul versiyalarga ega va maxsus bilim talab qilmaydi.",
            "O'rta daraja uchun: Adobe Spark (brendli animatsion kontent), Vyond (tushuntiruvchi animatsion videolar), Lottie (veb-animatsiyalar). Ular ko'proq imkoniyat beradi va professional natija hosil qiladi.",
            "Professional daraja uchun: Adobe After Effects (motion graphics), Adobe Animate (interaktiv animatsiya), Blender (3D animatsiya va virtual turlar). Bu vositalar chuqur bilim talab qiladi, lekin cheksiz imkoniyatlar beradi.",
        ]),
    ],
})

TOPICS.append({
    "num": 8,
    "title": "Jahon davlatlari turizm, mehmonxona biznesida qo'llanilayotgan animatsion dasturlarni namoyish etish",
    "reja": [
        "Xalqaro turizm brendlarining animatsion strategiyalari.",
        "Mehmonxona zanjirlari tomonidan ishlatiladigan animatsion dasturlar.",
        "Virtual reallik va kengaytirilgan reallik turizmdagi qo'llanilishi.",
        "O'zbekiston turizmi uchun xalqaro tajriba saboqlari.",
    ],
    "sections": [
        ("Xalqaro turizm brendlarining animatsion strategiyalari", [
            "Booking.com, Airbnb, TripAdvisor kabi global platformalar o'z saytlari va ilovalarida animatsion elementlardan keng foydalanadi: interaktiv xaritalar, animatsion qidiruv filtrlari, yuklanish animatsiyalari va gamifikatsiya elementlari. Bu foydalanuvchi tajribasini yaxshilaydi.",
            "Aviakompaniyalar (Emirates, Turkish Airlines) animatsion xavfsizlik ko'rsatmalari, interaktiv parvoz xaritalari va brendli animatsion reklama videolar yaratadi. Ular professional animatsiya studiyalari bilan hamkorlik qiladi.",
            "Turistik davlatlar (Yaponiya, Turkiya, Dubay) o'z turizm brendlarini targ'ib qilish uchun yuqori sifatli animatsion videolar, interaktiv veb-saytlar va virtual turlardan foydalanadi. Bular millionlab turistlarni jalb qilishda muhim rol o'ynaydi.",
        ]),
        ("Mehmonxona zanjirlari tomonidan ishlatiladigan animatsion dasturlar", [
            "Marriott, Hilton, Hyatt kabi yirik mehmonxona zanjirlari o'z saytlarida 3D virtual turlar, animatsion bronlash interfeyslari va interaktiv xaritalardan foydalanadi. Bu mijozlarga xonani oldindan ko'rish va tanlash imkonini beradi.",
            "Mehmonxona mobil ilovalarida animatsiya: animatsion xush kelibsiz ekrani, interaktiv xizmatlar menyusi, gamifikatsiya elementlari (ball to'plash animatsiyasi) va push-bildirishnomalar animatsiyasi. Bu mijozlar sadoqatini oshiradi.",
            "Digital signage (raqamli axborot ekranlari) — mehmonxona lobby, lift va yo'laklarda animatsion kontent namoyish etish. Bu restoran menyusi, tadbirlar jadvali va xizmatlar reklamsini dinamik ko'rsatish uchun ishlatiladi.",
        ]),
        ("Virtual reallik va kengaytirilgan reallik turizmdagi qo'llanilishi", [
            "Virtual reallik (VR) — foydalanuvchini to'liq virtual muhitga olib kiradi. Turizm sohasida VR: mehmonxona xonalarini virtual ko'rish, turistik joylarning 360-darajali panoramalari va sayohatni oldindan his qilish tajribasi.",
            "Kengaytirilgan reallik (AR) — real dunyoga virtual elementlarni qo'shadi. Turizm sohasida AR: telefon kamerasi orqali tarixiy joylar haqida ma'lumot ko'rish, restoranda menyuni 3D ko'rinishda ko'rish va navigatsiya uchun AR yo'l ko'rsatish.",
            "Metaverse va immersive tajribalar — kelajakda turistik joylarni virtual dunyoda ziyorat qilish, virtual mehmonxonalarda qolish va virtual sayohat tajribasi ommaviy bo'lishi kutilmoqda. Bu texnologiyalar hozirdan rivojlanmoqda.",
        ]),
        ("O'zbekiston turizmi uchun xalqaro tajriba saboqlari", [
            "O'zbekiston turizmi uchun xalqaro tajribadan olinishi mumkin bo'lgan asosiy saboqlar: sifatli vizual kontent yaratish, interaktiv veb-saytlar ishlab chiqish, ijtimoiy tarmoqlarda animatsion marketing va virtual turlar tayyorlash.",
            "Samarqand, Buxoro, Xiva kabi tarixiy shaharlarning 3D virtual turlari yaratish xalqaro turistlarni jalb qilishda kuchli vosita bo'ladi. 360-darajali panoramalar va drone video materiallari global auditoriyaga yetib boradi.",
            "O'zbekiston mehmonxonalari va restoranlarining veb-saytlarini zamonaviy animatsion elementlar, interaktiv bronlash tizimlari va multimediya kontent bilan boyitish — xalqaro standartlarga yetishning muhim qadami.",
        ]),
    ],
})

TOPICS.append({
    "num": 9,
    "title": "Videoapparatlar bilan ishlash: uyali telefon, LCD proektor, android platforma",
    "reja": [
        "Uyali telefon yordamida video kontent yaratish.",
        "LCD proektorlar va ulardan foydalanish texnologiyasi.",
        "Android platformasi va turizm ilovalari.",
        "Videoapparaturalarni turizm biznesida qo'llash.",
    ],
    "sections": [
        ("Uyali telefon yordamida video kontent yaratish", [
            "Zamonaviy smartfonlar yuqori sifatli video (4K) yozish imkoniyatiga ega bo'lib, turizm sohasida professional kontent yaratish uchun yetarli. Telefon kamerasi yordamida mehmonxona xonalari, restoran taomlarning videolari va turistik joylar haqida reportajlar tayyorlanadi.",
            "Video tahrirlash ilovalari (CapCut, InShot, Adobe Premiere Rush, KineMaster) yordamida yozilgan video kesiladi, musiqa qo'shiladi, matn va effektlar joylashtiriladi. Bu professional ko'rinishdagi video kontent yaratish imkonini beradi.",
            "Jonli efir (Live streaming) — ijtimoiy tarmoqlarda real vaqtda video uzatish. Turizm korxonalari jonli efir orqali mehmonxona ekskursiyasi, restoran ochilishi yoki turistik tadbirni namoyish etishi mumkin. Bu auditoriya bilan bevosita muloqot vositasi.",
        ]),
        ("LCD proektorlar va ulardan foydalanish texnologiyasi", [
            "LCD proektorlar — tasvir va videoni katta ekranga proyeksiya qiluvchi qurilmalar. Turizm sohasida ular prezentatsiyalar namoyishi, o'quv seminarlari, konferensiyalar va marketing tadbirlari uchun ishlatiladi.",
            "Proektor tanlashda muhim parametrlar: yorqinlik (lumen — konferens-zal uchun 3000+ lumen), o'lcham (resolution — Full HD yoki 4K), kontrast nisbati, ulanish portlari (HDMI, VGA, USB, Wi-Fi) va lampa umri.",
            "Proektorni ulash va sozlash: kompyuter yoki telefonga HDMI/USB orqali ulash, ekran o'lchamini sozlash, fokusni to'g'rilash va tasvir sifatini optimallash. Simsiz proektorlar Wi-Fi orqali ulanadi, bu esa qulay va harakatchan.",
        ]),
        ("Android platformasi va turizm ilovalari", [
            "Android — eng keng tarqalgan mobil operatsion tizim bo'lib, turizm sohasida ko'plab ilovalarni qo'llab-quvvatlaydi. Google Play orqali minglab turistik ilovalar yuklab olinadi: navigatsiya, bronlash, tarjimon, valyuta konvertori va yo'riqnoma.",
            "Turizm uchun muhim Android ilovalar: Google Maps (navigatsiya), Booking.com/Airbnb (bronlash), Google Translate (tarjima), TripAdvisor (sharhlar), XE Currency (valyuta) va Flightradar24 (parvozlarni kuzatish).",
            "Turizm korxonalari o'z Android ilovalarini yaratishi mumkin: mehmonxona ilovasi (bronlash, xizmatlar buyurtma qilish), restoran ilovasi (menyu ko'rish, buyurtma berish, to'lash) va turistik ilova (marshrut, gid, bileti). Bu mijozlar qulayligini oshiradi.",
        ]),
        ("Videoapparaturalarni turizm biznesida qo'llash", [
            "Turizm korxonalarida ishlatiladigan video uskunalar: CCTV kameralar (xavfsizlik), veb-kameralar (onlayn konferensiya), action kameralar (ekstremal turizm kontenti) va drone (havoviy video). Har biri o'z maqsadiga xizmat qiladi.",
            "Video konferensiya tizimlari (Zoom, Microsoft Teams, Google Meet) turizm agentliklariga xalqaro hamkorlar bilan aloqa qilish, masofaviy mijozlarga xizmat ko'rsatish va onlayn trening o'tkazish imkonini beradi.",
            "Digital signage tizimlari — mehmonxona va restoranlarda raqamli ekranlarda animatsion reklama, menyu, xizmatlar va tadbirlar haqida ma'lumot namoyish etish. Bu qurilmalar markazlashtirilgan holda masofadan boshqariladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 10,
    "title": "Uyali aloqa protokollaridan foydalanish usullari",
    "reja": [
        "Uyali aloqa texnologiyalari va ularning avlodlari.",
        "Mobil ma'lumot uzatish protokollari.",
        "Turizm sohasida mobil aloqa imkoniyatlari.",
        "Mobil ilovalar va push-bildirishnomalar.",
    ],
    "sections": [
        ("Uyali aloqa texnologiyalari va ularning avlodlari", [
            "Uyali aloqa texnologiyalari avlodlar bo'yicha rivojlangan: 2G (GSM — ovozli aloqa va SMS), 3G (UMTS — mobil internet), 4G/LTE (yuqori tezlikli internet, video qo'ng'iroqlar) va 5G (o'ta yuqori tezlik, past kechikish, IoT qo'llab-quvvatlash).",
            "Har bir avlodning turizm sohasiga ta'siri: 3G — mobil bronlash imkoniyati paydo bo'ldi; 4G — video kontent va real vaqtli xizmatlar rivojlandi; 5G — virtual/kengaytirilgan reallik, IoT asosidagi aqlli mehmonxonalar imkoniyatini ochadi.",
            "O'zbekistonda 4G/LTE tarmog'i keng tarqalgan bo'lib, turistlarga yuqori tezlikli internet xizmati taqdim etiladi. 5G texnologiyasining joriy etilishi turizm xizmatlarini yanada rivojlantiradi.",
        ]),
        ("Mobil ma'lumot uzatish protokollari", [
            "SMS (Short Message Service) — qisqa matnli xabar uzatish protokoli. Turizm sohasida bronlash tasdiqlash, parvoz ma'lumotlari va maxsus takliflarni yuborish uchun ishlatiladi. SMS barcha telefonlarda internet talab qilmasdan ishlaydi.",
            "MMS (Multimedia Messaging Service) — rasm, video va audio fayllarni xabar sifatida yuborish. USSD — tezkor interaktiv xizmatlar (masalan, balans tekshirish). WAP (Wireless Application Protocol) — mobil qurilmalar uchun internet protokoli (hozirda eskirgan).",
            "Zamonaviy protokollar: VoLTE (Voice over LTE — yuqori sifatli ovozli qo'ng'iroq), VoWiFi (Wi-Fi orqali qo'ng'iroq), RCS (Rich Communication Services — SMS ning zamonaviy versiyasi). Bu protokollar aloqa sifatini oshiradi.",
        ]),
        ("Turizm sohasida mobil aloqa imkoniyatlari", [
            "Rouming xizmatlari — turist chet elda ham o'z raqami bilan aloqa qilishi mumkin. Zamonaviy tendensiya — eSIM (raqamli SIM-karta), bu turistga mahalliy tarif rejimidan jismoniy karta almashtirisiz foydalanish imkonini beradi.",
            "Mobil to'lov tizimlari (Apple Pay, Google Pay, Samsung Pay) — turist naqd pulsiz, telefon orqali to'lov qiladi. QR-kodli to'lovlar ham keng tarqalmoqda. Bu mehmonxona va restoranlarda to'lov jarayonini tezlashtiradi.",
            "Geojoylashuv xizmatlari — turistning joylashuvini GPS orqali aniqlash. Bu navigatsiya, yaqin atrofdagi xizmatlarni topish va personallashtirilgan takliflar yuborish uchun ishlatiladi. Mehmonxonalar geojoylashuvga asoslangan reklama yuborishi mumkin.",
        ]),
        ("Mobil ilovalar va push-bildirishnomalar", [
            "Push-bildirishnomalar — mobil ilovadan foydalanuvchi telefoniga yuboriladigan xabarlar. Turizm korxonalari ular orqali maxsus takliflar, bronlash eslatmalari, parvoz holati va chegirmalar haqida xabar beradi. Ular email ga qaraganda 3-10 barobar ko'proq ochiladi.",
            "Geofencing — ma'lum hududga kirgan foydalanuvchiga avtomatik xabar yuborish texnologiyasi. Masalan, turist aeroportga kirganida yaqin mehmonxonalar taklifini olishi yoki mehmonxonaga yaqinlashganda maxsus xush kelibsiz xabarni olishi.",
            "Mobil ilovalar turizm korxonasining barcha xizmatlarini bitta joyga jamlaydigan platforma sifatida ishlaydi: bronlash, to'lov, xizmatlar buyurtmasi, aloqa, sharhlar va sodiqlik dasturi. Bu mijoz bilan doimiy aloqani ta'minlaydi.",
        ]),
    ],
})

TOPICS.append({
    "num": 11,
    "title": "Tarmoqlarning mazmun va mohiyati. Turizm, mehmonxona va restoran xizmatlarini reklama qilish va sotishda ijtimoiy tarmoqning o'rni",
    "reja": [
        "Ijtimoiy tarmoqlar tushunchasi va ularning turlari.",
        "Turizm sohasida ijtimoiy tarmoqlarning roli.",
        "SMM strategiyasi va kontent yaratish.",
        "Ijtimoiy tarmoqlarda reklama va sotish usullari.",
    ],
    "sections": [
        ("Ijtimoiy tarmoqlar tushunchasi va ularning turlari", [
            "Ijtimoiy tarmoqlar — foydalanuvchilarning o'zaro muloqot qilishi, kontent almashishi va jamoa yaratishi uchun mo'ljallangan onlayn platformalardir. Asosiy turlari: umumiy (Facebook, Instagram), kasbiy (LinkedIn), video (YouTube, TikTok), xabar almashish (Telegram, WhatsApp) va sharh platformalari (TripAdvisor).",
            "Har bir ijtimoiy tarmoqning o'z auditoriyasi va kontent formati bor: Instagram — vizual kontent (foto, Stories, Reels), Facebook — matn va jamoa, TikTok — qisqa videolar, YouTube — uzun video, Telegram — kanal va guruhlar, LinkedIn — biznes aloqalar.",
            "O'zbekistonda eng ko'p ishlatiladigan platformalar: Telegram (eng ommaviy), Instagram, Facebook va YouTube. Turizm korxonalari uchun bu platformalarda faol bo'lish — mijozlarga yetib borish va reklama qilishning eng samarali yo'li.",
        ]),
        ("Turizm sohasida ijtimoiy tarmoqlarning roli", [
            "Ijtimoiy tarmoqlar turizm sohasida uchta asosiy vazifani bajaradi: ma'lumot tarqatish (yangiliklar, takliflar), brend qurilishi (taniqlilik, ishonch) va savdo (bronlash, sotish). Mijozlarning 87 foizi sayohat rejalarida ijtimoiy tarmoqlardan ma'lumot oladi.",
            "User Generated Content (UGC) — turistlarning o'zlari yaratgan kontent (foto, video, sharh). Bu eng ishonchli reklama turi bo'lib, boshqa potensial turistlarga kuchli ta'sir ko'rsatadi. Mehmonxona va restoranlar UGC ni rag'batlantirishi kerak.",
            "Influencer marketing — turizm sohasida mashhur blogerlar va sayohatchilar bilan hamkorlik. Ular mehmonxona, restoran yoki turistik joyni o'z auditoriyasiga tavsiya qiladi. Bu an'anaviy reklamadan samarali bo'lishi mumkin.",
        ]),
        ("SMM strategiyasi va kontent yaratish", [
            "SMM (Social Media Marketing) strategiyasi quyidagi bosqichlardan iborat: maqsadni belgilash, auditoriyani aniqlash, platformalarni tanlash, kontent rejasini tuzish, kontent yaratish va natijalarni o'lchash. Har bir bosqich strategiya muvaffaqiyatiga ta'sir qiladi.",
            "Kontent turlari: ma'lumotli (foydali maslahatlar), ilhomlantiruvchi (chiroyli foto/video), o'qituvchi (qanday qilish kerak), ko'ngilochar (viktorinalar, konkurslar) va savdo (takliflar, aksiyalar). 80/20 qoidasi: 80% foydali kontent, 20% savdo kontenti.",
            "Kontent kalendari — oldindan rejalashtirilgan nashrlar jadvali. U muntazam nashr etishni ta'minlaydi, bayramlar va mavsumiy takliflarni o'z vaqtida joylashtiradi. Canva, Later yoki Hootsuite kabi vositalar kontent rejalashtirishni osonlashtiradi.",
        ]),
        ("Ijtimoiy tarmoqlarda reklama va sotish usullari", [
            "Targetli reklama — ijtimoiy tarmoqlarda aniq auditoriyaga (yoshi, joylashuvi, qiziqishlari bo'yicha) yo'naltirilgan pulli reklama. Facebook/Instagram Ads, TikTok Ads va Telegram Ads turizm xizmatlarini maqsadli mijozlarga yetkazadi.",
            "Social commerce — ijtimoiy tarmoqlar orqali to'g'ridan-to'g'ri sotish. Instagram Shopping, Facebook Marketplace va Telegram botlar orqali turistik paketlarni bronlash, mehmonxona xonalarini band qilish va restoran stolini zahiralash mumkin.",
            "Analitika va samaradorlikni o'lchash: obunachilar soni, qamrov (reach), o'zaro ta'sir (engagement), havolaga o'tish (clicks) va konversiya (sotib olish). Bu ko'rsatkichlar reklama strategiyasini takomillashtirish uchun tahlil qilinadi.",
        ]),
    ],
})

TOPICS.append({
    "num": 12,
    "title": "Dasturning imkoniyatlari bilan tanishtirish. SNAPS, STRUCTURE, PSPAR, MultiNet - kichik guruhlardan tarkib topgan ijtimoiy tarmoq dasturlari",
    "reja": [
        "Ijtimoiy tarmoq tahlil dasturlari haqida umumiy ma'lumot.",
        "SNAPS va STRUCTURE dasturlarining imkoniyatlari.",
        "PSPAR va MultiNet dasturlari.",
        "Ijtimoiy tarmoq tahlilini turizm sohasida qo'llash.",
    ],
    "sections": [
        ("Ijtimoiy tarmoq tahlil dasturlari haqida umumiy ma'lumot", [
            "Ijtimoiy tarmoq tahlili (Social Network Analysis — SNA) — bu shaxslar, guruhlar yoki tashkilotlar o'rtasidagi munosabatlar va aloqalarni o'rganish usuli. SNA dasturlari bu aloqalarni vizualizatsiya qiladi va tahlil qilish imkonini beradi.",
            "SNA dasturlari tugunlar (nodes — shaxslar yoki tashkilotlar) va bog'lanishlar (edges — ular o'rtasidagi munosabatlar) asosida tarmoq grafini quradi. Bu graf orqali eng ta'sirli elementlar, klasterlar va axborot oqimlari aniqlanadi.",
            "Turizm sohasida SNA — turistik korxonalar o'rtasidagi hamkorlik tarmoqlari, mijozlar guruhlari o'rtasidagi aloqalar va ma'lumot tarqalish yo'llarini tahlil qilish uchun qo'llaniladi. Bu strategik qarorlar qabul qilishga yordam beradi.",
        ]),
        ("SNAPS va STRUCTURE dasturlarining imkoniyatlari", [
            "SNAPS (Social Network Analysis Procedures) — kichik guruhlar ichidagi ijtimoiy munosabatlarni tahlil qilish uchun mo'ljallangan dastur. U guruh a'zolari o'rtasidagi aloqalar kuchini, markaziylikni va guruh strukturasini aniqlaydi.",
            "STRUCTURE — tarmoq tuzilmasini tahlil qiluvchi dastur bo'lib, u bloklar va pozitsiyalarni aniqlash, ekvivalentlik tahlili va ierarxik klasterlash imkoniyatlariga ega. Bu dastur tashkilot ichidagi rasmiy va norasmiy aloqalarni ko'rsatadi.",
            "Har ikkala dastur ham kiritilgan ma'lumotlar (so'rovnoma natijalari, kuzatuvlar) asosida tarmoq matritsasini quradi va turli ko'rsatkichlarni hisoblaydi: zichlik (density), markaziylik (centrality), klasterlashtirish koeffitsienti va boshqalar.",
        ]),
        ("PSPAR va MultiNet dasturlari", [
            "PSPAR — turli o'lchamdagi guruhlar tuzilmasini tadqiq qilish uchun ishlatiladi. U tarmoqdagi subguruhlarni aniqlaydi, har bir a'zoning rolini belgilaydi va guruh dinamikasini tahlil qiladi. Natijalar jadval va grafik ko'rinishda taqdim etiladi.",
            "MultiNet — bir nechta turdagi aloqalarni bir vaqtda tahlil qiluvchi dastur (multiplex networks). Masalan, bir guruh ichida ham do'stlik, ham ish munosabatlari, ham axborot almashinuvi tarmoqlarini parallel o'rganish mumkin.",
            "Bu dasturlar kichik (10-50 kishilik) guruhlar tahlili uchun optimallashtirilgan. Ular ilmiy tadqiqotlar, tashkilot ichki aloqalari tahlili va jamoaviy ishlash samaradorligini o'rganishda keng qo'llaniladi.",
        ]),
        ("Ijtimoiy tarmoq tahlilini turizm sohasida qo'llash", [
            "Turizm klasterlari tahlili: bir mintaqadagi mehmonxonalar, restoranlar, transport kompaniyalari va turistik operatorlar o'rtasidagi hamkorlik tarmoqlarini tahlil qilish. Bu zaif aloqalarni aniqlash va hamkorlikni kuchaytirishga yordam beradi.",
            "Mijozlar tarmoqi tahlili: mijozlar qanday guruhlarga bo'linadi, qaysi mijozlar ta'sir o'tkazuvchi (influencer), axborot qaysi yo'l bilan tarqaladi. Bu marketing strategiyasini aniqlashtirish va og'zaki reklama (word-of-mouth) samaradorligini oshirish uchun foydali.",
            "Xodimlar aloqasi tahlili: mehmonxona yoki restoran ichidagi rasmiy va norasmiy aloqalar, axborot oqimi va hamkorlik darajasini o'rganish. Bu boshqaruv samaradorligini oshirish va jamoaviy muhitni yaxshilash uchun qo'llaniladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 13,
    "title": "Ma'lumotlarni uzatish protokollari (WAP - Wi-Fi)",
    "reja": [
        "WAP protokoli va uning tarixi.",
        "Wi-Fi texnologiyasi va standartlari.",
        "Bluetooth, NFC va boshqa simsiz protokollar.",
        "Turizm sohasida simsiz texnologiyalarning qo'llanilishi.",
    ],
    "sections": [
        ("WAP protokoli va uning tarixi", [
            "WAP (Wireless Application Protocol) — mobil qurilmalar uchun internetga kirish protokoli bo'lib, 1990-yillar oxirida yaratilgan. U cheklangan imkoniyatli qurilmalar (oddiy telefonlar) uchun veb-sahifalarning soddalashtirilgan versiyasini ko'rsatishga mo'ljallangan edi.",
            "WAP texnologiyasi WML (Wireless Markup Language) tilini ishlatgan va maxsus WAP-gateway orqali ishlagan. Ammo 3G va smartfonlar paydo bo'lgach, WAP eskirdi — zamonaviy qurilmalar to'liq HTML sahifalarni ko'rsata oladi.",
            "WAP ning turizm sohasidagi ahamiyati tarixiy: u mobil bronlash va turistik ma'lumotga mobil qurilma orqali kirish g'oyasining dastlabki amalga oshirilishi edi. Bugungi mobil turizm xizmatlari WAP dan boshlangan evolyutsiyaning natijasi.",
        ]),
        ("Wi-Fi texnologiyasi va standartlari", [
            "Wi-Fi — IEEE 802.11 standartlari asosidagi simsiz mahalliy tarmoq texnologiyasi. Avlodlari: 802.11b/g (2.4 GHz, 54 Mbps), 802.11n/Wi-Fi 4 (300 Mbps), 802.11ac/Wi-Fi 5 (1+ Gbps), 802.11ax/Wi-Fi 6 (10+ Gbps). Har yangi avlod tezlik va samaradorlikni oshiradi.",
            "Wi-Fi tarmoq komponentlari: router (marshrutizator), access point (kirish nuqtasi), modem (internet ulanishi) va adapter (qabul qiluvchi qurilma). Xavfsizlik protokollari: WPA2 va WPA3 — parol bilan himoyalangan ulanish.",
            "Mehmonxona va restoranlarda Wi-Fi zaruriy xizmat hisoblanadi. Mehmonlar uchun bepul Wi-Fi taqdim etish — raqobat afzalligi. Wi-Fi tarmog'i sifati (tezligi, qamrovi, barqarorligi) mehmonlar tajribasiga to'g'ridan-to'g'ri ta'sir qiladi.",
        ]),
        ("Bluetooth, NFC va boshqa simsiz protokollar", [
            "Bluetooth — qisqa masofada (10-100 metr) qurilmalar o'rtasida ma'lumot almashish texnologiyasi. Turizm sohasida: audio gidlar, beacon qurilmalari (yaqin joylashuv xabarlari) va simsiz quloqchinlar uchun ishlatiladi.",
            "NFC (Near Field Communication) — juda qisqa masofada (4 sm gacha) ma'lumot almashish. Kontaktsiz to'lov (bank kartalari, Apple/Google Pay), elektron kalitlar (mehmonxona xona kaliti) va ma'lumot uzatish (NFC teglari) uchun ishlatiladi.",
            "RFID (Radio Frequency Identification) — radio to'lqinlar orqali identifikatsiya. Mehmonxonalarda: xona kaliti kartalari, bagaj kuzatuvi, inventar boshqarish. Restoranlarda: buyurtma kuzatuvi va mahsulot nazorati uchun qo'llaniladi.",
        ]),
        ("Turizm sohasida simsiz texnologiyalarning qo'llanilishi", [
            "Smart hotel (aqlli mehmonxona): mehmon telefonidan Wi-Fi orqali xona haroratini boshqarish, Bluetooth kalit bilan eshikni ochish, NFC orqali to'lov qilish va mobil ilova orqali xizmatlarni buyurtma qilish — bular simsiz texnologiyalar integratsiyasi.",
            "Beacon texnologiyasi — mehmonxona yoki muzeyda joylashtirilgan kichik qurilmalar. Turist yaqinlashganda Bluetooth orqali telefonga ma'lumot yuboradi: tarixiy joy haqida ma'lumot, yaqin atrofdagi restoran takliflari yoki maxsus chegirmalar.",
            "Wi-Fi analytics — mehmonxona Wi-Fi tarmog'i orqali mehmonlar harakati, ko'p tashrif buyuriladigan joylar va o'rtacha qolish vaqtini tahlil qilish. Bu ma'lumotlar xizmat sifatini oshirish va marketing strategiyasini optimallash uchun ishlatiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 14,
    "title": "Internetda adresatsiyalash tartibi. Internetdan ma'lumotni izlash va nusxa ko'chirish",
    "reja": [
        "Internet manzillash tizimi: IP, domen va URL.",
        "Qidiruv tizimlari va samarali qidirish usullari.",
        "Internetdan ma'lumot olish va nusxa ko'chirish.",
        "Internet xavfsizligi va mualliflik huquqlari.",
    ],
    "sections": [
        ("Internet manzillash tizimi: IP, domen va URL", [
            "Internetdagi har bir qurilma noyob IP-manzilga ega: IPv4 (masalan, 192.168.1.1) yoki IPv6. Domen nomi (masalan, booking.com) IP-manzilning inson uchun qulay shakli. DNS (Domain Name System) domen nomini IP-manzilga aylantiradi.",
            "URL (Uniform Resource Locator) — internet resursi manzilining to'liq shakli: protokol://domen/yo'l. Masalan: https://www.mehmonxona.uz/xonalar/lyuks.html. Bu yerda https — protokol, www.mehmonxona.uz — domen, /xonalar/lyuks.html — sahifa yo'li.",
            "Domen zonalari: .com (tijorat), .org (tashkilot), .uz (O'zbekiston), .travel (turizm). Turizm korxonalari uchun mos domen tanlash muhim: qisqa, esda qoladigan va faoliyatni aks ettiruvchi. Masalan: silkroad-tours.uz yoki grand-hotel.uz.",
        ]),
        ("Qidiruv tizimlari va samarali qidirish usullari", [
            "Qidiruv tizimlari: Google (global lider), Yandex (MDH mamlakatlari), Bing (Microsoft). Ular milliardlab veb-sahifalarni indekslaydi va foydalanuvchi so'roviga mos natijalarni tartiblaydi. Google turizm qidiruvlarida ham eng ko'p ishlatiladi.",
            "Samarali qidirish usullari: qo'shtirnoqlar (\"aniq ibora\"), minus (-istalmagan so'z), sayt ichida qidirish (site:booking.com), fayl turini qidirish (filetype:pdf), OR operatori va vaqt bo'yicha filtrlash. Bu usullar kerakli ma'lumotni tezroq topishga yordam beradi.",
            "Turizm sohasida qidiruv: mehmonxona izlash, aviabilet narxlarini taqqoslash, turistik joylar haqida ma'lumot, viza talablari va ob-havo ma'lumotlari. Meta-qidiruv tizimlari (Google Flights, Kayak, Skyscanner) bir nechta manbadan narxlarni taqqoslaydi.",
        ]),
        ("Internetdan ma'lumot olish va nusxa ko'chirish", [
            "Veb-sahifadan ma'lumot olish usullari: matn nusxalash (Ctrl+C), rasmni saqlash (o'ng tugma - Save Image), sahifani to'liq saqlash (Ctrl+S) va PDF formatida saqlash (Print - Save as PDF). Skrinshotlar ham ma'lumotni saqlashning qulay usuli.",
            "Fayllarni yuklab olish (download): brauzer orqali havolani bosish, yuklab olish menejeri (Download Manager) dan foydalanish. Katta fayllarni yuklab olishda IDM (Internet Download Manager) yoki bepul alternativlari (Free Download Manager) ishlatiladi.",
            "Veb-sahifalarni oflayn ko'rish uchun saqlash, xatcho'plar (bookmarks) orqali kerakli sahifalarni yig'ish va bulutli xizmatlarga (Pocket, Evernote) maqolalarni saqlash — bu ma'lumotlarni tartibga solish va keyinchalik foydalanish usullari.",
        ]),
        ("Internet xavfsizligi va mualliflik huquqlari", [
            "Internet xavfsizligi: ishonchli saytlarni tanish (https, kalit belgisi), fishing xabarlardan ehtiyot bo'lish, kuchli parollar ishlatish, ikki bosqichli autentifikatsiya va umumiy Wi-Fi da VPN ishlatish. Turizm sohasida onlayn firibgarlik (soxta bronlash saytlari) keng tarqalgan.",
            "Mualliflik huquqlari (copyright): internetdagi kontentning ko'pchiligi himoyalangan. Rasmlar, matnlar va videolarni ruxsatsiz ishlatish qonunbuzarlik. Bepul foydalanish uchun Creative Commons litsenziyali yoki stok-fotobank (Unsplash, Pexels) materiallaridan foydalanish kerak.",
            "Turizm korxonalari veb-kontentni yaratishda mualliflik huquqlariga rioya qilishi zarur: faqat litsenziyalangan rasmlar, o'z original kontenti va ruxsat berilgan materiallar ishlatilishi kerak. Bu huquqiy muammolarni oldini oladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 15,
    "title": "Sayt dizayni va imkoniyatlari. Saytda amalga oshirilayotgan biznes jarayonlarni o'rganib chiqish",
    "reja": [
        "Veb-sayt tushunchasi va uning turizm sohasidagi ahamiyati.",
        "Sayt dizayni tamoyillari va UX/UI asoslari.",
        "Turizm saytlarida amalga oshiriluvchi biznes jarayonlar.",
        "Sayt yaratish platformalari va vositalari.",
    ],
    "sections": [
        ("Veb-sayt tushunchasi va uning turizm sohasidagi ahamiyati", [
            "Veb-sayt — internetdagi veb-sahifalar to'plami bo'lib, korxona haqida ma'lumot berish, xizmatlarni taqdim etish va mijozlar bilan aloqa o'rnatish uchun xizmat qiladi. Turizm sohasida sayt — korxonaning raqamli yuzi va asosiy sotish kanali.",
            "Turizm korxonasi veb-sayti vazifalari: ma'lumot berish (xizmatlar, narxlar, joylashuv), bronlash imkoniyati, vizual taqdimot (foto, video), mijozlar sharhlari va aloqa. Sayt sifati bevosita sotish hajmiga ta'sir qiladi.",
            "Statistika ko'rsatishicha, turistlarning 70% dan ortig'i sayohat rejalarini internetda o'rganadi va 50% dan ko'prog'i onlayn bronlash qiladi. Shu sababli sifatli, tez ishlaydigan va mobil qurilmalarga moslashgan sayt — turizm korxonasi muvaffaqiyatining zaruriy sharti.",
        ]),
        ("Sayt dizayni tamoyillari va UX/UI asoslari", [
            "UX (User Experience) — foydalanuvchi tajribasi: sayt qanchalik qulay, tushunarli va samarali ishlashi. UI (User Interface) — foydalanuvchi interfeysi: saytning vizual ko'rinishi, ranglar, shriftlar va elementlar joylashuvi. Ikkalasi birgalikda sayt sifatini belgilaydi.",
            "Sayt dizayni tamoyillari: oddiylik (ortiqcha elementlarsiz), navigatsiya qulayligi (3 bosqichda istalgan sahifaga yetish), mobil moslik (responsive design), tezkor yuklanish (3 soniyadan kam) va vizual ierarxiya (muhim narsalar birinchi ko'rinadi).",
            "Turizm saytlari dizaynida muhim: sifatli fotografiyalar (asosiy jalb vositasi), aniq narxlar, oson bronlash tugmasi (CTA — Call to Action), ishonch belgilari (sertifikatlar, sharhlar) va ko'p tillilik (xalqaro turistlar uchun).",
        ]),
        ("Turizm saytlarida amalga oshiriluvchi biznes jarayonlar", [
            "Onlayn bronlash jarayoni: xonani/turni tanlash, sanalarni belgilash, ma'lumotlarni kiritish, to'lovni amalga oshirish va tasdiqlash xatini olish. Bu jarayon to'liq avtomatlashtirilgan bo'lib, 24/7 ishlaydi va inson ishtirokini talab qilmaydi.",
            "CRM integratsiyasi: sayt orqali kelgan so'rovlar avtomatik ravishda mijozlar bazasiga tushadi, ular bilan aloqa tarixi saqlanadi va personallashtirilgan takliflar yuboriladi. Bu mijozlar bilan uzoq muddatli munosabatlar qurishga yordam beradi.",
            "Analitika va optimallashtirish: Google Analytics orqali saytga tashrif buyuruvchilar soni, manbalari, xulq-atvori va konversiya darajasi kuzatiladi. Bu ma'lumotlar asosida sayt doimiy ravishda takomillashtiriladi, bu esa sotishni oshiradi.",
        ]),
        ("Sayt yaratish platformalari va vositalari", [
            "Tayyor platformalar (CMS): WordPress (eng ommaviy, moslashuvchan), Wix (sodda, vizual quruvchi), Tilda (dizaynga yo'naltirilgan), Squarespace (premium dizayn). Ular dasturlash bilimisiz sayt yaratish imkonini beradi.",
            "Turizm uchun maxsus platformalar: mehmonxonalar uchun — Cloudbeds, Little Hotelier; restoranlar uchun — Gloria Food, ChowNow; turlar uchun — Rezdy, Bookeo. Ular sohaga mos tayyor funksionallikni taqdim etadi.",
            "Sayt yaratish bosqichlari: maqsad va auditoriyani aniqlash, domen va hosting tanlash, platforma tanlash, dizayn ishlab chiqish, kontent kiritish, SEO optimallash, test o'tkazish va nashr etish. Saytni muntazam yangilab turish ham muhim.",
        ]),
    ],
})

TOPICS.append({
    "num": 16,
    "title": "Turistik korxonalar, mehmonxona, restoran saytlarini tahlil qilish",
    "reja": [
        "Sayt tahlilining maqsadi va mezonlari.",
        "Turistik korxonalar saytlarini baholash.",
        "Mehmonxona va restoran saytlarini baholash.",
        "Sayt tahlili natijalari asosida takliflar ishlab chiqish.",
    ],
    "sections": [
        ("Sayt tahlilining maqsadi va mezonlari", [
            "Sayt tahlili — bu veb-saytning texnik, dizayn va funksional jihatlarini baholab, uning samaradorligini aniqlash va yaxshilash yo'llarini topish jarayoni. Tahlil maqsadi — saytning biznes maqsadlariga qanchalik xizmat qilishini o'lchash.",
            "Tahlil mezonlari: dizayn va vizual ko'rinish, navigatsiya qulayligi, kontent sifati va dolzarbligi, mobil moslik, yuklanish tezligi, SEO optimallashi, xavfsizlik (SSL), bronlash funksionalligi va ijtimoiy tarmoqlar integratsiyasi.",
            "Tahlil vositalari: Google PageSpeed Insights (tezlik), GTmetrix (performance), Mobile-Friendly Test (mobil moslik), SEMrush/Ahrefs (SEO), Hotjar (foydalanuvchi xulq-atvori). Bu vositalar aniq raqamlar va tavsiyalar beradi.",
        ]),
        ("Turistik korxonalar saytlarini baholash", [
            "Turistik operatorlar saytida muhim elementlar: turlar katalogi (filtrlash imkoniyati bilan), narxlar va paket tarkibi, fotogalereya va videolar, online bronlash, to'lov tizimlari, mijozlar sharhlari va aloqa ma'lumotlari.",
            "Muvaffaqiyatli turizm saytlari namunalari: aniq qidiruv va filtrlash tizimi, ilhomlantiruvchi vizual kontent, personallashtirilgan takliflar, blog bo'limi (foydali maslahatlar) va sodiqlik dasturi. Bu elementlar mijozni saytda ushlab turadi.",
            "Xatolar: eskirgan ma'lumotlar, sekin yuklanish, murakkab navigatsiya, mobil versiyaning yo'qligi, bronlash jarayonining uzunligi va aloqa ma'lumotlarining topilmasligi. Bu xatolar mijozlarni yo'qotishga olib keladi.",
        ]),
        ("Mehmonxona va restoran saytlarini baholash", [
            "Mehmonxona sayti uchun zarur: xona turlari tavsifi va fotosuratlari, narxlar va mavjudlik kalendari, onlayn bronlash, joylashuv xaritasi, qo'shimcha xizmatlar (restoran, spa, konferens-zal) va mehmonlar sharhlari.",
            "Restoran sayti uchun zarur: menyu (narxlari bilan), interer fotosuratlari, ish vaqti va manzil, stol bronlash imkoniyati, delivery/olib ketish xizmati va maxsus takliflar (bayramlar, banketlar). Menyu PDF emas, balki saytda ko'rinadigan bo'lishi kerak.",
            "Taqqosiy tahlil usuli: bir nechta raqobatchi saytlarni jadvalda baholash (har bir mezon bo'yicha 1-5 ball). Bu o'z saytining kuchli va zaif tomonlarini aniqlash va yaxshilash uchun aniq rejani ishlab chiqish imkonini beradi.",
        ]),
        ("Sayt tahlili natijalari asosida takliflar ishlab chiqish", [
            "Tahlil natijalari hisobot shaklida rasmiylashtiriladi: joriy holat tavsifi, aniqlangan muammolar ro'yxati, har bir muammo uchun yechim taklifi va amalga oshirish ustuvorligi (yuqori, o'rta, past). Bu hisobot saytni yaxshilash uchun yo'l xaritasi bo'ladi.",
            "Texnik takliflar: hosting yaxshilash (tezlik uchun), SSL sertifikat o'rnatish (xavfsizlik), responsive dizayn joriy etish (mobil moslik) va SEO optimallash (qidiruv tizimlarida ko'rinish). Bu asosiy texnik masalalar.",
            "Kontent va funksional takliflar: professional fotografiya va video olish, kontent yangilash jadvali tuzish, onlayn bronlash tizimini joriy etish, chat-bot qo'shish va ijtimoiy tarmoqlar integratsiyasini kuchaytirish.",
        ]),
    ],
})

TOPICS.append({
    "num": 17,
    "title": "O'zbekiston turistik korxonalari, mehmonxonalarida zamonaviy bronlash tizimlaridan foydalanish",
    "reja": [
        "Bronlash tizimi tushunchasi va turlari.",
        "O'zbekistonda qo'llaniladigan bronlash platformalari.",
        "Mehmonxonalarda bronlash tizimlarini joriy etish.",
        "Onlayn bronlash jarayoni va uning afzalliklari.",
    ],
    "sections": [
        ("Bronlash tizimi tushunchasi va turlari", [
            "Bronlash tizimi — xizmatlarni (xona, tur, transport) oldindan band qilish imkonini beruvchi dasturiy ta'minotdir. Turlari: to'g'ridan-to'g'ri bronlash (korxonaning o'z sayti), OTA — Online Travel Agencies (Booking.com, Expedia) va GDS — Global Distribution Systems (Amadeus, Sabre).",
            "Channel Manager — bir nechta bronlash platformasidagi mavjudlik va narxlarni bir joydan boshqarish tizimi. U mehmonxonaga Booking, Airbnb, Expedia va o'z saytidagi bronlarni sinxron holda boshqarish imkonini beradi, ikkilantirilgan bronlash xavfini bartaraf etadi.",
            "Bronlash tizimlari real vaqtda ishlaydi: xona band qilinganda barcha platformalarda darhol yangilanadi. Bu ortiqcha bronlash (overbooking) muammosini hal qiladi va inventar boshqaruvini optimallashtiradi.",
        ]),
        ("O'zbekistonda qo'llaniladigan bronlash platformalari", [
            "Xalqaro platformalar: Booking.com (eng ommaviy), Airbnb (xususiy turar joylar), Expedia, Agoda va TripAdvisor. O'zbekistondagi mehmonxonalarning aksariyati Booking.com da ro'yxatga olingan va u orqali xalqaro mehmonlarni qabul qiladi.",
            "Mahalliy platformalar: MyTaxi (transport), Uzairways.com (aviabilet), mehmonxonalarning o'z saytlari. O'zbekistonda mahalliy bronlash platformalari rivojlanish bosqichida — bu sohada yangi startaplar va loyihalar paydo bo'lmoqda.",
            "Turizm operatorlari uchun: iVisa (viza xizmatlari), GetYourGuide (ekskursiyalar), Viator (tajribalar). O'zbekiston turizm korxonalari ushbu platformalarda o'z xizmatlarini joylashtirish orqali xalqaro bozorga chiqishi mumkin.",
        ]),
        ("Mehmonxonalarda bronlash tizimlarini joriy etish", [
            "Bronlash tizimini joriy etish bosqichlari: ehtiyojlarni aniqlash, tizim tanlash, texnik tayyorgarlik (internet, kompyuter), tizimni sozlash (xonalar, narxlar, qoidalar kiritish), xodimlarni o'qitish va sinov ishlashi.",
            "Tizim tanlash mezonlari: narxi (oylik to'lov yoki komissiya), qo'llab-quvvatlash tili, integratsiya imkoniyatlari (PMS, buxgalteriya), foydalanish qulayligi, mobil ilova mavjudligi va mijozlar sharhlari.",
            "Kichik mehmonxonalar uchun oddiy va arzon yechimlar: Cloudbeds, Little Hotelier, Beds24. Katta mehmonxonalar uchun to'liq tizimlar: Opera (Oracle), Protel, Mews. Har bir mehmonxona o'z hajmi va byudjetiga mos tizim tanlashi kerak.",
        ]),
        ("Onlayn bronlash jarayoni va uning afzalliklari", [
            "Onlayn bronlash jarayoni mijoz nuqtai nazaridan: 1) saytga kirish, 2) sanalar va mehmonlar sonini tanlash, 3) mavjud xonalarni ko'rish, 4) xonani tanlash, 5) shaxsiy ma'lumotlarni kiritish, 6) to'lovni amalga oshirish, 7) tasdiqlash emailini olish.",
            "Afzalliklari mijoz uchun: 24/7 mavjudlik, bir nechta variantni taqqoslash imkoniyati, vaqtni tejash, real vaqtda narxlarni ko'rish va darhol tasdiqlash. Afzalliklari korxona uchun: avtomatlashtirilgan jarayon, xodim vaqtini tejash va xatolarni kamaytirish.",
            "Bronlash tizimi ma'lumotlarini tahlil qilish: qaysi kanal ko'proq bronlash olib keladi, o'rtacha qolish muddati, oldindan bronlash muddati va bekor qilish darajasi. Bu ma'lumotlar narx strategiyasi va marketing rejasini optimallashtirish uchun ishlatiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 18,
    "title": "Turizm, mehmonxona xizmatlarini bronlashning global tizimlarini (GDS) o'rganish. Amadeus, Sabre, Fidelio, Galileo tizimlarining ishlash prinsiplari",
    "reja": [
        "GDS tushunchasi va uning turizm sohasidagi o'rni.",
        "Amadeus va Sabre tizimlarining ishlash prinsiplari.",
        "Galileo va Worldspan tizimlari.",
        "Fidelio (Opera) mehmonxona boshqaruv tizimi.",
    ],
    "sections": [
        ("GDS tushunchasi va uning turizm sohasidagi o'rni", [
            "GDS (Global Distribution System) — aviakompaniyalar, mehmonxonalar, rent-a-car va boshqa turizm xizmatlari inventarini butun dunyo bo'ylab turli agentliklarga real vaqtda taqdim etuvchi markazlashtirilgan elektron tizim. GDS turizm tarqatish zanjirining asosini tashkil etadi.",
            "GDS ning ishlash prinsipi: xizmat ko'rsatuvchi (aviakompaniya, mehmonxona) o'z inventarini GDS ga joylashtiradi, turli mamlakatdagi travel agentlar esa GDS orqali bu xizmatlarni qidiradi, bronlaydi va to'lovni amalga oshiradi. Bu millionlab tranzaksiyalarni real vaqtda bajaradi.",
            "GDS bozori asosan to'rtta yirik tizimdan iborat: Amadeus (Yevropa va global), Sabre (Shimoliy Amerika), Travelport (Galileo + Worldspan). Ular birgalikda dunyo bo'ylab kuniga millionlab bronlashlarni qayta ishlaydi.",
        ]),
        ("Amadeus va Sabre tizimlarining ishlash prinsiplari", [
            "Amadeus — Ispaniyada joylashgan va 190 dan ortiq mamlakatda ishlatiladigan eng yirik GDS. U aviabiletlar, mehmonxonalar, temir yo'l, avtomobil ijarasi va sug'urta xizmatlarini bronlash imkonini beradi. Amadeus Selling Platform — agentlar uchun asosiy ish interfeysi.",
            "Sabre — AQShda joylashgan GDS bo'lib, asosan Shimoliy va Janubiy Amerikada kuchli pozitsiyaga ega. Sabre Red 360 — agentlar uchun zamonaviy veb-asosidagi interfeys. U tez qidirish, narxlarni taqqoslash va murakkab marshrutlarni bronlash imkoniyatini beradi.",
            "Har ikkala tizim ham maxsus buyruqlar (cryptic commands) va grafik interfeys (GUI) orqali ishlaydi. Agentlar maxsus o'quv kurslaridan o'tib sertifikat oladi. Zamonaviy versiyalari API integratsiyasi va veb-interfeys orqali ishlash imkonini beradi.",
        ]),
        ("Galileo va Worldspan tizimlari", [
            "Galileo — Travelport kompaniyasiga tegishli GDS bo'lib, asosan Yevropa, Osiyo va Tinch okeani mintaqasida keng tarqalgan. U aviabiletlar, mehmonxonalar va avtomobil ijarasi xizmatlarini taqdim etadi. Galileo Desktop va Smartpoint interfeyslari ishlatiladi.",
            "Worldspan — ham Travelport tarkibiga kiruvchi GDS, asosan onlayn travel agentliklar (OTA) uchun texnologik yechimlar taqdim etadi. U Orbitz, Priceline kabi yirik OTA larning texnologik asosini tashkil etadi.",
            "Travelport o'z GDS larini Universal API orqali birlashtirib, dasturchilarga barcha uchta tizim (Galileo, Worldspan, Apollo) inventariga yagona interfeys orqali kirish imkonini beradi. Bu zamonaviy integratsiya yondashuvidir.",
        ]),
        ("Fidelio (Opera) mehmonxona boshqaruv tizimi", [
            "Fidelio/Opera PMS (Property Management System) — Oracle kompaniyasiga tegishli, dunyodagi eng keng tarqalgan mehmonxona boshqaruv tizimi. U xonalar inventari, bronlash, ro'yxatga olish (check-in), hisob-kitob va hisobot modullaridan iborat.",
            "Opera tizimining asosiy modullari: Reservation (bronlash), Front Desk (qabul xizmati), Cashiering (kassa), Housekeeping (xonadon xizmati), AR/AP (debitor/kreditor), Reports (hisobotlar) va Profiles (mijozlar bazasi).",
            "Opera Cloud — zamonaviy bulutli versiya bo'lib, mahalliy server o'rnatishni talab qilmaydi. U mobil qurilmalardan ishlash, boshqa tizimlar bilan oson integratsiya va avtomatik yangilanish imkoniyatlarini taqdim etadi. Bu kichik va o'rta mehmonxonalar uchun ham qulay yechim.",
        ]),
    ],
})

TOPICS.append({
    "num": 19,
    "title": "Turizm, mehmonxona va restoran xizmatlarini sotishda reklamaning o'rni (logotip, grafik va video reklama yaratish)",
    "reja": [
        "Reklama tushunchasi va turizm sohasidagi ahamiyati.",
        "Logotip yaratish va brending asoslari.",
        "Grafik reklama materiallari tayyorlash.",
        "Video reklama yaratish texnologiyalari.",
    ],
    "sections": [
        ("Reklama tushunchasi va turizm sohasidagi ahamiyati", [
            "Reklama — mahsulot yoki xizmat haqida maqsadli auditoriyaga axborot yetkazish va ularni harakatga undash vositasi. Turizm sohasida reklama turistik mahsulotni tanishtirisn, brend taniqligini oshirish va sotishni rag'batlantirish uchun zarur.",
            "Turizm reklamasining o'ziga xosligi: u hissiy ta'sirga ko'proq tayanadi (orzu, zavq, yangilik hissi), vizual kontent muhim rol o'ynaydi va mavsumiylikka bog'liq. Samarali turizm reklamasi mijozda sayohat qilish ishtiyoqini uyg'otadi.",
            "Reklama kanallari: raqamli (ijtimoiy tarmoqlar, qidiruv tizimlari, email, saytlar), an'anaviy (gazeta, jurnal, TV, radio, tashqi reklama). Zamonaviy turizm korxonalari asosan raqamli reklamaga yo'nalgan, chunki u aniqroq maqsadga yo'naltiriladi va natijasi o'lchanadigan.",
        ]),
        ("Logotip yaratish va brending asoslari", [
            "Logotip — korxona vizual identifikatsiyasining asosiy elementi, uning simvoli va yuz belgisi. Yaxshi logotip: sodda (esda qoladi), o'ziga xos (farqlanadi), moslashuvchan (turli o'lcham va fonda ishlaydi) va korxona faoliyatini aks ettiradi.",
            "Logotip yaratish vositalari: professional — Adobe Illustrator, CorelDRAW; oddiy — Canva, LogoMaker, Hatchful. Logotip uchun shrift, rang va shakl tanlash muhim: ranglar psixologiyasi (ko'k — ishonch, yashil — tabiat, qizil — energiya) hisobga olinadi.",
            "Brend identifikatsiyasi logotipdan tashqari: rang palitrasi (asosiy va qo'shimcha ranglar), shrift oilasi, vizual uslub, ovoz toni (tone of voice) va brand guideline (qo'llanma). Bu elementlarning yagona qo'llanilishi kuchli brend yaratadi.",
        ]),
        ("Grafik reklama materiallari tayyorlash", [
            "Grafik reklama turlari: banner (veb-sayt va ijtimoiy tarmoq uchun), buklet va brosyura (bosma), flayer (tarqatma varaqalar), poster (afisha) va vizitka. Har bir turning o'z o'lchami, formati va tarqatish usuli bor.",
            "Grafik dizayn dasturlari: Adobe Photoshop (rasmlar tahrirlash), Adobe Illustrator (vektorli grafika), Canva (onlayn dizayn), Figma (veb-dizayn). Canva boshlang'ich darajadagi foydalanuvchilar uchun tayyor shablonlar bilan ishlash imkonini beradi.",
            "Samarali grafik reklama tamoyillari: bitta aniq xabar, kuchli vizual, o'qishga qulay shrift, brend ranglari va aniq harakatga chaqiruv (CTA — Call to Action). Turizm reklamasida sifatli fotografiya va ilhomlantiruvchi tasvir hal qiluvchi rol o'ynaydi.",
        ]),
        ("Video reklama yaratish texnologiyalari", [
            "Video reklama turlari: promo-rolik (30-60 sek), imij-video (1-3 min), sharh-video (mijozlar fikri), behind-the-scenes (sahna ortida), virtual tur (360 daraja). Har bir tur o'z maqsadi va auditoriyasiga ega.",
            "Video yaratish bosqichlari: ssenariy yozish (script), suratga olish (filming), tahrirlash (editing), musiqa va effektlar qo'shish va nashr etish. Smartfon bilan ham sifatli video olish mumkin — muhimi yorug'lik, barqaror kamera va yaxshi audio.",
            "Video tahrirlash dasturlari: Adobe Premiere Pro (professional), Final Cut Pro (Mac), DaVinci Resolve (bepul, kuchli), CapCut/InShot (mobil). Turizm video reklamasi uchun drone suratlar, time-lapse va silliq o'tishlar samarali texnikalar.",
        ]),
    ],
})

TOPICS.append({
    "num": 20,
    "title": "Reklama materiallarini loyihalarini ishlab chiqish, tayyorlash va taqdimot qilish",
    "reja": [
        "Reklama loyihasini rejalashtirish bosqichlari.",
        "Reklama materiallarini tayyorlash texnologiyasi.",
        "Reklama materiallarini taqdimot qilish usullari.",
        "Reklama kampaniyasi samaradorligini baholash.",
    ],
    "sections": [
        ("Reklama loyihasini rejalashtirish bosqichlari", [
            "Reklama loyihasini ishlab chiqish quyidagi bosqichlarni o'z ichiga oladi: brief (topshiriq) olish, maqsadli auditoriyani aniqlash, xabar va kreativ konsepsiyani ishlab chiqish, kanal va formatlarni tanlash, byudjetni belgilash va ish jadvalini tuzish.",
            "Brief — bu buyurtmachi tomonidan berilgan topshiriq bo'lib, unda reklama maqsadi, maqsadli auditoriya, asosiy xabar, byudjet va muddatlar ko'rsatiladi. To'g'ri brief — muvaffaqiyatli reklama loyihasining asosidir.",
            "Kreativ konsepsiya — reklama g'oyasining asosi. Turizm reklamasida konsepsiya odatda hissiy tajribaga (sarguzasht, dam olish, kashfiyot) asoslanadi. Masalan: 'O'zbekiston — Ipak yo'lining yuragi' — bu konsepsiya tarix va romantikani birlashtiradi.",
        ]),
        ("Reklama materiallarini tayyorlash texnologiyasi", [
            "Bosma materiallar tayyorlash: dizaynni yaratish (Photoshop/Canva), matn yozish (copywriting), tasvirlarni tanlash, maket (layout) tuzish va bosmaxonaga topshirish. Bosma uchun rasm sifati kamida 300 DPI, rang modeli CMYK bo'lishi kerak.",
            "Raqamli materiallar tayyorlash: veb-bannerlar (GIF, HTML5), ijtimoiy tarmoq postlari (platformaga mos o'lcham), email shablon va video roliklar. Raqamli kontent uchun RGB rang modeli, 72-150 DPI va platforma talablariga mos o'lchamlar ishlatiladi.",
            "Kontent yaratish jarayoni: tadqiqot (auditoriya va raqobatchilar), ideatsiya (g'oyalar ishlab chiqish), prototip (eskiz/maket), tasdiqlash (buyurtmachi bilan kelishish), ishlab chiqarish (final versiya) va tarqatish (nashr etish). Har bir bosqich sifatni ta'minlaydi.",
        ]),
        ("Reklama materiallarini taqdimot qilish usullari", [
            "Buyurtmachiga taqdimot qilish: PowerPoint yoki PDF formatida loyihani tushuntirish — konsepsiya, maqsadli auditoriya, kanallar, materiallar namunalari va kutilayotgan natijalar. Vizual material (mockup) real ko'rinishni tasavvur qilishga yordam beradi.",
            "Mockup — bu reklama materialnig real muhitdagi ko'rinishini simulyatsiya qilish. Masalan: bilbord fotosuratingiz ko'cha fonida, bukletingiz qo'lda, banneringiz veb-saytda. Bu buyurtmachiga yakuniy natijani oldindan ko'rish imkonini beradi.",
            "Taqdimot muvaffaqiyati omillari: aniq va qisqa bayon, kuchli vizual materiallar, auditoriyani tushunish, savollarga tayyorgarlik va professional muomala. Taqdimotning birinchi 30 soniyasi — diqqatni jalb qilishning eng muhim lahzasi.",
        ]),
        ("Reklama kampaniyasi samaradorligini baholash", [
            "KPI (Key Performance Indicators) — reklama muvaffaqiyatini o'lchovchi ko'rsatkichlar: qamrov (reach), taassurotlar (impressions), bosishlar (clicks), konversiya (conversion), investitsiya qaytimi (ROI) va brend taniqligining oshishi.",
            "Raqamli reklama samaradorligini o'lchash oson: Google Ads va Facebook Ads o'z analitikasini beradi. Bosma va tashqi reklama uchun so'rovnomalar, promo-kodlar va maxsus telefon raqamlaridan foydalaniladi.",
            "A/B testing — reklama materialining ikki variantini taqqoslab, qaysi biri samaraliroq ekanini aniqlash. Masalan: ikki xil sarlavha, ikki xil rasm yoki ikki xil CTA. Natijalar asosida eng samarali variant tanlanadi va kengaytiriladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 21,
    "title": "Elektron to'lovlarni amalga oshirish vositalari",
    "reja": [
        "Elektron to'lov tushunchasi va turlari.",
        "Bank kartalari va onlayn to'lov tizimlari.",
        "Mobil to'lov va elektron hamyonlar.",
        "Turizm sohasida elektron to'lovlarning qo'llanilishi.",
    ],
    "sections": [
        ("Elektron to'lov tushunchasi va turlari", [
            "Elektron to'lov — bu naqd pulsiz, elektron vositalar orqali amalga oshiriladigan moliyaviy tranzaksiya. Turlar: bank kartasi orqali (Visa, MasterCard, UzCard, Humo), elektron hamyon orqali (Payme, Click, PayPal), mobil to'lov (Apple Pay, Google Pay) va bank o'tkazmasi.",
            "Elektron to'lov tizimining komponentlari: to'lovchi (mijoz), qabul qiluvchi (savdogar), to'lov shlyuzi (gateway — tranzaksiyani qayta ishlaydi), protsessing markazi (bank bilan aloqa) va emitent bank (karta chiqargan bank).",
            "To'lov xavfsizligi standartlari: PCI DSS (karta ma'lumotlarini himoyalash), 3D Secure (qo'shimcha tasdiqlash), SSL/TLS (ma'lumot shifrlash) va tokenizatsiya (karta raqamini vaqtinchalik token bilan almashtirish). Bu standartlar firibgarlikni oldini oladi.",
        ]),
        ("Bank kartalari va onlayn to'lov tizimlari", [
            "Bank kartalari turlari: debet karta (hisobdagi mablag' doirasida), kredit karta (qarz mablag'i), prepaid karta (oldindan to'ldirilgan). O'zbekistonda UzCard va Humo mahalliy to'lov tizimlari, Visa va MasterCard xalqaro tizimlar sifatida ishlaydi.",
            "Onlayn to'lov shlyuzlari: PayMe, Click, Apelsin (O'zbekistonda), Stripe, PayPal (xalqaro). Ular veb-sayt yoki ilovaga o'rnatiladi va mijozga to'lovni xavfsiz amalga oshirish imkonini beradi. Komissiya odatda 1-3% ni tashkil etadi.",
            "Turizm saytlarida to'lov integratsiyasi: bronlash tugmasidan keyin to'lov sahifasiga o'tish, karta ma'lumotlarini kiritish, 3D Secure tasdiqlash va to'lov kvitansiyasini olish. Bu jarayon 2-3 daqiqa ichida tugaydi va 24/7 ishlaydi.",
        ]),
        ("Mobil to'lov va elektron hamyonlar", [
            "Mobil to'lov tizimlari: Payme, Click, Apelsin (O'zbekistonda eng ommaviy). Ular orqali telefon raqami yoki QR-kod skanerlash orqali tez va qulay to'lov amalga oshiriladi. Mehmonxona va restoranlarda QR-kodli to'lov keng tarqalmoqda.",
            "NFC to'lovlar (Apple Pay, Google Pay, Samsung Pay): telefon yoki soatni terminal ga tegintirib to'lash. Bu eng tez to'lov usuli (1-2 soniya). O'zbekistonda bu xizmatlar bosqichma-bosqich joriy etilmoqda.",
            "Elektron hamyon (e-wallet): Payme, Click hamyoni, PayPal — ularda mablag' saqlanadi va turli to'lovlar amalga oshiriladi. Turizm sohasida elektron hamyonlar xalqaro to'lovlarni osonlashtiradi va valyuta konvertatsiyasini avtomatlaydi.",
        ]),
        ("Turizm sohasida elektron to'lovlarning qo'llanilishi", [
            "Mehmonxonalarda: xona bronlash to'lovi (oldindan yoki chiqishda), qo'shimcha xizmatlar to'lovi (restoran, spa, minibar), depozit va garov to'lovlari. PMS tizimi barcha to'lovlarni bir hisobda birlashtiradi.",
            "Restoranlarda: stol boshida to'lov (POS terminal), onlayn buyurtma to'lovi (delivery), QR-kodli menyu orqali to'lov va choy puli (tip) elektron to'lovi. Split payment — bir hisobni bir nechta kishi o'rtasida bo'lish imkoniyati ham mavjud.",
            "Turistik xizmatlar uchun: tur paketlari uchun onlayn to'lov (bo'lib to'lash imkoniyati bilan), aviabilet va transport to'lovlari, sug'urta va viza to'lovlari. Xalqaro turistlar uchun ko'p valyutali to'lov qabul qilish muhim.",
        ]),
    ],
})

TOPICS.append({
    "num": 22,
    "title": "Terminallar va kassa apparatlarining imkoniyatlari. Ularning ishlash prinsiplari",
    "reja": [
        "POS-terminal tushunchasi va turlari.",
        "Kassa apparatlari va fiskal qurilmalar.",
        "POS-tizimlar va ularning funksionalligi.",
        "Turizm va restoran sohasida POS-tizimlarning qo'llanilishi.",
    ],
    "sections": [
        ("POS-terminal tushunchasi va turlari", [
            "POS-terminal (Point of Sale) — savdo nuqtasida bank kartalari orqali to'lovni qabul qiluvchi elektron qurilma. Turlari: statsionar (do'kon yoki resepshen uchun), portativ (ofitsiant stol boshiga olib boradi) va mobil (smartfonga ulanadigan kichik qurilma — mPOS).",
            "Ishlash prinsipi: karta terminalga kiritiladi yoki tegintiriladi, terminal karta ma'lumotlarini o'qiydi, protsessing markaziga so'rov yuboradi, bank tranzaksiyani tasdiqlaydi va chek chiqariladi. Bu jarayon 3-10 soniya davom etadi.",
            "Zamonaviy terminallar kontaktsiz to'lovni (NFC), QR-kodli to'lovni, pin-kodsiz to'lovni (kichik summalar uchun) va choy pulini kiritish imkoniyatini qo'llab-quvvatlaydi. Ba'zilari Android operatsion tizimida ishlaydi va qo'shimcha ilovalar o'rnatilishi mumkin.",
        ]),
        ("Kassa apparatlari va fiskal qurilmalar", [
            "Kassa apparati (KKM) — savdo operatsiyalarini qayd etuvchi va fiskal chek chiqaruvchi qurilma. O'zbekistonda barcha savdo va xizmat ko'rsatish nuqtalari fiskal kassa ishlatishi shart (soliq nazorati maqsadida).",
            "Fiskal xotira moduli — barcha tranzaksiyalarni o'chirib bo'lmaydigan holda saqlaydigan maxsus qurilma. U soliq inspeksiyasiga ma'lumot uzatadi. Onlayn-kassa esa real vaqtda ma'lumotni soliq organiga yuboradi (O'zbekistonda soliq.uz tizimi).",
            "Chek ma'lumotlari: korxona nomi va STIR, sana va vaqt, tovar/xizmat nomi, soni va narxi, QQS summasi, to'lov usuli (naqd/karta) va fiskal belgi. QR-kodli chek mijozga elektron shaklda ham yuborilishi mumkin.",
        ]),
        ("POS-tizimlar va ularning funksionalligi", [
            "POS-tizim — bu faqat to'lov qabul qiluvchi terminal emas, balki savdo va xizmat ko'rsatishni to'liq boshqaruvchi dasturiy-apparat kompleks. U kassa, inventar boshqaruvi, xodimlar hisobi, mijozlar bazasi va hisobotlarni birlashtiradi.",
            "POS-tizim tarkibi: kompyuter yoki planshet (asosiy qurilma), kassa tortmasi (naqd pul uchun), chek printer, shtrix-kod skaneri, POS-terminal (karta uchun) va dasturiy ta'minot. Bulutli POS-tizimlar server o'rnatishni talab qilmaydi.",
            "POS-tizim funksiyalari: tovar/xizmatlar katalogi, buyurtma qabul qilish, turli to'lov usullari, chegirmalar va aksiyalar boshqarish, smena ochish/yopish, xodimlar kirish huquqlari va batafsil hisobotlar (kunlik, oylik, kategoriya bo'yicha).",
        ]),
        ("Turizm va restoran sohasida POS-tizimlarning qo'llanilishi", [
            "Restoran POS-tizimlari: iiko, R-Keeper, Poster, Square. Ular buyurtmani qabul qilish, oshxonaga yuborish, stol boshqaruvi, menyu boshqarish, inventar nazorati va xodimlar vaqtini hisobga olish funksiyalariga ega.",
            "Mehmonxona POS-tizimlari PMS (Property Management System) bilan integratsiyalashadi: mehmon restoranda ovqatlanganda, hisob avtomatik uning xona hisobiga qo'shiladi. Chiqishda (check-out) barcha xarajatlar bitta hisobda ko'rsatiladi.",
            "Zamonaviy tendensiyalar: QR-kodli menyu va buyurtma (telefon orqali), self-service kiosk (o'z-o'ziga xizmat), ofitsiant planshetidan buyurtma va mobil POS (portativ terminal). Bu texnologiyalar xizmat tezligini oshiradi va xodim ish yukini kamaytiradi.",
        ]),
    ],
})

TOPICS.append({
    "num": 23,
    "title": "Turistik ofislarda ishlatiladigan boshqaruv tizimlari va dasturiy ta'minot",
    "reja": [
        "Turistik ofis boshqaruv tizimi tushunchasi.",
        "CRM tizimlari va mijozlar bilan ishlash.",
        "Turlarni shakllantirish va sotish dasturlari.",
        "Ofis ishi va hisobot avtomatlashtirish dasturlari.",
    ],
    "sections": [
        ("Turistik ofis boshqaruv tizimi tushunchasi", [
            "Turistik ofis boshqaruv tizimi — bu turoperator yoki turagent faoliyatini avtomatlashtiruvchi dasturiy ta'minot. U turlarni shakllantirish, narxlash, bronlash, hujjatlarni rasmiylashtirish, moliyaviy hisoblar va mijozlar bilan ishlashni birlashtiradi.",
            "Boshqaruv tizimi afzalliklari: qo'lda ishni kamaytirish, xatolarni oldini olish, tezkorlik, shaffoflik va hisobotlar orqali qarorlar qabul qilish. U kichik agentlikdan tortib yirik operatorgacha barcha darajadagi korxonalar uchun zarur.",
            "Mashhur turistik boshqaruv tizimlari: SAMO-TourAgent, U-ON Travel, Bitrix24 (CRM), TourControl va boshqalar. Har birining o'z imkoniyatlari va narx siyosati bor. Tizim tanlashda korxona hajmi va ehtiyojlari hisobga olinadi.",
        ]),
        ("CRM tizimlari va mijozlar bilan ishlash", [
            "CRM (Customer Relationship Management) — mijozlar bilan munosabatlarni boshqarish tizimi. U mijozlar bazasini yuritadi, aloqa tarixini saqlaydi, sotish jarayonini kuzatadi va personallashtirilgan xizmat ko'rsatishni ta'minlaydi.",
            "Turizm uchun CRM funksiyalari: mijoz profili (shaxsiy ma'lumotlar, afzalliklar, sayohat tarixi), lead boshqarish (potensial mijozlarni sotuvga aylantirish), avtomatik eslatmalar (tug'ilgan kun, sayohat yildonaligi) va email/SMS kampaniyalar.",
            "Mashhur CRM tizimlar: Bitrix24 (keng imkoniyatli, o'zbek tili bor), AmoCRM (sodda va qulay), Salesforce (xalqaro standart), HubSpot (bepul boshlang'ich versiya). CRM joriy etish mijozlar qaytish darajasini 25-30% ga oshirishi mumkin.",
        ]),
        ("Turlarni shakllantirish va sotish dasturlari", [
            "Tur shakllantirish dasturlari: marshrutni rejalashtirish, xizmatlarni (transport, turar joy, ovqatlanish, ekskursiya) birlashtirish, narxni kalkulyatsiya qilish va paket tuzish. Bu jarayonni avtomatlashtirish turoperator samaradorligini oshiradi.",
            "Onlayn sotish kanallari: korxona veb-sayti (to'g'ridan-to'g'ri bronlash), OTA platformalari (GetYourGuide, Viator), B2B platformalar (boshqa agentlarga sotish) va ijtimoiy tarmoqlar. Ko'p kanalli sotish (omnichannel) strategiyasi daromadni maksimallashtiradi.",
            "Narxlash va yield management: mavsumga, talabga va raqobatga qarab narxlarni dinamik boshqarish. Dasturlar o'tmishdagi ma'lumotlar asosida optimal narxni tavsiya qiladi. Erta bronlash chegirmasi va oxirgi daqiqa takliflari ham avtomatlashtiriladi.",
        ]),
        ("Ofis ishi va hisobot avtomatlashtirish dasturlari", [
            "Ofis ishi avtomatlashtirishga: shartnomalar generatsiyasi (shablon asosida), hisob-fakturalar yaratish, to'lovlarni kuzatish, yetkazib beruvchilar bilan hisob-kitob va arxivlash kiradi. Bu qog'oz ishini kamaytiradi va tezlikni oshiradi.",
            "Hisobot tizimlari: kunlik sotish hisoboti, oylik moliyaviy hisobot, xodimlar samaradorligi, marshrutlar mashhurligi va mijozlar statistikasi. Dashboardlar (boshqaruv panellari) asosiy ko'rsatkichlarni real vaqtda ko'rsatadi.",
            "Integratsiya: boshqaruv tizimi buxgalteriya dasturi (1C, QuickBooks), email xizmati, telefoniya va messenger bilan integratsiyalashishi kerak. Bu ma'lumotlarning bir tizimdan boshqasiga avtomatik oqishini ta'minlaydi.",
        ]),
    ],
})

TOPICS.append({
    "num": 24,
    "title": "Ofis dasturlari yordamida korxona faoliyatini yuritish. Dasturlar bilan ishlash",
    "reja": [
        "Microsoft Office paketi va uning komponentlari.",
        "Google Workspace va bulutli ofis vositalari.",
        "Hujjat aylanishini avtomatlashtirish.",
        "Turizm ofisida kundalik dasturiy vositalar.",
    ],
    "sections": [
        ("Microsoft Office paketi va uning komponentlari", [
            "Microsoft Office — eng keng tarqalgan ofis dasturlari paketi: Word (hujjatlar), Excel (jadvallar va hisob-kitoblar), PowerPoint (prezentatsiyalar), Outlook (email va kalendar), Access (ma'lumotlar bazasi) va OneNote (eslatmalar).",
            "Turizm ofisida Word: shartnomalar, xatlar, yo'riqnomalar; Excel: narxnomalar, hisobotlar, jadvallar; PowerPoint: taqdimotlar, trening; Outlook: mijozlar va hamkorlar bilan yozishmalar, uchrashuv rejalash. Har bir dastur o'z sohasida kuchli.",
            "Microsoft 365 (bulutli versiya) — internet orqali ishlash, bir nechta qurilmada sinxronlash, real vaqtda hamkorlikda ishlash va avtomatik saqlash imkoniyatlarini beradi. Bu zamonaviy ofis ishi uchun standart yechim.",
        ]),
        ("Google Workspace va bulutli ofis vositalari", [
            "Google Workspace: Google Docs (matn), Google Sheets (jadval), Google Slides (prezentatsiya), Gmail (email), Google Drive (saqlash), Google Calendar (kalendar) va Google Meet (video konferensiya). Hammasi brauzer orqali ishlaydi, o'rnatish talab etilmaydi.",
            "Afzalliklari: bepul boshlang'ich versiya, real vaqtda hamkorlik (bir faylda bir nechta kishi ishlashi), avtomatik saqlash (hech narsa yo'qolmaydi), istalgan qurilmadan kirish va sharhlar orqali muloqot.",
            "Turizm ofisida Google Workspace: jamoaviy ish rejasi (Sheets), mijozlarga taklifnoma (Docs), xodimlar uchrashuvi (Calendar + Meet), reklama materiallari (Slides) va fayllarni ulashish (Drive). Bu kichik korxonalar uchun ayniqsa qulay.",
        ]),
        ("Hujjat aylanishini avtomatlashtirish", [
            "Hujjat aylanish tizimi (DMS — Document Management System) — hujjatlarni yaratish, saqlash, qidirish, tasdiqlash va arxivlash jarayonini avtomatlashtiradigan dastur. Bu qog'oz hujjatlarning raqamli muqobiliga o'tishni ta'minlaydi.",
            "Turizm ofisida hujjat aylanishi: shartnomalar (yaratish - tasdiqlash - imzolash - saqlash), hisobotlar (tayyorlash - tekshirish - tasdiqlash), kiruvchi xatlar (ro'yxatga olish - bajaruvchiga yo'naltirish - javob berish - arxiv). Har bir bosqich kuzatiladi.",
            "Elektron imzo — hujjatni raqamli shaklda tasdiqlash usuli. O'zbekistonda E-IMZO tizimi orqali rasmiy hujjatlarga huquqiy kuchga ega elektron imzo qo'yish mumkin. Bu masofaviy shartnomalar imzolash va tezkor hujjat almashinuvi imkonini beradi.",
        ]),
        ("Turizm ofisida kundalik dasturiy vositalar", [
            "Aloqa va muloqot: Telegram (tezkor xabar, kanallar), Zoom/Google Meet (video uchrashuv), email (rasmiy yozishma). Turizm agentligida mijoz, hamkorlar va xodimlar bilan doimiy aloqa zarur.",
            "Rejalashtirish va loyiha boshqarish: Trello (kanboan taxtasi), Notion (hamma narsa bitta joyda), Google Calendar (jadval) va Todoist (vazifalar ro'yxati). Bu vositalar jamoaviy ishni tartibga soladi va muddatlarni nazorat qiladi.",
            "Moliya va buxgalteriya: 1C (buxgalteriya hisobi), Excel (oddiy hisob-kitoblar), bank ilovalari (to'lovlarni kuzatish). Kichik agentliklar uchun oddiy Excel jadvali yetarli bo'lishi mumkin, katta korxonalar esa maxsus dasturlar ishlatadi.",
        ]),
    ],
})

TOPICS.append({
    "num": 25,
    "title": "Mehmonxona faoliyatini boshqarishda qo'llaniladigan dasturlar",
    "reja": [
        "PMS (Property Management System) tushunchasi va vazifalari.",
        "Mashhur mehmonxona boshqaruv dasturlari.",
        "PMS ning asosiy modullari va funksiyalari.",
        "PMS tizimini tanlash va joriy etish.",
    ],
    "sections": [
        ("PMS (Property Management System) tushunchasi va vazifalari", [
            "PMS — mehmonxona boshqaruv tizimi bo'lib, u xonalar fondi, bronlash, qabul xizmati, xonadon, moliyaviy operatsiyalar va hisobotlarni yagona platformada birlashtiradi. Bu mehmonxona kundalik operatsiyalarining asosiy dasturiy ta'minoti.",
            "PMS ning asosiy vazifalari: xonalar mavjudligini boshqarish, bronlashlarni qabul qilish va tasdiqlash, mehmonlarni ro'yxatga olish va chiqarish, hisob-kitob yuritish, xonadon xizmatini boshqarish va boshqaruv hisobotlari yaratish.",
            "PMS boshqa tizimlar bilan integratsiyalanadi: Channel Manager (bronlash kanallari), POS (restoran), CRM (mijozlar), buxgalteriya dasturi, elektron kalit tizimi va minibar tizimi. Bu yaxlit ekotizim barcha jarayonlarni birlashtiradi.",
        ]),
        ("Mashhur mehmonxona boshqaruv dasturlari", [
            "Katta mehmonxonalar uchun: Opera PMS (Oracle) — xalqaro standart, Protel (Planet), Mews — zamonaviy bulutli tizim. Bu tizimlar kuchli funksionallik, ko'p tillilik va xalqaro integratsiyalarni taqdim etadi.",
            "O'rta va kichik mehmonxonalar uchun: Cloudbeds (bulutli, hamyonbop), Little Hotelier (kichik mehmonxonalarga maxsus), Hotelogix, RoomRaccoon va Beds24. Ular oddiy interfeys, tez sozlash va past narx bilan ajralib turadi.",
            "O'zbekistonda qo'llaniladigan tizimlar: Opera (yirik mehmonxonalar), 1C:Hotel (mahalliy yechim), Bnovo va Travelline (MDH bozori uchun). Tizim tanlash mehmonxona hajmi, byudjeti va xizmatlar doirasiga bog'liq.",
        ]),
        ("PMS ning asosiy modullari va funksiyalari", [
            "Reservation (bronlash moduli): bronlarni qabul qilish, o'zgartirish va bekor qilish; guruh bronlash; kutish ro'yxati (waitlist); allotment boshqarish. Front Desk (qabul): check-in/check-out, xona tayinlash, mehmon profili va maxsus so'rovlar.",
            "Housekeeping (xonadon moduli): xonalar holati (toza, iflos, ta'mirda), xonadon xodimlariga vazifa berish, minibar nazorati. Cashiering (kassa): mehmon hisobi, to'lovlar qabul qilish, hisob-faktura yaratish va folio boshqarish.",
            "Reports (hisobotlar): bandlik darajasi (occupancy rate), o'rtacha kunlik narx (ADR), xona boshiga daromad (RevPAR), prognoz va tarixiy taqqoslash. Bu ko'rsatkichlar mehmonxona samaradorligini o'lchaydi va strategik qarorlarni asoslaydi.",
        ]),
        ("PMS tizimini tanlash va joriy etish", [
            "Tanlash mezonlari: funksionallik (kerakli modullar bormi), foydalanish qulayligi, integratsiya imkoniyatlari, narxi (bir martalik yoki oylik), texnik qo'llab-quvvatlash, o'quv materiallar mavjudligi va mos platformalar (bulutli yoki lokal).",
            "Joriy etish bosqichlari: ehtiyojlarni aniqlash, tizim tanlash, shartnoma imzolash, texnik tayyorgarlik, ma'lumotlarni ko'chirish (migration), xodimlarni o'qitish, sinov ishlashi va to'liq ishga tushirish. Bu jarayon 2-8 hafta davom etishi mumkin.",
            "Muvaffaqiyat omillari: rahbariyat qo'llab-quvvatlashi, xodimlarning tayyor bo'lishi, bosqichma-bosqich joriy etish va doimiy texnik yordam. Xodimlar eski usuldan yangisiga o'tishda qo'shimcha motivatsiya va sabr-toqat talab qilinadi.",
        ]),
    ],
})

TOPICS.append({
    "num": 26,
    "title": "Kichik va xususiy mehmonxonalarning boshqaruv tizimlari va dasturlari",
    "reja": [
        "Kichik mehmonxonalarning o'ziga xos xususiyatlari.",
        "Kichik mehmonxonalar uchun mos boshqaruv dasturlari.",
        "Bepul va arzon dasturiy yechimlar.",
        "Kichik mehmonxonalarda texnologiya joriy etish strategiyasi.",
    ],
    "sections": [
        ("Kichik mehmonxonalarning o'ziga xos xususiyatlari", [
            "Kichik mehmonxonalar (boutique hotel, guest house, hostel, B&B) odatda 5-50 xonaga ega, kam xodimlar bilan ishlaydi va cheklangan byudjetga ega. Shu sababli ular uchun oddiy, arzon va tez o'rganiladi gan dasturiy yechimlar kerak.",
            "Kichik mehmonxonalarning afzalligi — shaxsiy yondashuv va moslashuvchanlik. Texnologiya bu afzallikni yo'qotmasdan, kundalik tartibli ishlarni (bronlash, hisob-kitob, tozalik nazorati) avtomatlashtirib, egasi va xodimlarning mehmonlarga ko'proq vaqt ajratishini ta'minlashi kerak.",
            "O'zbekistonda kichik va xususiy mehmonxonalar soni tez o'smoqda — turistik shaharlar (Samarqand, Buxoro, Xiva, Toshkent) da ko'plab family guest house va boutique hotellar ochilmoqda. Ularga mos texnologik yechimlar talab ortmoqda.",
        ]),
        ("Kichik mehmonxonalar uchun mos boshqaruv dasturlari", [
            "Cloudbeds — kichik va o'rta mehmonxonalar uchun all-in-one bulutli tizim: PMS + Channel Manager + Booking Engine. Sodda interfeys, ko'p tilli va 300+ kanal integratsiyasi. Narxi xona soniga qarab oylik to'lov.",
            "Little Hotelier (SiteMinder) — 1-30 xonali mehmonxonalar uchun maxsus yaratilgan: oddiy PMS, kanal boshqaruvchi va onlayn bronlash mexanizmi. Mobil ilovasi bor — egasi telefondan boshqarishi mumkin.",
            "Boshqa variantlar: Beds24 (bepul plan mavjud), eZee Absolute (keng funksional), RoomRaccoon (avtomatizatsiya kuchli) va Hostelworld (hostellar uchun maxsus). Har biri kichik mehmonxonaning muayyan ehtiyojlariga javob beradi.",
        ]),
        ("Bepul va arzon dasturiy yechimlar", [
            "Bepul yechimlar: Google Calendar (oddiy bronlash jadvali), Google Sheets (mehmonlar ro'yxati va hisoblar), Beds24 bepul plan (1-2 xona uchun). Juda kichik (1-5 xonali) mehmonxonalar uchun bu yetarli bo'lishi mumkin.",
            "Arzon yechimlar (oyiga 20-50 dollar): Cloudbeds starter, Little Hotelier, Lodgify. Ular professional darajadagi funksionallikni kichik narxda taqdim etadi. Ko'pchiligida 14-30 kunlik bepul sinov muddati bor.",
            "Telegram bot — O'zbekistonda kichik mehmonxonalar uchun arzon alternativa: bronlash so'rovlarini qabul qilish, mehmonlarga avtomatik javob berish va ma'lumot yuborish. Oddiy botni dasturchi 1-2 kunda yaratishi mumkin.",
        ]),
        ("Kichik mehmonxonalarda texnologiya joriy etish strategiyasi", [
            "Bosqichma-bosqich yondashuv: 1-bosqich — OTA larda (Booking.com) ro'yxatdan o'tish; 2-bosqich — oddiy PMS joriy etish; 3-bosqich — o'z saytida onlayn bronlash; 4-bosqich — CRM va marketing avtomatlash. Bir vaqtda hammasini qilishga urinmaslik kerak.",
            "Xodimlarni o'qitish: kichik mehmonxonada ko'pincha egasi va 1-2 xodim ishlaydi. Ular uchun sodda video qo'llanmalar, amaliy mashg'ulotlar va doimiy qo'llab-quvvatlash muhim. Murakkab tizimlardan qochish va intuitiv dasturlarni tanlash kerak.",
            "Investitsiya qaytimi (ROI): texnologiya xarajatlarining o'zini qoplashini hisoblash. Masalan: OTA komissiyasi 15-20%, o'z saytidan to'g'ridan-to'g'ri bronlash — 0-5%. Agar tizim oyiga 5+ to'g'ridan-to'g'ri bronlash olib kelsa, o'zini tez qoplaydi.",
        ]),
    ],
})

TOPICS.append({
    "num": 27,
    "title": "Mijozlarni ro'yxatga olish va bronlarni bekor qilish",
    "reja": [
        "Mehmonxonada ro'yxatga olish (check-in) jarayoni.",
        "Elektron ro'yxatga olish va self-service texnologiyalari.",
        "Bronlarni bekor qilish siyosati va jarayoni.",
        "No-show va overbooking boshqaruvi.",
    ],
    "sections": [
        ("Mehmonxonada ro'yxatga olish (check-in) jarayoni", [
            "Check-in — mehmonning mehmonxonaga kelganda ro'yxatga olinishi va xona kalitini olish jarayoni. Bosqichlari: 1) bronlashni tekshirish (PMS da qidirish), 2) shaxsiy hujjatni tekshirish (pasport), 3) ro'yxatga olish kartasini to'ldirish, 4) to'lovni tasdiqlash, 5) xona kalitini berish.",
            "PMS tizimida check-in: mehmon ismi yoki bron raqami bo'yicha bronlash topiladi, xona tayinlanadi (assign), mehmon ma'lumotlari kiritiladi, depozit olinadi (kerak bo'lsa), kalit aktivlashtiriladi va xona holati 'band' (occupied) ga o'zgartiriladi.",
            "Chet ellik mehmonlar uchun qo'shimcha talab: O'zbekistonda xorijiy fuqarolarning ro'yxatga olish (registratsiya) ma'lumotlari migratsiya xizmatiga yuborilishi kerak. Ba'zi PMS tizimlar bu jarayonni avtomatlashtiradi.",
        ]),
        ("Elektron ro'yxatga olish va self-service texnologiyalari", [
            "Online check-in — mehmon kelishidan oldin (odatda 24-48 soat) mobil ilova yoki email orqali ro'yxatga olish. Mehmon shaxsiy ma'lumotlarni, to'lov ma'lumotlarini kiritadi va kelganda faqat kalitni oladi (yoki mobil kalit ishlatadi).",
            "Self-service kiosk — mehmonxona lobbidagi terminal qurilma. Mehmon o'zi bron raqami yoki passport skaneri orqali check-in qiladi, karta to'lovini amalga oshiradi va kalit oladi. Bu qabul xizmatidagi navbatni kamaytiradi.",
            "Mobil kalit (digital key) — mehmon telefoni xona kaliti sifatida ishlaydi (Bluetooth/NFC orqali). Bu qabul stoliga umuman bormasdan xonaga kirish imkonini beradi. Marriott, Hilton va Hyatt kabi zanjirlar bu texnologiyani faol joriy etmoqda.",
        ]),
        ("Bronlarni bekor qilish siyosati va jarayoni", [
            "Bekor qilish siyosati — bronlashni bekor qilish shartlarini belgilaydi: bepul bekor qilish muddati (masalan, kelishdan 24-48 soat oldin), jarima miqdori (birinchi kecha narxi yoki 100%), qaytarib bo'lmaydigan (non-refundable) bronlash va force majeure holatlari.",
            "PMS da bekor qilish jarayoni: bron topiladi, bekor qilish sababi tanlanadi, tizim siyosatga ko'ra jarima hisoblaydi (yoki hisoblamaydi), xona qaytadan mavjud (available) holatga o'tadi va mijozga tasdiqlash emaili yuboriladi.",
            "Bekor qilish tahlili: qaysi kanaldan kelgan bronlar ko'proq bekor qilinadi, qaysi mavsumda bekor qilish yuqori va o'rtacha bekor qilish muddati qancha. Bu ma'lumotlar siyosatni optimallashtirish va daromadni himoyalash uchun ishlatiladi.",
        ]),
        ("No-show va overbooking boshqaruvi", [
            "No-show — mehmonning bron qilib kelmasligi. Bu mehmonxona uchun daromad yo'qotish. Uni kamaytirish usullari: oldindan to'lov talab qilish, kelish kuniga 1 kun qolganida tasdiqlash so'rash va no-show jarimasi siyosati.",
            "Overbooking — mavjud xonalar sonidan ko'proq bronlashni qabul qilish strategiyasi. Bu no-show va bekor qilishlarni hisobga olib, bandlik darajasini maksimallash uchun ishlatiladi. Ammo hamma mehmonlar kelsa, muammo yuzaga keladi.",
            "Overbooking boshqaruvi: PMS tarixiy ma'lumotlar asosida qancha overbooking xavfsiz ekanini hisoblaydi. Agar barcha mehmonlar kelsa — mehmon boshqa mehmonxonaga yo'naltiriladi (walking guest), transport va birinchi kecha to'lovi mehmonxona hisobidan qoplanadi.",
        ]),
    ],
})

TOPICS.append({
    "num": 28,
    "title": "Turizm, mehmonxona xizmatlarini ijtimoiy tarmoq, elektron pochta, faks, telefon, internet orqali bron qilish",
    "reja": [
        "Turli kanallar orqali bronlash usullari.",
        "Ijtimoiy tarmoqlar va messenjerlar orqali bronlash.",
        "Email va telefon orqali bronlash jarayoni.",
        "Ko'p kanalli bronlash strategiyasi (omnichannel).",
    ],
    "sections": [
        ("Turli kanallar orqali bronlash usullari", [
            "Mehmonxona bronlash kanallari: to'g'ridan-to'g'ri (veb-sayt, telefon, email, walk-in), OTA (Booking.com, Expedia), GDS (turoperatorlar uchun), ijtimoiy tarmoqlar (Instagram, Facebook) va messenjerlar (Telegram, WhatsApp). Har bir kanal o'z auditoriyasiga ega.",
            "Kanal samaradorligini baholash: qaysi kanal ko'proq bronlash olib keladi, qaysi birining komissiyasi past va qaysi biri sifatli (ko'p tunadigan, ko'p sarflaydigan) mehmonlarni jalb qiladi. Bu ma'lumotlar marketing byudjetini taqsimlashda yordam beradi.",
            "Channel Manager — barcha kanallarni bir joydan boshqarish tizimi. Narx yoki mavjudlik o'zgarganda barcha platformalarda bir vaqtda yangilanadi. Bu ikkilangan bronlash (double booking) xavfini bartaraf etadi.",
        ]),
        ("Ijtimoiy tarmoqlar va messenjerlar orqali bronlash", [
            "Instagram orqali bronlash: Direct xabar orqali so'rov, Instagram Shopping orqali xizmat tanlash va havolaga o'tib bronlash. Chatbot integratsiyasi avtomatik javob berish va bronlash jarayonini boshqarish imkonini beradi.",
            "Telegram orqali bronlash: Telegram bot yaratish — mijoz bot bilan suhbatlashib sanalarni tanlaydi, xona turini ko'radi va bronlashni tasdiqlaydi. O'zbekistonda Telegram eng ommaviy messenger bo'lgani uchun bu kanal juda samarali.",
            "Facebook/WhatsApp Business: avtomatik javoblar sozlash, katalog yaratish (xonalar yoki xizmatlar), va to'g'ridan-to'g'ri xabar orqali bronlash. WhatsApp Business API katta korxonalar uchun avtomatlashtirilgan xabar almashish imkonini beradi.",
        ]),
        ("Email va telefon orqali bronlash jarayoni", [
            "Email orqali bronlash: mijoz so'rov yuboradi (sanalar, mehmonlar soni, xona turi), mehmonxona mavjudlik va narxni tekshirib javob beradi, mijoz tasdiqlaydi va to'lovni amalga oshiradi. Bu jarayon 2-24 soat davom etishi mumkin.",
            "Telefon orqali bronlash: resepsionist mijoz bilan gaplashib ma'lumotlarni oladi, PMS da mavjudlikni tekshiradi, bronlashni kiritadi va tasdiqlash raqamini beradi. Bu shaxsiy muloqot imkonini beradi, lekin ko'proq vaqt talab qiladi.",
            "Faks — eskirgan, lekin ba'zi korporativ mijozlar va Yaponiya kabi mamlakatlarda hali ishlatiladigan usul. Bronlash formasi to'ldiriladi va faks orqali yuboriladi. Zamonaviy mehmonxonalarda online fax xizmatlari (eFax) orqali qog'ozsiz ishlash mumkin.",
        ]),
        ("Ko'p kanalli bronlash strategiyasi (omnichannel)", [
            "Omnichannel yondashuv — barcha kanallarda yagona va izchil tajriba yaratish. Mijoz qaysi kanaldan murojaat qilmasin (sayt, telefon, messenger, ijtimoiy tarmoq), bir xil ma'lumot, narx va xizmat sifatini olishi kerak.",
            "Strategiya qoidalari: barcha kanallarda narx pariteti (bir xil narx), tez javob berish (24 soat ichida), yagona mijoz bazasi (qaysi kanaldan kelganini bilish) va professional muloqot standarti (barcha kanallarda yagona ton).",
            "Texnologik yechim: CRM tizimi barcha kanallardan kelgan so'rovlarni bitta joyga yig'adi, har bir mijozning aloqa tarixini saqlaydi va keyingi murojaat uchun kontekst beradi. Bu mijoz tajribasini sezilarli yaxshilaydi.",
        ]),
    ],
})

TOPICS.append({
    "num": 29,
    "title": "Axborot texnologiyalari asosida turizm, mehmonxona va restoran xizmatlarini rivojlantirish, hujjatlar bilan ishlash",
    "reja": [
        "AT asosida xizmatlarni rivojlantirish strategiyalari.",
        "Raqamli transformatsiya bosqichlari.",
        "Hujjatlar bilan ishlashni raqamlashtirish.",
        "Kelajak tendensiyalari va innovatsiyalar.",
    ],
    "sections": [
        ("AT asosida xizmatlarni rivojlantirish strategiyalari", [
            "Raqamli strategiya — bu axborot texnologiyalari yordamida korxona maqsadlariga erishish rejasi. Turizm korxonasi uchun bu: onlayn mavjudlikni kuchaytirish, mijozlar tajribasini yaxshilash, operatsion samaradorlikni oshirish va yangi daromad manbalari yaratish.",
            "Personalizatsiya strategiyasi: mijozlar ma'lumotlari tahlili asosida har bir mehmon uchun individual takliflar, xizmatlar va muloqot. Masalan: qaytib kelgan mehmonning afzalliklarini eslab, mos xona va xizmatlarni taklif qilish.",
            "Raqamli marketing strategiyasi: SEO (qidiruv tizimlarida ko'rinish), kontekstli reklama, ijtimoiy tarmoqlar marketingi, email marketing va kontent marketing. Bu kanallar integratsiyalashgan holda ishlashi eng yaxshi natija beradi.",
        ]),
        ("Raqamli transformatsiya bosqichlari", [
            "1-bosqich — Raqamlashtirish (digitization): qog'oz jarayonlarni raqamli formatga o'tkazish. Masalan: qog'oz bronlash daftarini PMS ga almashtirish, qo'lda yoziladigan hisobotlarni Excel ga o'tkazish.",
            "2-bosqich — Raqamli optimallashtirish (digitalization): mavjud jarayonlarni raqamli vositalar bilan yaxshilash. Masalan: onlayn bronlash tizimi, avtomatlashtirilgan email xabarnomalar, elektron hisob-kitob.",
            "3-bosqich — Raqamli transformatsiya: biznes modelini tubdan o'zgartirish. Masalan: to'liq onlayn turizm agentligi, chatbot asosidagi mijozlarga xizmat ko'rsatish, sun'iy intellekt yordamida personalizatsiya. Bu eng yuqori daraja.",
        ]),
        ("Hujjatlar bilan ishlashni raqamlashtirish", [
            "Turizm korxonasidagi hujjatlar: shartnomalar (mijozlar, hamkorlar, yetkazib beruvchilar bilan), moliyaviy hujjatlar (hisob-fakturalar, aktlar), ichki hujjatlar (buyruqlar, yo'riqnomalar) va turistik hujjatlar (marshrut varaqlari, voucherlar).",
            "Raqamlashtirish usullari: hujjat shablonlari yaratish (Word/Google Docs), avtomatik to'ldirish (CRM dan ma'lumot olish), elektron imzo (tasdiqlash uchun), bulutli saqlash (Google Drive/OneDrive) va raqamli arxiv (tezkor qidirish imkoniyati).",
            "Hujjat aylanish tizimi afzalliklari: vaqtni tejash (hujjatni tez topish), xavfsizlik (kirish huquqlari boshqaruvi), shaffoflik (kim qachon nima qilganini ko'rish), qog'oz tejash va masofadan ishlash imkoniyati.",
        ]),
        ("Kelajak tendensiyalari va innovatsiyalar", [
            "Sun'iy intellekt (AI): chatbotlar (24/7 mijoz xizmati), personalizatsiya (individual takliflar), dinamik narxlash (talab asosida avtomatik narx o'zgartirish) va bashoratli tahlil (talabni oldindan aniqlash).",
            "IoT (Internet of Things): aqlli xona (harorat, yorug'lik, parda avtomatik boshqariladi), energiya tejash sensorlari, prediktiv texnik xizmat (buzilishni oldindan aniqlash) va mehmon joylashuvini kuzatish (xizmat ko'rsatish uchun).",
            "Blockchain va NFT: xavfsiz to'lovlar, sodiqlik dasturi tokenlari, soxta sharhlarni oldini olish va aqlli shartnomalar (smart contracts — shartlar bajarilganda avtomatik to'lov). Bu texnologiyalar turizm sohasida sinov bosqichida.",
        ]),
    ],
})

TOPICS.append({
    "num": 30,
    "title": "Turistik ofis hujjatlarini Word, Excel dasturlarida rasmiylashtirish",
    "reja": [
        "Turistik shartnomalarni Word da rasmiylashtirish.",
        "Turistik marshrut va voucherlarni tayyorlash.",
        "Excel da moliyaviy hujjatlar va hisobotlar.",
        "Hujjat shablonlari yaratish va avtomatlashtirish.",
    ],
    "sections": [
        ("Turistik shartnomalarni Word da rasmiylashtirish", [
            "Turistik xizmat ko'rsatish shartnomasi — turoperator/turagent va mijoz o'rtasidagi huquqiy hujjat. Shartnoma tarkibi: tomonlar ma'lumotlari, xizmat tavsifi (marshrut, muddatlar, xizmatlar ro'yxati), narx va to'lov shartlari, tomonlar majburiyatlari, bekor qilish shartlari va imzolar.",
            "Word da shartnoma rasmiylashtirish: A4 format, Times New Roman 12-14pt, 1-1.5 interval, raqamlangan bandlar, jadvallar (xizmatlar va narxlar uchun), joylash joy belgilangan imzo va muhr joyi. Sarlavha korxona logotipi bilan bezatilishi mumkin.",
            "Shablon yaratish: bir marta yaxshi rasmiyiashtirilgan shartnoma shablonini saqlash va har safar faqat mijoz ma'lumotlarini o'zgartirish. Word dagi form fields yoki mail merge funksiyasi bu jarayonni avtomatlashtiradi.",
        ]),
        ("Turistik marshrut va voucherlarni tayyorlash", [
            "Marshrut varag'i (itinerary) — turistik sayohatning batafsil dasturi. Tarkibi: kun bo'yicha reja, transport ma'lumotlari, turar joy manzillari, ekskursiya vaqtlari, ovqatlanish joylari va aloqa ma'lumotlari. U Word da jadval ko'rinishida tuziladi.",
            "Voucher — xizmat ko'rsatuvchiga (mehmonxona, restoran, transport) taqdim etiladigan tasdiqlovchi hujjat. Unda: korxona nomi, mehmon ismi, sanalar, xizmat turi, to'lov holati va turoperator aloqa ma'lumotlari ko'rsatiladi.",
            "Word da voucher tayyorlash: korxona logotipi va brend ranglari, jadval ichida ma'lumotlar, QR-kod (elektron tasdiqlash uchun). Voucher A5 o'lchamda (yarim varaq) ham tayyorlanishi mumkin. Elektron voucher PDF formatida email orqali yuboriladi.",
        ]),
        ("Excel da moliyaviy hujjatlar va hisobotlar", [
            "Hisob-faktura (invoice) yaratish: Excel da mijozga yuboriladigan to'lov hujjati. Tarkibi: korxona rekvizitlari, mijoz ma'lumotlari, xizmatlar ro'yxati (nomi, soni, narxi), jami summa, QQS va to'lov rekvizitlari. Formulalar jami summani avtomatik hisoblaydi.",
            "Oylik hisobot: sotish hajmi (turlar soni va summasi), xarajatlar (transport, turar joy, marketing), foyda, mijozlar statistikasi va kanal bo'yicha taqsimot. Diagrammalar vizual ko'rinish beradi. Pivot Table yig'ma tahlil uchun ishlatiladi.",
            "Byudjet rejasi: yillik daromad va xarajatlar rejasi, oylar bo'yicha taqsimot, reja va fakt taqqoslash. Shartli formatlash (Conditional Formatting) rejadan og'ishlarni rangli ajratib ko'rsatadi. Bu moliyaviy boshqaruvning asosiy vositasi.",
        ]),
        ("Hujjat shablonlari yaratish va avtomatlashtirish", [
            "Shablon (template) — har safar qayta yozmasdan, faqat o'zgaruvchan ma'lumotlarni to'ldirib ishlatiladigan tayyor hujjat. Turizm ofisi uchun zarur shablonlar: shartnoma, voucher, marshrut varag'i, hisob-faktura, tijorat taklifi va xat.",
            "Word shablonlari: .dotx formatida saqlanadi, ochilganda yangi hujjat yaratadi. Form fields (matn maydoni, sana tanlash, ro'yxatdan tanlash) kiritish mumkin. Quick Parts (tez kiritish bloklari) — ko'p ishlatiladigan matn bo'laklarini saqlash.",
            "Excel shablonlari: formulalar, formatlash va tuzilma tayyor, faqat raqamlar kiritiladi. Makroslar (VBA) — takroriy amallarni avtomatlashtirish. Masalan: bir tugma bilan oylik hisobotni generatsiya qilish yoki hisob-fakturani PDF ga aylantirish.",
        ]),
    ],
})

# ===================== ISHGA TUSHIRISH =====================
if __name__ == "__main__":
    write_docx("/projects/sandbox/konspekt-30-mavzu/Turizm_Konspekt_30_mavzu.docx")

# -*- coding: utf-8 -*-
"""
Konspekt generatori.
python-docx mavjud emasligi sababli .docx fayl (Office Open XML)
standart kutubxonalar (zipfile) yordamida qo'lda quriladi.
A4, Times New Roman 14pt, 1.5 interval, har bir mavzu yangi sahifadan.
"""
import zipfile
import datetime

# ------------------------------------------------------------------
# Mavzular ro'yxati. Har bir element:
#   {"num": int, "title": str, "reja": [4 ta string],
#    "sections": [(sarlavha, [paragraflar]) ... 4 ta]}
# ------------------------------------------------------------------
TOPICS = []


# ------------------------------------------------------------------
# XML yordamchi funksiyalari
# ------------------------------------------------------------------
def esc(text):
    """XML uchun maxsus belgilarni ekranlash."""
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;"))


def run(text, bold=False, italic=False, size=28, font="Times New Roman"):
    """Bitta matn 'run' (w:r) hosil qiladi. size = half-points (28 = 14pt)."""
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
    """
    Bitta paragraf (w:p) hosil qiladi.
    align: both|left|center|right
    indent_first: birinchi qator chekinishi (twips), 709 ~ 1.25 sm
    line: 360 = 1.5 interval (240 = 1.0)
    """
    ppr = '<w:pPr>'
    if keep_next:
        ppr += '<w:keepNext/>'
    ppr += '<w:spacing w:after="%d" w:line="%d" w:lineRule="auto"/>' % (space_after, line)
    if indent_first:
        ppr += '<w:ind w:firstLine="%d"/>' % indent_first
    ppr += '<w:jc w:val="%s"/>' % align
    ppr += '</w:pPr>'
    return '<w:p>%s%s</w:p>' % (ppr, run(text, bold=bold, italic=italic, size=size))


def para_runs(runs_xml, align="both", indent_first=0, space_after=120, line=360):
    """Tayyor run-lardan paragraf yasash (aralash formatlash uchun)."""
    ppr = '<w:pPr>'
    ppr += '<w:spacing w:after="%d" w:line="%d" w:lineRule="auto"/>' % (space_after, line)
    if indent_first:
        ppr += '<w:ind w:firstLine="%d"/>' % indent_first
    ppr += '<w:jc w:val="%s"/>' % align
    ppr += '</w:pPr>'
    return '<w:p>%s%s</w:p>' % (ppr, runs_xml)


def page_break():
    """Sahifa uzilishi (yangi mavzu uchun)."""
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def empty_para():
    return '<w:p><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="auto"/></w:pPr></w:p>'


# ------------------------------------------------------------------
# Bitta mavzuni XML ga aylantirish
# ------------------------------------------------------------------
def build_topic(t, is_first):
    parts = []
    if not is_first:
        parts.append(page_break())

    # Sarlavha
    parts.append(para("%d-MAVZU. %s" % (t["num"], t["title"].upper()),
                      bold=True, size=30, align="center",
                      indent_first=0, space_after=180, keep_next=True))

    # Reja
    parts.append(para("Reja:", bold=True, size=28, align="left",
                      indent_first=0, space_after=60, keep_next=True))
    for i, r in enumerate(t["reja"], 1):
        parts.append(para("%d. %s" % (i, r), size=28, align="left",
                          indent_first=360, space_after=40))
    parts.append(empty_para())

    # Bo'limlar
    for idx, (head, paras) in enumerate(t["sections"], 1):
        parts.append(para("%d. %s" % (idx, head), bold=True, size=28,
                          align="left", indent_first=0, space_after=80,
                          keep_next=True))
        for p in paras:
            parts.append(para(p, size=28, align="both",
                             indent_first=709, space_after=120))
    return "".join(parts)


# ------------------------------------------------------------------
# Hujjatni yig'ish
# ------------------------------------------------------------------
def build_document():
    body = []
    for i, t in enumerate(TOPICS):
        body.append(build_topic(t, is_first=(i == 0)))

    # A4 sahifa: 11906 x 16838 twips. Margins: left 3sm, right 1.5sm, top/bottom 2sm
    sect = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
        '<w:pgNumType w:start="1"/>'
        '</w:sectPr>'
    )

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:body>' + "".join(body) + sect + '</w:body></w:document>'
    )
    return document


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    '</Types>'
)

RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '</Relationships>'
)

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>'
)

STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
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
    '</w:styles>'
)


def core_props():
    now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Texnik tizimlarda axborot texnologiyalari - Konspekt</dc:title>'
        '<dc:creator>Konspekt generatori</dc:creator>'
        '<cp:lastModifiedBy>Konspekt generatori</cp:lastModifiedBy>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
        '</cp:coreProperties>' % (now, now)
    )


APP_PROPS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    '<Application>Konspekt generatori</Application></Properties>'
)


def write_docx(path="Konspekt.docx"):
    document = build_document()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("docProps/core.xml", core_props())
        z.writestr("docProps/app.xml", APP_PROPS)
    # Statistika
    words = 0
    for t in TOPICS:
        for _, paras in t["sections"]:
            for p in paras:
                words += len(p.split())
    print("Yaratildi:", path)
    print("Mavzular soni:", len(TOPICS))
    print("Taxminiy so'zlar (faqat matn tanasi):", words)


# main() oxirida (mavzular qo'shilgandan keyin) chaqiriladi.



# ==================================================================
# MAVZULAR KONTENTI
# ==================================================================

TOPICS.append({
    "num": 1,
    "title": "Texnik tizimlarda axborot texnologiyalari faniga kirish",
    "reja": [
        "Axborot texnologiyalari tushunchasi va uning rivojlanish tarixi.",
        "Axborot, ma’lumot va bilim tushunchalari, ularning xususiyatlari.",
        "Texnik tizimlarda axborot texnologiyalarining tutgan o‘rni va ahamiyati.",
        "Fanning maqsadi, vazifalari hamda boshqa fanlar bilan bog‘liqligi.",
    ],
    "sections": [
        ("Axborot texnologiyalari tushunchasi va uning rivojlanish tarixi", [
            "Axborot texnologiyalari (AT) — bu axborotni yig‘ish, saqlash, qayta ishlash, uzatish va undan foydalanish jarayonlarini amalga oshiruvchi usullar, vositalar hamda dasturiy-texnik ta’minotlar majmuasidir. Zamonaviy sharoitda axborot texnologiyalari nafaqat hisoblash texnikasi, balki aloqa vositalari, dasturiy mahsulotlar va tarmoq tizimlarini ham qamrab oladi. Ularning asosiy vazifasi insonga kerakli axborotni o‘z vaqtida, aniq va qulay shaklda taqdim etishdir.",
            "Axborot texnologiyalarining rivojlanish tarixini bir necha bosqichga bo‘lish mumkin. Dastlabki bosqich qo‘lda yozish va og‘zaki uzatishga asoslangan bo‘lsa, keyinchalik bosmaxona ixtirosi axborotni ommaviy tarqatish imkonini berdi. XX asrning o‘rtalarida elektron hisoblash mashinalarining (EHM) paydo bo‘lishi axborotni avtomatik qayta ishlash davrini boshlab berdi.",
            "Bugungi kunda axborot texnologiyalari mikroprotsessorlar, shaxsiy kompyuterlar, Internet tarmog‘i va bulutli hisoblash texnologiyalari asosida jadal rivojlanmoqda. Har bir yangi bosqich axborotni qayta ishlash tezligini oshirib, uning hajmini kengaytirib bormoqda. Shu sababli zamonaviy jamiyat ko‘pincha “axborot jamiyati” deb ataladi.",
        ]),
        ("Axborot, ma’lumot va bilim tushunchalari, ularning xususiyatlari", [
            "Ma’lumot (data) — bu real dunyo hodisalari, jismlari yoki jarayonlari haqidagi qayd etilgan, lekin hali talqin qilinmagan belgilar, raqamlar yoki faktlardir. Ma’lumotlarning o‘zi to‘g‘ridan-to‘g‘ri qaror qabul qilish uchun yetarli bo‘lmasligi mumkin, chunki ular ma’no kasb etishi uchun qayta ishlanishi lozim.",
            "Axborot (information) — bu inson uchun ma’no kasb etadigan, talqin qilingan va muayyan kontekstda foydali bo‘lgan ma’lumotdir. Boshqacha aytganda, ma’lumotlar qayta ishlangach, tartibga solinib, izohlangach axborotga aylanadi. Axborotning muhim xususiyatlari: aniqlik, to‘liqlik, o‘z vaqtidaligi, ishonchliligi va foydaliligidir.",
            "Bilim (knowledge) — bu axborot asosida shakllangan, tajriba va tahlil natijasida umumlashtirilgan, yangi vaziyatlarda qo‘llash mumkin bo‘lgan tushunchalar majmuasidir. Ma’lumot, axborot va bilim o‘zaro bog‘liq pog‘onalarni tashkil etadi: ma’lumotdan axborot, axborotdan esa bilim hosil bo‘ladi. Texnik tizimlarda bu uch tushunchani to‘g‘ri farqlash samarali qarorlar qabul qilish uchun asos bo‘ladi.",
        ]),
        ("Texnik tizimlarda axborot texnologiyalarining tutgan o‘rni va ahamiyati", [
            "Texnik tizim deganda biror maqsadga erishish uchun o‘zaro bog‘langan elementlar (qurilmalar, mexanizmlar, dasturiy vositalar) majmuasi tushuniladi. Bunday tizimlarda axborot texnologiyalari boshqaruv, monitoring, hisoblash va loyihalash jarayonlarini avtomatlashtirish uchun keng qo‘llaniladi.",
            "Masalan, ishlab chiqarish korxonalarida texnologik jarayonlarni boshqarishda avtomatlashtirilgan tizimlar, transport sohasida navigatsiya va dispetcherlik tizimlari, energetikada esa hisob-kitob va nazorat tizimlari axborot texnologiyalariga asoslanadi. Bu tizimlar inson xatosini kamaytiradi, ish unumdorligini oshiradi va resurslarni tejaydi.",
            "Axborot texnologiyalarining texnik tizimlardagi ahamiyati nafaqat jarayonlarni tezlashtirishda, balki katta hajmdagi ma’lumotlarni tahlil qilib, ishonchli qarorlar qabul qilishga ko‘maklashishda ham namoyon bo‘ladi. Shuning uchun zamonaviy muhandis axborot texnologiyalaridan unumli foydalana olishi zarur.",
        ]),
        ("Fanning maqsadi, vazifalari hamda boshqa fanlar bilan bog‘liqligi", [
            "“Texnik tizimlarda axborot texnologiyalari” fanining asosiy maqsadi — talabalarda axborotni qayta ishlashning zamonaviy usullari, dasturiy vositalari va texnik vositalari haqida nazariy bilim va amaliy ko‘nikmalarni shakllantirishdir. Fan muhandislik faoliyatida axborot texnologiyalaridan samarali foydalanishni o‘rgatadi.",
            "Fanning vazifalari sirasiga matn va jadval muharrirlari bilan ishlash, avtomatlashtirilgan loyihalash tizimlarini o‘zlashtirish, matematik hisoblash dasturlaridan foydalanish va dasturlash asoslarini egallash kiradi. Bu ko‘nikmalar bo‘lajak mutaxassisning kasbiy salohiyatini oshiradi.",
            "Fan informatika, matematika, fizika, dasturlash va muhandislik fanlari bilan chambarchas bog‘liq. U bir tomondan matematik va informatik bilimlarga tayanadi, ikkinchi tomondan esa ularni texnik masalalarni yechishda amaliy qo‘llashga yo‘naltiradi. Shu bois fanni o‘zlashtirish yaxlit tizimli yondashuvni talab qiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 2,
    "title": "Axborot kommunikatsion texnologiyalarining komponentlari. Kompyuter tizimlari",
    "reja": [
        "Axborot kommunikatsion texnologiyalari (AKT) tushunchasi va tarkibi.",
        "Kompyuterning apparat ta’minoti (hardware) komponentlari.",
        "Dasturiy ta’minot (software) va uning turlari.",
        "Kompyuter tizimlarining tasnifi va ishlash prinsipi.",
    ],
    "sections": [
        ("Axborot kommunikatsion texnologiyalari (AKT) tushunchasi va tarkibi", [
            "Axborot kommunikatsion texnologiyalari (AKT) — axborotni yaratish, saqlash, qayta ishlash va uzatishni ta’minlovchi texnik vositalar, dasturlar hamda aloqa tizimlarining yaxlit majmuasidir. AKT axborot texnologiyalari bilan telekommunikatsiya vositalarining birlashuvi natijasida shakllangan.",
            "AKT tarkibiga apparat vositalari (kompyuterlar, serverlar, tarmoq qurilmalari), dasturiy ta’minot, ma’lumotlar bazasi, aloqa kanallari va foydalanuvchilar kiradi. Bu komponentlarning o‘zaro muvofiq ishlashi axborotni samarali boshqarish imkonini beradi.",
            "Zamonaviy jamiyatda AKT ta’lim, tibbiyot, sanoat, boshqaruv va boshqa sohalarda keng qo‘llaniladi. Ular masofadan ta’lim olish, elektron hujjat almashinuvi va onlayn xizmatlarning rivojlanishiga asos bo‘lmoqda.",
        ]),
        ("Kompyuterning apparat ta’minoti (hardware) komponentlari", [
            "Apparat ta’minoti — kompyuterning qo‘l bilan ushlash mumkin bo‘lgan fizik qismlari majmuasidir. Uning markaziy qismi protsessor (CPU) bo‘lib, u barcha hisoblash va boshqaruv amallarini bajaradi. Protsessorning tezligi va yadrolar soni kompyuter unumdorligini belgilaydi.",
            "Operativ xotira (RAM) hozirda bajarilayotgan dasturlar va ma’lumotlarni vaqtincha saqlaydi; u tez ishlaydi, lekin energiya o‘chsa ma’lumotlar yo‘qoladi. Doimiy xotira qurilmalari — qattiq disk (HDD) va yarimo‘tkazgichli disk (SSD) — ma’lumotlarni uzoq muddat saqlash uchun ishlatiladi.",
            "Kiritish qurilmalari (klaviatura, sichqoncha, skaner) orqali foydalanuvchi ma’lumot kiritadi, chiqarish qurilmalari (monitor, printer, akustik tizim) esa natijalarni taqdim etadi. Bularning barchasini ona plata (motherboard) o‘zaro bog‘lab, yagona tizim sifatida ishlashini ta’minlaydi.",
        ]),
        ("Dasturiy ta’minot (software) va uning turlari", [
            "Dasturiy ta’minot — kompyuterga muayyan vazifalarni bajarishni buyuruvchi dasturlar va ular bilan bog‘liq hujjatlar majmuasidir. Dasturiy ta’minotsiz apparat qism o‘z-o‘zicha hech qanday amal bajara olmaydi.",
            "Dasturiy ta’minot uch asosiy turga bo‘linadi: tizimli, amaliy va instrumental. Tizimli dasturiy ta’minot (operatsion tizimlar, drayverlar) kompyuter resurslarini boshqaradi. Amaliy dasturiy ta’minot (matn muharrirlari, jadval protsessorlari, brauzerlar) foydalanuvchining aniq masalalarini yechishga xizmat qiladi.",
            "Instrumental dasturiy ta’minot esa yangi dasturlar yaratish uchun mo‘ljallangan vositalar — dasturlash tillari, kompilyatorlar va integratsiyalashgan ishlab chiqish muhitlarini o‘z ichiga oladi. Bu turlarning birgalikda ishlashi kompyuterning to‘liq funksional ishlashini ta’minlaydi.",
        ]),
        ("Kompyuter tizimlarining tasnifi va ishlash prinsipi", [
            "Kompyuter tizimlari turli mezonlarga ko‘ra tasniflanadi. Ko‘lami va quvvatiga ko‘ra ular superkompyuterlar, katta EHMlar (mainframe), serverlar, shaxsiy kompyuterlar va mobil qurilmalarga bo‘linadi. Har bir tur ma’lum maqsad va ish hajmiga moslashtirilgan.",
            "Aksariyat zamonaviy kompyuterlar Jon fon Neyman taklif etgan prinsipga asoslanadi: dastur va ma’lumotlar yagona xotirada saqlanadi, protsessor esa buyruqlarni ketma-ket o‘qib bajaradi. Bu prinsip kiritish, qayta ishlash, saqlash va chiqarish bosqichlarini o‘z ichiga oladi.",
            "Kompyuter tizimlari yakka holda yoki tarmoqqa birlashtirilgan holda ishlashi mumkin. Tarmoqqa ulangan tizimlar resurslarni birgalikda ishlatish va ma’lumot almashish imkonini beradi. Shu tariqa kompyuter tizimlari zamonaviy axborot infratuzilmasining asosini tashkil etadi.",
        ]),
    ],
})



TOPICS.append({
    "num": 3,
    "title": "Word dasturida hujjatlar ustida ishlash",
    "reja": [
        "Microsoft Word dasturi va uning interfeysi bilan tanishish.",
        "Matn kiritish, tahrirlash va formatlash amallari.",
        "Hujjatga jadval, rasm va boshqa obyektlarni qo‘shish.",
        "Hujjatni saqlash, chop etish va sahifa parametrlarini sozlash.",
    ],
    "sections": [
        ("Microsoft Word dasturi va uning interfeysi bilan tanishish", [
            "Microsoft Word — matnli hujjatlarni yaratish, tahrirlash va formatlashga mo‘ljallangan eng keng tarqalgan matn muharriridir. U Microsoft Office paketining tarkibiy qismi bo‘lib, hujjatlarni professional ko‘rinishda rasmiylashtirish imkonini beradi. Word yordamida arizalar, ma’ruzalar, kitoblar va boshqa hujjatlar tayyorlanadi.",
            "Word dasturi interfeysi tasma (Ribbon) deb ataluvchi yuqori paneldan iborat bo‘lib, unda “Bosh sahifa”, “Qo‘yish”, “Dizayn”, “Sahifa maketi” kabi yorliqlar joylashgan. Har bir yorliq tegishli buyruqlar guruhini o‘z ichiga oladi. Tez kirish paneli esa eng ko‘p ishlatiladigan buyruqlarni qo‘l ostida saqlaydi.",
            "Ish maydonining markazida hujjat varag‘i joylashadi, chetlarida o‘lchov chizg‘ichlari va aylantirish liniyalari mavjud. Pastki holat qatorida sahifa raqami, so‘zlar soni va ko‘rinish rejimlari aks etadi. Bu elementlar bilan tanishish samarali ishlash uchun zarur.",
        ]),
        ("Matn kiritish, tahrirlash va formatlash amallari", [
            "Hujjatga matn klaviatura yordamida kiritiladi. Kursor matn kiritiladigan joyni belgilaydi. Matnni belgilash (ajratish) sichqoncha yoki Shift tugmasi bilan amalga oshiriladi. Belgilangan matnni nusxalash (Ctrl+C), kesish (Ctrl+X) va joylashtirish (Ctrl+V) buyruqlari orqali ko‘chirish mumkin.",
            "Matnni formatlash — uning tashqi ko‘rinishini o‘zgartirishdir. Shrift turi, o‘lchami, rangi, qalin (Bold), qiya (Italic) va tagiga chizilgan (Underline) ko‘rinishlari “Bosh sahifa” yorlig‘ida sozlanadi. To‘g‘ri tanlangan shrift hujjatni o‘qishga qulay qiladi.",
            "Paragraf darajasida tekislash (chap, o‘ng, markaz, eniga), qatorlar orasidagi interval, abzas chekinishi va ro‘yxatlar (markerli yoki raqamli) sozlanadi. Uslublar (Styles) yordamida sarlavhalar va matnga yagona ko‘rinish berish mumkin, bu esa mundarija avtomatik tuzishni osonlashtiradi.",
        ]),
        ("Hujjatga jadval, rasm va boshqa obyektlarni qo‘shish", [
            "Word hujjatga turli obyektlarni qo‘shish imkonini beradi. Jadval qo‘yish uchun “Qo‘yish” yorlig‘idagi “Jadval” buyrug‘idan foydalaniladi; ustun va satrlar soni belgilanib, ma’lumotlar tartibli ko‘rinishda joylashtiriladi. Jadval kataklarini birlashtirish, ranglash va chegaralarini sozlash mumkin.",
            "Hujjatga rasm, shakl (figure), SmartArt diagrammalari va diagrammalar qo‘shish ham “Qo‘yish” yorlig‘i orqali bajariladi. Rasmni o‘lchamini o‘zgartirish, matn bilan o‘rab joylashtirish va effektlar berish hujjatni ko‘rgazmali qiladi.",
            "Bundan tashqari, hujjatga kolontitullar (yuqori va pastki sarlavhalar), sahifa raqamlari, izohlar va giperhavolalarni qo‘shish mumkin. Bu vositalar yirik va murakkab hujjatlarni tartibli rasmiylashtirishda muhim ahamiyatga ega.",
        ]),
        ("Hujjatni saqlash, chop etish va sahifa parametrlarini sozlash", [
            "Tayyorlangan hujjatni saqlash uchun “Fayl” menyusidagi “Saqlash” yoki “Boshqacha saqlash” buyruqlaridan foydalaniladi. Word hujjatlari odatda .docx formatida saqlanadi, biroq PDF, RTF kabi boshqa formatlarda ham saqlash mumkin. Ishni vaqti-vaqti bilan saqlab borish ma’lumot yo‘qolishining oldini oladi.",
            "Sahifa parametrlari “Sahifa maketi” (Layout) yorlig‘ida sozlanadi. Bu yerda qog‘oz o‘lchami (masalan, A4), sahifa yo‘nalishi (kitob yoki albom), maydonlar (margins) va ustunlar soni belgilanadi. To‘g‘ri sozlangan parametrlar hujjatning bosma ko‘rinishini belgilaydi.",
            "Hujjatni chop etishdan oldin “Fayl” menyusidagi “Chop etish” bo‘limida oldindan ko‘rish (preview) orqali natijani tekshirish tavsiya etiladi. Bu yerda printer, nusxalar soni va chop etiladigan sahifalar diapazoni tanlanadi. Shunday qilib, hujjat yakuniy ko‘rinishga keltiriladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 4,
    "title": "Texnik tizimlarda murakkab hujjatlar yaratish va qayta ishlash texnologiyasi",
    "reja": [
        "Murakkab hujjat tushunchasi va uning tuzilmasi.",
        "Uslublar, mundarija va havolalar bilan ishlash.",
        "Formulalar, diagrammalar va texnik elementlarni joylashtirish.",
        "Hujjatlarni birgalikda tahrirlash va himoyalash texnologiyalari.",
    ],
    "sections": [
        ("Murakkab hujjat tushunchasi va uning tuzilmasi", [
            "Murakkab hujjat — bu nafaqat oddiy matndan, balki turli xil elementlar: jadvallar, rasmlar, formulalar, diagrammalar, mundarija va havolalardan tashkil topgan yirik hujjatdir. Texnik sohalarda bunday hujjatlarga loyiha hujjatlari, hisobotlar, qo‘llanmalar va ilmiy ishlar misol bo‘ladi.",
            "Murakkab hujjatning tuzilmasi mantiqiy darajalardan iborat bo‘ladi: bosh sarlavha, bo‘limlar, kichik bo‘limlar va paragraflar. Bu ierarxiya hujjatni o‘qish va undan foydalanishni osonlashtiradi. To‘g‘ri tuzilmaviy yondashuv hujjatni boshqarishni soddalashtiradi.",
            "Murakkab hujjatlarni yaratishda standart talablariga (masalan, davlat standartlari yoki korxona ichki me’yorlari) rioya qilish muhim. Bu hujjatning rasmiy va professional ko‘rinishini ta’minlaydi hamda uni boshqa mutaxassislar tomonidan tushunilishini osonlashtiradi.",
        ]),
        ("Uslublar, mundarija va havolalar bilan ishlash", [
            "Uslublar (Styles) — matn va sarlavhalarga oldindan belgilangan formatlash to‘plamini bir buyruq bilan qo‘llash imkonini beradi. Sarlavhalarga “Sarlavha 1”, “Sarlavha 2” uslublarini berish orqali hujjat tuzilmasi belgilanadi. Bu nafaqat ko‘rinishni birxillashtiradi, balki avtomatik mundarija tuzishga asos yaratadi.",
            "Mundarija (Table of Contents) — hujjat bo‘limlari va ularning sahifa raqamlari ro‘yxatidir. Word uni uslublar asosida avtomatik yaratadi va hujjat o‘zgarganda yangilash mumkin. Bu yirik hujjatlarda kerakli bo‘limni tez topishni ta’minlaydi.",
            "Havolalar (cross-references) va giperhavolalar hujjat ichidagi yoki tashqi manbalarga bog‘lanishni yaratadi. Adabiyotlar ro‘yxati, izohlar va iqtiboslar bilan ishlash texnik hujjatlarning ilmiy aniqligini oshiradi.",
        ]),
        ("Formulalar, diagrammalar va texnik elementlarni joylashtirish", [
            "Texnik hujjatlarda matematik va fizik formulalar muhim o‘rin tutadi. Word ichidagi “Tenglama” (Equation) muharriri yordamida murakkab formulalarni — kasrlar, ildizlar, integrallar, indekslar va yuqori darajalarni to‘g‘ri ko‘rinishda yozish mumkin.",
            "Diagrammalar (Chart) ma’lumotlarni ko‘rgazmali tarzda taqdim etadi. Ustunli, chiziqli, doiraviy va boshqa diagramma turlari tahlil natijalarini vizual ifodalaydi. Diagrammalar ko‘pincha Excel ma’lumotlari bilan bog‘liq holda yaratiladi.",
            "Bundan tashqari, texnik hujjatlarga chizmalar, sxema va belgilarni joylashtirish kerak bo‘ladi. SmartArt va shakllar yordamida jarayon sxemalari va tuzilmaviy diagrammalar tuziladi. Bu elementlar texnik g‘oyalarni aniq va tushunarli ifodalashga xizmat qiladi.",
        ]),
        ("Hujjatlarni birgalikda tahrirlash va himoyalash texnologiyalari", [
            "Zamonaviy hujjat ishlash texnologiyalari bir nechta mualliflarning bitta hujjat ustida birgalikda ishlashiga imkon beradi. Bulutli xizmatlar (masalan, OneDrive) orqali hujjatni real vaqtda birga tahrirlash mumkin. “O‘zgarishlarni kuzatish” (Track Changes) funksiyasi har bir muharrir kiritgan tuzatishlarni qayd etadi.",
            "Sharhlar (Comments) yordamida hujjat matniga izoh va takliflar qo‘shiladi, bu jamoaviy ishni muvofiqlashtiradi. O‘zgarishlarni qabul qilish yoki rad etish orqali yakuniy variant shakllantiriladi.",
            "Hujjatni himoyalash maqsadida parol qo‘yish, tahrirlashni cheklash yoki faqat o‘qish rejimini yoqish mumkin. Bu maxfiy yoki muhim texnik hujjatlarning ruxsatsiz o‘zgartirilishi va tarqalishining oldini oladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 5,
    "title": "Zamonaviy avtomatlashtirilgan loyihalash tizimlari",
    "reja": [
        "Avtomatlashtirilgan loyihalash tizimi (CAD) tushunchasi.",
        "CAD tizimlarining tarkibi va asosiy imkoniyatlari.",
        "Mashhur CAD dasturlari va ularning xususiyatlari.",
        "CAD tizimlaridan foydalanishning afzalliklari va istiqbollari.",
    ],
    "sections": [
        ("Avtomatlashtirilgan loyihalash tizimi (CAD) tushunchasi", [
            "Avtomatlashtirilgan loyihalash tizimi (ingl. CAD — Computer-Aided Design) — bu muhandislik obyektlarini loyihalash jarayonini kompyuter yordamida avtomatlashtiruvchi dasturiy va texnik vositalar majmuasidir. CAD tizimlari chizmalar yaratish, hisob-kitoblar bajarish va loyiha hujjatlarini tayyorlashni soddalashtiradi.",
            "An’anaviy qo‘lda loyihalashda ko‘p vaqt va mehnat sarflanardi, xatolik ehtimoli yuqori edi. CAD tizimlari esa loyihani tez, aniq va o‘zgartirishga qulay shaklda yaratishga imkon beradi. Loyihaning istalgan qismini osongina tahrirlash mumkin.",
            "CAD tizimlari ikki o‘lchovli (2D) chizmalar bilan birga uch o‘lchovli (3D) modellarni yaratishni ham qo‘llab-quvvatlaydi. 3D modellashtirish obyektni har tomondan ko‘rish, tahlil qilish va virtual sinovdan o‘tkazish imkonini beradi.",
        ]),
        ("CAD tizimlarining tarkibi va asosiy imkoniyatlari", [
            "CAD tizimi texnik ta’minot (kompyuter, grafik plansheti, plotter), dasturiy ta’minot va ma’lumotlar bazasidan iborat. Dasturiy ta’minot geometrik modellashtirish, chizmachilik va hisoblash modullarini o‘z ichiga oladi.",
            "Tizimning asosiy imkoniyatlari qatoriga aniq o‘lchamli chizmalar yaratish, parametrik modellashtirish, qatlamlar (layers) bilan ishlash, o‘lchamlarni avtomatik qo‘yish va kutubxonalardan tayyor elementlardan foydalanish kiradi. Parametrik modellashtirishda obyekt o‘lchamini o‘zgartirish butun model bo‘ylab avtomatik aks etadi.",
            "CAD tizimlari ko‘pincha muhandislik tahlili (CAE) va ishlab chiqarishni tayyorlash (CAM) tizimlari bilan integratsiyalashadi. Bu loyihalashdan ishlab chiqarishgacha bo‘lgan jarayonni yagona raqamli muhitda boshqarish imkonini beradi.",
        ]),
        ("Mashhur CAD dasturlari va ularning xususiyatlari", [
            "Dunyoda eng keng tarqalgan CAD dasturlaridan biri AutoCAD bo‘lib, u 2D va 3D chizmachilikda universalligi bilan ajralib turadi. U qurilish, mashinasozlik va elektrotexnika sohalarida keng qo‘llaniladi.",
            "SolidWorks va Autodesk Inventor dasturlari uch o‘lchovli parametrik modellashtirishga ixtisoslashgan bo‘lib, mashinasozlik detallari va yig‘ma birikmalarni loyihalashda samaralidir. KOMPAS-3D esa MDH mamlakatlarida commondir va standart talablariga moslashgan.",
            "Qurilish sohasida Revit kabi BIM (Building Information Modeling) tizimlari, elektronika sohasida esa Altium Designer va boshqa maxsus dasturlar qo‘llaniladi. Har bir dastur muayyan sohaning ehtiyojlariga moslashtirilgan vositalarga ega.",
        ]),
        ("CAD tizimlaridan foydalanishning afzalliklari va istiqbollari", [
            "CAD tizimlaridan foydalanishning asosiy afzalligi — loyihalash tezligi va aniqligining ortishidir. Loyihalardagi xatolar ishlab chiqarishdan oldin aniqlanadi, bu esa vaqt va moddiy resurslarni tejaydi. Tayyor elementlar kutubxonasi takroriy ishlarni kamaytiradi.",
            "CAD tizimlari loyiha hujjatlarini standartlarga muvofiq avtomatik shakllantiradi, loyihani osongina o‘zgartirish va turli variantlarni taqqoslash imkonini beradi. Raqamli model boshqa mutaxassislar bilan tez almashinish va jamoaviy ishlashni osonlashtiradi.",
            "Kelajakda CAD tizimlari sun’iy intellekt, bulutli hisoblash va virtual hamda kengaytirilgan reallik (VR/AR) texnologiyalari bilan tobora ko‘proq integratsiyalashmoqda. Bu loyihalashni yanada avtomatlashtirib, generativ dizayn kabi yangi imkoniyatlarni ochmoqda.",
        ]),
    ],
})



TOPICS.append({
    "num": 6,
    "title": "Avtomatlashtirilgan loyihalash tizimlarining texnik sohalarda qo'llanilishi",
    "reja": [
        "CAD tizimlarining mashinasozlikda qo'llanilishi.",
        "Qurilish va arxitektura sohasida CAD tizimlari.",
        "Elektrotexnika va elektronika loyihalashda CAD.",
        "CAD tizimlarini joriy etish bosqichlari va muammolari.",
    ],
    "sections": [
        ("CAD tizimlarining mashinasozlikda qo'llanilishi", [
            "Mashinasozlik sohasida CAD tizimlari detallar va mexanizmlarning 3D modellarini yaratish, yig'ma birikmalar loyihalash hamda konstruktorlik hujjatlarini tayyorlashda keng qo'llaniladi. SolidWorks, Inventor, KOMPAS kabi dasturlar chizmachilik standartlariga muvofiq ishchi chizmalar hosil qiladi.",
            "Parametrik modellashtirish texnologiyasi detalning istalgan o'lchamini o'zgartirish bilan bog'liq barcha elementlarning avtomatik yangilanishini ta'minlaydi. Bu konstruktor uchun turli variantlarni tezda sinab ko'rish imkonini beradi va loyihalash vaqtini sezilarli qisqartiradi.",
            "CAD tizimlari CAE (Computer-Aided Engineering) modullari bilan integratsiyalashib, detalning mustahkamligini, issiqlik holatini va dinamik xarakteristikalarini hisoblash imkonini beradi. Shu tariqa loyiha bosqichidayoq mahsulotning ishonchliligi tekshiriladi.",
        ]),
        ("Qurilish va arxitektura sohasida CAD tizimlari", [
            "Qurilish sohasida CAD tizimlari bino va inshootlarning plan, kesim va fasadlarini chizish, inshootlarning fazoviy modelini yaratish uchun ishlatiladi. AutoCAD, Revit va ArchiCAD bu sohaning asosiy vositalaridir.",
            "BIM (Building Information Modeling) texnologiyasi binoning nafaqat geometrik shaklini, balki materiallar, kommunikatsiyalar, smeta va qurilish jadvali haqidagi ma'lumotlarni yagona modelda saqlaydi. Bu qurilish ishtirokchilari o'rtasida samarali hamkorlikni ta'minlaydi.",
            "Arxitekturalik vizualizatsiya yordamida loyihaning fotorealistik rasmlari va animatsiyalari yaratiladi, bu buyurtmachiga loyiha tasavvurini aniq beradi. Virtual tur texnologiyasi esa qurilish boshlanmasdan binoni ichkaridan ko'rib chiqish imkonini beradi.",
        ]),
        ("Elektrotexnika va elektronika loyihalashda CAD", [
            "Elektrotexnika sohasida CAD tizimlari elektr sxemalarini chizish, simlar ulanish jadvallari va kabel yo'nalishlarini loyihalash uchun qo'llaniladi. AutoCAD Electrical va EPLAN bu sohaning mashhur vositalaridir.",
            "Elektron platalar (PCB) loyihalashda Altium Designer, KiCad va Eagle dasturlari ishlatiladi. Ular elektron sxemani chizishdan boshlab, plata trass yo'nalishlarini joylashtirish va ishlab chiqarishga fayllar tayyorlashgacha barcha bosqichlarni qamrab oladi.",
            "Mikroprotsessorli tizimlarni loyihalashda CAD dasturlari signallarning yaxlitligini tekshirish, elektromagnit muvofiqlikni tahlil qilish va issiqlik rejimini modellashtirish imkonini beradi. Bu murakkab elektron mahsulotlar sifatini oshiradi.",
        ]),
        ("CAD tizimlarini joriy etish bosqichlari va muammolari", [
            "CAD tizimini korxonaga joriy etish bir necha bosqichda amalga oshiriladi: ehtiyojlarni tahlil qilish, dastur tanlash, apparat ta'minotni tayyorlash, dasturni o'rnatish va sozlash, xodimlarni o'qitish hamda sinov ekspluatatsiyasi.",
            "Joriy etish jarayonida yuzaga kelishi mumkin bo'lgan muammolar: xodimlarning yangi texnologiyaga moslashishdagi qiyinchiliklari, katta boshlang'ich xarajatlar, mavjud hujjatlarni raqamli formatga o'tkazish zarurati va turli dasturlar o'rtasida ma'lumot almashinuvi masalalari.",
            "Muammolarni bartaraf etish uchun bosqichma-bosqich joriy etish strategiyasi, xodimlarni tizimli o'qitish va tajribali mutaxassislardan texnik yordam olish muhim. To'g'ri rejalashtirilgan joriy etish uzoq muddatda katta samara beradi.",
        ]),
    ],
})

TOPICS.append({
    "num": 7,
    "title": "Excel dasturida hujjatlar ustida ishlash",
    "reja": [
        "Microsoft Excel dasturi va uning ish maydoni tuzilishi.",
        "Ma'lumotlar kiritish, formatlash va tartibga solish.",
        "Formulalar va funksiyalar bilan ishlash.",
        "Diagrammalar yaratish va ma'lumotlarni tahlil qilish vositalari.",
    ],
    "sections": [
        ("Microsoft Excel dasturi va uning ish maydoni tuzilishi", [
            "Microsoft Excel — elektron jadval (spreadsheet) dasturi bo'lib, raqamli ma'lumotlarni kiritish, tartibga solish, hisoblash va tahlil qilishga mo'ljallangan. U Microsoft Office paketining ajralmas qismidir va texnik sohalarda keng qo'llaniladi.",
            "Excel ish maydoni ishchi kitob (workbook) va varaqlar (worksheets) dan tashkil topadi. Har bir varaq ustun (A, B, C...) va satrlardan (1, 2, 3...) iborat bo'lib, ularning kesishishida katakcha (cell) joylashadi. Har bir katakchaning o'z manzili bor (masalan, A1, B5).",
            "Dastur interfeysi yuqorida tasma (Ribbon), formula qatori va ish maydoni joylashgan. Pastda varaqlar orasida o'tish yorliqlari va holat qatori mavjud. Bu tuzilish ma'lumotlarni qulay tartibda joylashtirish va ular ustida turli amallar bajarish imkonini beradi.",
        ]),
        ("Ma'lumotlar kiritish, formatlash va tartibga solish", [
            "Excel katakchalariga sonlar, matnlar, sanalar va formulalar kiritiladi. Ma'lumotni kiritgach Enter tugmasi bosiladi. Bir nechta katakchaga bir xil ma'lumotni tez kiritish uchun avtoto'ldirish (AutoFill) funksiyasidan foydalaniladi — masalan, ketma-ket raqamlar yoki hafta kunlari.",
            "Katakchalarni formatlash orqali shrift, chegara, fon rangi, raqam formati (valyuta, foiz, sana) va tekislash sozlanadi. Shartli formatlash (Conditional Formatting) ma'lum shartlarga javob beruvchi katakchalarni avtomatik ranglaydi, bu esa ma'lumotdagi muhim qiymatlarni ajratib ko'rsatadi.",
            "Ma'lumotlarni saralash (Sort) va filtrlash (Filter) vositalari katta jadvallardan kerakli ma'lumotni tez topish imkonini beradi. Saralash bir yoki bir nechta ustun bo'yicha o'sish yoki kamayish tartibida amalga oshiriladi.",
        ]),
        ("Formulalar va funksiyalar bilan ishlash", [
            "Formulalar Excel ning eng muhim vositasi bo'lib, katakchalardagi qiymatlar ustida avtomatik hisob-kitoblar bajaradi. Formulalar doimo '=' belgisi bilan boshlanadi. Masalan, =A1+B1 ifodasi ikkita katakcha qiymatlarini qo'shadi.",
            "Excelda 400 dan ortiq tayyor funksiyalar mavjud. Eng ko'p ishlatiladiganlari: SUM (yig'indisi), AVERAGE (o'rta qiymati), MAX, MIN (eng katta va kichik), COUNT (soni), IF (shart tekshirish). Funksiyalar murakkab hisob-kitoblarni soddalashtiradi.",
            "Nisbiy va mutlaq havolalar formulalarni nusxalashda muhim. Nisbiy havola (A1) nusxalanganda avtomatik o'zgaradi, mutlaq havola ($A$1) esa o'zgarmaydi. Aralash havolalar ($A1 yoki A$1) faqat bitta yo'nalishda qotib turadi.",
        ]),
        ("Diagrammalar yaratish va ma'lumotlarni tahlil qilish vositalari", [
            "Excelda ma'lumotlarni ko'rgazmali ifodalash uchun diagrammalar ishlatiladi. Ma'lumotni belgilab, 'Qo'yish' yorlig'idagi diagramma turini tanlash kifoya. Ustunli, chiziqli, doiraviy, nuqtali va boshqa diagramma turlari mavjud.",
            "Diagramma elementlari — sarlavha, oq (legend), ma'lumot belgilari va o'qlar — sozlanishi mumkin. Diagramma ma'lumotlar bilan bog'liq bo'lib, jadval qiymatlari o'zgarganda avtomatik yangilanadi. Bu dinamik tahlil va hisobotlar tayyorlashda qulaydir.",
            "Excel ma'lumotlarni tahlil qilishning kuchli vositalariga ega: yig'ma jadvallar (Pivot Table) katta hajmdagi ma'lumotlarni guruhlab xulosa chiqaradi, 'Nima bo'ladi-agar' (What-If) tahlili esa turli stsenariylarni modellashtirish imkonini beradi.",
        ]),
    ],
})

TOPICS.append({
    "num": 8,
    "title": "Texnik tizimlarida tarmoq texnologiyalari. Tarmoq texnologiyalaridan foydalanish",
    "reja": [
        "Kompyuter tarmoqlari tushunchasi va tasnifi.",
        "Tarmoq qurilmalari va aloqa muhitlari.",
        "TCP/IP protokoli va Internet ishlash prinsipi.",
        "Texnik sohalarda tarmoq texnologiyalarini qo'llash.",
    ],
    "sections": [
        ("Kompyuter tarmoqlari tushunchasi va tasnifi", [
            "Kompyuter tarmog'i — bu axborot almashish va resurslardan birgalikda foydalanish maqsadida o'zaro bog'langan kompyuterlar va qurilmalar majmuasidir. Tarmoqlar bir nechta kompyuterning printerdan birgalikda foydalanishidan tortib, butun dunyo bo'ylab axborot almashishgacha keng spektrni qamrab oladi.",
            "Tarmoqlar qamroviga ko'ra tasniflanadi: LAN (mahalliy tarmoq) — bitta bino yoki korxona ichidagi tarmoq; WAN (keng tarmoq) — geografik jihatdan keng hududni qamrab oluvchi tarmoq; MAN (shahar tarmog'i) — shahar miqyosidagi tarmoq. Internet esa eng katta global tarmoqdir.",
            "Topologiyasi bo'yicha tarmoqlar yulduzsimon (star), shinali (bus), halqasimon (ring) va aralash turlarga bo'linadi. Har bir topologiyaning o'z afzalliklari va kamchiliklari bor: masalan, yulduzsimon topologiyada bitta qurilma uzilsa, boshqalarga ta'sir qilmaydi.",
        ]),
        ("Tarmoq qurilmalari va aloqa muhitlari", [
            "Tarmoq qurilmalariga marshrutizator (router), kommutator (switch), konsentrator (hub), tarmoq adapteri (NIC) va simsiz nuqta (access point) kiradi. Marshrutizator tarmoqlar orasida ma'lumotlarni yo'naltiradi, kommutator esa mahalliy tarmoq ichida paketlarni to'g'ri adresga yetkazadi.",
            "Aloqa muhitlari simli va simsiz turlarga bo'linadi. Simli muhitga o'rama juftlik (twisted pair, Cat5e/Cat6), optik tolali kabel (fiber optic) va koaksial kabel kiradi. Optik kabel eng yuqori tezlik va uzoq masofani ta'minlaydi.",
            "Simsiz texnologiyalar (Wi-Fi, Bluetooth, 4G/5G) kabelga bo'lgan ehtiyojni bartaraf etadi va harakatchanlik beradi. Wi-Fi tarmog'i mahalliy muhitda, mobil tarmoqlar esa keng hududda simsiz ulanish imkonini beradi.",
        ]),
        ("TCP/IP protokoli va Internet ishlash prinsipi", [
            "TCP/IP — Internetning asosiy protokollar to'plamidir. TCP (Transmission Control Protocol) ma'lumotlarni paketlarga bo'lib, ishonchli yetkazilishini ta'minlaydi; IP (Internet Protocol) esa har bir paketni manzilga yo'naltiradi. Bu ikki protokol birgalikda tarmoq aloqasining asosini tashkil etadi.",
            "Internetda har bir qurilmaga noyob IP-manzil beriladi. DNS (Domain Name System) xizmati domen nomlarini (masalan, www.example.com) IP-manzillarga aylantiradi. Foydalanuvchi brauzerda manzil kiritganda, DNS so'rovi orqali kerakli server topiladi va bog'lanish o'rnatiladi.",
            "Internet orqali elektron pochta (SMTP/POP3/IMAP), veb-sahifalar (HTTP/HTTPS), fayllar uzatish (FTP) va boshqa xizmatlar ishlaydi. Har bir xizmat maxsus protokolga asoslanadi va muayyan port raqami orqali amalga oshiriladi.",
        ]),
        ("Texnik sohalarda tarmoq texnologiyalarini qo'llash", [
            "Texnik sohalarda tarmoq texnologiyalari sanoat korxonalarini boshqarish (SCADA tizimlari), masofaviy monitoring, loyihalarni birgalikda ishlash va ma'lumotlarni markazlashtirilgan saqlash uchun qo'llaniladi.",
            "Narsalar Interneti (IoT) texnologiyasi sensorlar va qurilmalarni tarmoqqa ulash orqali texnik tizimlarni masofadan kuzatish va boshqarish imkonini beradi. Masalan, ishlab chiqarish liniyasidagi harorat va bosim datchiklarining ko'rsatkichlari real vaqtda markaziy serverga uzatiladi.",
            "Bulutli texnologiyalar (cloud computing) korxonalarga o'z serverlarini sotib olmasdan, masofaviy serverlarda ma'lumotlarni saqlash va dasturlardan foydalanish imkonini beradi. Bu xarajatlarni kamaytiradi va resurslardan moslashuvchan foydalanishni ta'minlaydi.",
        ]),
    ],
})

TOPICS.append({
    "num": 9,
    "title": "MathCad dasturida matematik hisoblashlarni bajarish",
    "reja": [
        "MathCad dasturi haqida umumiy ma'lumot va interfeysi.",
        "O'zgaruvchilar, formulalar va matematik ifodalarni kiritish.",
        "Tenglamalar yechish va grafiklar qurish.",
        "MathCad ning texnik hisob-kitoblardagi qo'llanilishi.",
    ],
    "sections": [
        ("MathCad dasturi haqida umumiy ma'lumot va interfeysi", [
            "MathCad — muhandislik va ilmiy hisoblashlarni bajarish uchun mo'ljallangan dasturiy mahsulotdir. Uning asosiy xususiyati shundaki, matematik ifodalar daftardagi kabi tabiiy ko'rinishda kiritiladi va natijalar shu yerda ko'rsatiladi. Bu uni oddiy dasturlash tillaridan farqlaydi.",
            "MathCad interfeysi hujjat varag'iga o'xshash bo'lib, foydalanuvchi varoq bo'ylab formulalar, matnlar va grafiklarni erkin joylashtiradi. Yuqoridagi asboblar paneli matematik operatorlar, grafiklar va matritsa vositalari uchun tugmalarni o'z ichiga oladi.",
            "Dastur kuchli hisoblash imkoniyatlariga ega: oddiy arifmetik amallardan tortib, differentsial tenglamalar yechish, statistik tahlil va optimizatsiyagacha. Natijalar birliklarni hisobga olgan holda chiqariladi, bu muhandislik hisob-kitoblarida xatolarni kamaytiradi.",
        ]),
        ("O'zgaruvchilar, formulalar va matematik ifodalarni kiritish", [
            "MathCad da o'zgaruvchi tayinlash ':=' operatori yordamida amalga oshiriladi (masalan, x:=5). O'zgaruvchiga son, matn yoki ifoda qiymati berilishi mumkin. Bir marta aniqlangan o'zgaruvchi hujjatning keyingi qismlarida ishlatiladi.",
            "Formulalar klaviatura va matematik panel yordamida kiritiladi. Kasrlar, ildizlar, darajalar, indekslar va boshqa matematik belgilar tabiiy ko'rinishda yoziladi. Masalan, kvadrat ildiz belgisini kiritish uchun maxsus tugma bosiladi va ostiga ifoda yoziladi.",
            "MathCad birliklar tizimini qo'llab-quvvatlaydi: o'zgaruvchiga metr, kilogramm, sekund kabi birliklar biriktirilishi mumkin. Dastur birliklarning muvofiqligini avtomatik tekshiradi va natijani kerakli birlikda ifodalaydi. Bu fizik hisob-kitoblarda aniqlikni oshiradi.",
        ]),
        ("Tenglamalar yechish va grafiklar qurish", [
            "MathCad da algebraik tenglamalarni yechish uchun solve funksiyasi yoki Given-Find bloki ishlatiladi. Given kalit so'zidan keyin tenglik va tengsizliklar yoziladi, Find funksiyasi esa noma'lumlarning qiymatini topadi. Bir nechta noma'lumli tenglamalar tizimini ham yechish mumkin.",
            "Grafiklar 2D va 3D ko'rinishda quriladi. 2D grafik yaratish uchun XY-Plot maydoni kiritiladi, gorizontal o'qqa argument, vertikal o'qqa esa funksiya yoziladi. Bir nechta funksiyaning grafiklari bitta koordinata tizimida tasvirlanishi mumkin.",
            "3D grafiklar sirt (surface), kontur va fazoviy egri chiziqlarni ifodalaydi. Ular ikki o'zgaruvchili funksiyalarni vizualizatsiya qilish uchun ishlatiladi. Grafiklar ranglar, masshtab va burchak bo'yicha sozlanishi mumkin.",
        ]),
        ("MathCad ning texnik hisob-kitoblardagi qo'llanilishi", [
            "MathCad mashinasozlikda detallar mustahkamligini hisoblash, elektrotexnikada zanjir parametrlarini tahlil qilish, qurilishda yuk ko'tarish qobiliyatini tekshirishda keng qo'llaniladi. Uning formulalarni tabiiy shaklda ko'rsatishi hisob-kitob hujjatlarini tushunishni osonlashtiradi.",
            "Dastur differentsial tenglamalarni sonli usullar bilan yechish imkoniyatiga ega bo'lib, bu dinamik tizimlarni modellashtirish uchun muhimdir. Masalan, tebranish jarayonlari, issiqlik tarqalishi va suyuqlik oqimi tahlil qilinishi mumkin.",
            "MathCad hujjatni Word yoki PDF formatida eksport qilish imkonini beradi, bu esa hisob-kitob natijalarini hisobotlarga kiritishni osonlashtiradi. Shunday qilib, dastur muhandis uchun nafaqat hisoblash vositasi, balki hisob-kitob hujjatlarini rasmiylashtirishning qulay vositasidir.",
        ]),
    ],
})

TOPICS.append({
    "num": 10,
    "title": "Zamonaviy dasturlash texnologiyalari",
    "reja": [
        "Dasturlash tushunchasi va uning rivojlanish bosqichlari.",
        "Zamonaviy dasturlash paradigmalari.",
        "Integratsiyalashgan ishlab chiqish muhitlari (IDE).",
        "Dasturlash texnologiyalarining texnik sohalardagi o'rni.",
    ],
    "sections": [
        ("Dasturlash tushunchasi va uning rivojlanish bosqichlari", [
            "Dasturlash — bu kompyuterga muayyan vazifani bajarish uchun buyruqlar ketma-ketligini (dastur) yaratish jarayonidir. Dasturchi masalani tahlil qilib, uni kompyuter tushunadigan tilga o'giradi. Dasturlash kompyuter fanining eng muhim yo'nalishlaridan biridir.",
            "Dasturlashning rivojlanish bosqichlari: birinchi avlod — mashina tili (ikkilik kodlar); ikkinchi avlod — assembler (mnemonik buyruqlar); uchinchi avlod — yuqori darajali tillar (C, Pascal, Fortran); to'rtinchi avlod — muammoga yo'naltirilgan tillar (SQL); beshinchi avlod — sun'iy intellekt tillari va vizual dasturlash.",
            "Har bir yangi avlod dasturlashni inson uchun qulayroq va abstrakt darajasini yuqoriroq qildi. Zamonaviy tillar murakkab tizimlarni kamroq kod bilan yaratish imkonini beradi, bu esa ishlab chiqish tezligini oshiradi.",
        ]),
        ("Zamonaviy dasturlash paradigmalari", [
            "Dasturlash paradigmasi — bu dastur yozish usuli va falsafasidir. Strukturali dasturlash dasturni ketma-ket, tarmoqlanuvchi va takrorlanuvchi tuzilmalar orqali yozishni nazarda tutadi. Bu yondashuv kodni tartibli va tushunarli qiladi.",
            "Obyektga yo'naltirilgan dasturlash (OOP) — zamonaviy dasturlashning eng keng tarqalgan paradigmasi bo'lib, ma'lumot va funksiyalarni obyekt sifatida birlashtirishga asoslanadi. OOP ning asosiy tamoyillari: inkapsulyatsiya, vorislik, polimorfizm va abstraksiya.",
            "Funktsional dasturlash matematik funksiyalarga asoslanib, holat o'zgarishlarisiz va yon ta'sirsiz dasturlashni ko'zda tutadi. Reaktiv dasturlash esa ma'lumot oqimlariga reaksiya qilish tamoyiliga asoslanadi. Zamonaviy tillar ko'pincha bir nechta paradigmani qo'llab-quvvatlaydi.",
        ]),
        ("Integratsiyalashgan ishlab chiqish muhitlari (IDE)", [
            "IDE (Integrated Development Environment) — dasturchi uchun barcha zarur vositalarni birlashtirgan dasturiy muhitdir. U matn muharriri, kompilyator/interpretator, tuzatuvchi (debugger) va loyiha boshqaruvchisini o'z ichiga oladi.",
            "Mashhur IDE larga Visual Studio, Visual Studio Code, IntelliJ IDEA, Eclipse va PyCharm kiradi. Ular kodni rang bilan ajratish (syntax highlighting), avtoto'ldirish (autocomplete), xatolarni real vaqtda ko'rsatish va versiyalarni boshqarish integratsiyasini taqdim etadi.",
            "IDE foydalanish dasturchi samaradorligini sezilarli oshiradi: xatolar tezroq topiladi, kod yozish tezlashadi va loyihaning barcha fayllari bilan qulay ishlash imkoniyati yaratiladi. Zamonaviy IDE lar plaginlar orqali yangi tillar va texnologiyalarni qo'llab-quvvatlashi mumkin.",
        ]),
        ("Dasturlash texnologiyalarining texnik sohalardagi o'rni", [
            "Texnik sohalarda dasturlash texnologiyalari avtomatlashtirish tizimlarini yaratish, qurilmalarni boshqaruvchi dasturlar yozish, ma'lumotlarni qayta ishlash va modellashtirish uchun qo'llaniladi. Har bir soha o'ziga mos tillar va vositalardan foydalanadi.",
            "Masalan, o'rnatilgan tizimlar (embedded systems) uchun C va C++ tillari, ma'lumotlar tahlili uchun Python, veb-ilovalar uchun JavaScript va mobil ilovalar uchun Swift yoki Kotlin ishlatiladi. Ilmiy hisob-kitoblarda MATLAB va Python keng tarqalgan.",
            "Zamonaviy dasturlash texnologiyalari DevOps, bulutli hisoblash, mikroservislar arxitekturasi va sun'iy intellekt kabi yo'nalishlarni ham qamrab oladi. Bu texnologiyalar dasturlarni ishlab chiqish, sinovdan o'tkazish va foydalanuvchilarga yetkazish jarayonini tezlashtiradi va sifatini oshiradi.",
        ]),
    ],
})



TOPICS.append({
    "num": 11,
    "title": "Dasturlash tillari va tizimlari, ularning ishlatilishi va tasnifi. Algoritmlar",
    "reja": [
        "Dasturlash tillarining tasnifi va avlodlari.",
        "Dasturlash tizimlarining tarkibi va ishlash prinsipi.",
        "Algoritm tushunchasi va uning xossalari.",
        "Algoritmlarni ifodalash usullari.",
    ],
    "sections": [
        ("Dasturlash tillarining tasnifi va avlodlari", [
            "Dasturlash tillari turli mezonlarga ko'ra tasniflanadi. Mashina tiliga yaqinlik darajasiga ko'ra past darajali (mashina tili, assembler) va yuqori darajali (C, C++, Java, Python) tillarga bo'linadi. Past darajali tillar protsessorga yaqin bo'lib, tez ishlaydi, lekin murakkab; yuqori darajali tillar esa inson tiliga yaqinroq.",
            "Dasturlash paradigmasi bo'yicha tillar protsedurali (C, Pascal), obyektga yo'naltirilgan (C++, Java, C#), funktsional (Haskell, Lisp) va skript (Python, JavaScript) tillarga bo'linadi. Ko'p zamonaviy tillar bir nechta paradigmani qo'llab-quvvatlaydi.",
            "Qo'llanilish sohasiga ko'ra tizimli dasturlash tillari (C, C++), ilmiy hisoblash tillari (Fortran, MATLAB), veb-dasturlash (JavaScript, PHP), ma'lumotlar bazasi (SQL) va maxsus maqsadli tillar (R, VHDL) ajratiladi.",
        ]),
        ("Dasturlash tizimlarining tarkibi va ishlash prinsipi", [
            "Dasturlash tizimi — dasturchiga dastur yaratish, tekshirish va bajarish imkonini beruvchi vositalar majmuasidir. Uning asosiy komponentlari: matn muharriri, translyator (kompilyator yoki interpretator), bog'lovchi (linker) va tuzatuvchi (debugger).",
            "Kompilyator butun dasturni mashina tiliga bir marotaba to'liq tarjima qiladi va bajariluvchi fayl hosil qiladi (C, C++). Interpretator esa dasturni satr-satr o'qib bajaradi (Python, JavaScript). Har bir yondashuvning o'z afzalliklari bor: kompilyatsiya tezroq ishlaydi, interpretatsiya esa tuzatishni osonlashtiradi.",
            "Zamonaviy dasturlash tizimlari IDE ko'rinishida taqdim etiladi, ularga versiyalarni boshqarish (Git), paket menejerlari va avtomatik test vositalari ham integratsiyalashgan. Bu dastur ishlab chiqish jarayonini yaxlit va qulay qiladi.",
        ]),
        ("Algoritm tushunchasi va uning xossalari", [
            "Algoritm — biror masalani yechish uchun aniq belgilangan, chekli sondagi buyruqlar (qadamlar) ketma-ketligidir. Algoritm noaniq yoki ikki ma'noli bo'lmasligi kerak — har bir qadamdan keyin keyingisi aniq bo'lishi lozim.",
            "Algoritmning asosiy xossalari: diskretlik (bosqichma-bosqich bajarilish), tushunarlilik (ijrochiga aniq buyruqlar), aniqlik (har bir qadamning yagona talqini), natijaviylik (chekli qadamda natijaga erishish) va ommaviylik (bir xil turdagi masalalar uchun yaroqlilik).",
            "Algoritm tushunchasi dasturlashning asosidir — har qanday dastur mohiyatan algoritmning biror dasturlash tilida yozilgan shaklidir. Shu sababli masalani avval algoritmik jihatdan yechib, keyin dasturga o'girish to'g'ri yondashuv hisoblanadi.",
        ]),
        ("Algoritmlarni ifodalash usullari", [
            "Algoritmlarni ifodalashning bir nechta usuli mavjud: so'zlar bilan (tabiiy tilda), blok-sxema ko'rinishida, psevdokod yordamida va dasturlash tilida. Har bir usul o'z qulaylik va aniqlik darajasiga ega.",
            "Blok-sxema — algoritmni grafik shaklda ifodalash usuli bo'lib, turli geometrik shakllar ishlatiladi: to'g'ri to'rtburchak — amal (jarayon), romb — shart tekshirish, parallelogramm — kiritish/chiqarish, oval — boshlanishi va tugashi. Bu usul algoritmni ko'rgazmali tasvirlaydi.",
            "Psevdokod — algoritmni dasturlash tiliga yaqin, lekin formal sintaksisga bog'lanmagan shaklda ifodalash usuli. U so'zlar va tuzilmalarni ishlatsada, aniq tilning qoidalariga qat'iy rioya qilmaydi. Bu esa algoritmni tez yozish va tushunishni osonlashtiradi.",
        ]),
    ],
})

TOPICS.append({
    "num": 12,
    "title": "Matlab dasturida matematik hisoblashlarni bajarish",
    "reja": [
        "MATLAB dasturi haqida umumiy ma'lumot va uning imkoniyatlari.",
        "MATLAB muhitida o'zgaruvchilar va matritsa amallari.",
        "MATLAB da grafik qurish va vizualizatsiya.",
        "MATLAB ning muhandislik masalalarida qo'llanilishi.",
    ],
    "sections": [
        ("MATLAB dasturi haqida umumiy ma'lumot va uning imkoniyatlari", [
            "MATLAB (Matrix Laboratory) — MathWorks kompaniyasi tomonidan ishlab chiqilgan matematik hisoblash, ma'lumotlar tahlili va vizualizatsiya uchun mo'ljallangan dasturiy muhitdir. Uning nomi matritsa laboratoriyasi ma'nosini beradi, chunki asosiy ma'lumot turi matritsadir.",
            "MATLAB o'z ichiga dasturlash tilini, interaktiv muhitni, grafiklar yaratish vositalarini va ko'plab maxsus kutubxonalarni (toolbox) oladi. Toolbox larga signal qayta ishlash, boshqaruv nazariyasi, statistika, tasvirlarni qayta ishlash va boshqalar kiradi.",
            "Dastur mashinasozlik, elektrotexnika, aviatsiya, biotexnologiya va boshqa sohalarda muhandislik hisob-kitoblari uchun standart vositaga aylangan. U o'quv jarayonida ham keng qo'llaniladi.",
        ]),
        ("MATLAB muhitida o'zgaruvchilar va matritsa amallari", [
            "MATLAB da o'zgaruvchi oldindan e'lon qilmasdan to'g'ridan-to'g'ri qiymat berish orqali yaratiladi (masalan, a = 5). Barcha o'zgaruvchilar avtomatik ravishda matritsa sifatida saqlanadi: skalyar — 1x1 matritsa, vektor — 1xn yoki nx1 matritsa, ikki o'lchovli massiv esa mxn matritsa.",
            "Matritsa amallari MATLAB ning kuchli tomoni: qo'shish (+), ayirish (-), ko'paytirish (*), transponerlash ('), teskari matritsa (inv), determinant (det) va xos qiymatlar (eig). Nuqta bilan boshlangan amallar (.*, ./, .^) elementma-element hisoblash uchun ishlatiladi.",
            "MATLAB buyruq oynasi (Command Window) da ifodalar kiritib natija olish mumkin, shuningdek m-fayllar (skriptlar va funksiyalar) yaratib murakkab hisob-kitoblarni avtomatlashtirish mumkin. Workspace oynasi barcha joriy o'zgaruvchilarni ko'rsatadi.",
        ]),
        ("MATLAB da grafik qurish va vizualizatsiya", [
            "MATLAB da 2D grafik qurish uchun plot funksiyasi ishlatiladi: plot(x, y) — x va y vektorlariga mos nuqtalarni chiziq bilan bog'laydi. Chiziq rangi, turi va markerlari maxsus parametrlar bilan sozlanadi. Bir oynada bir nechta grafik chizish uchun hold on buyrug'i ishlatiladi.",
            "3D grafiklar uchun plot3, surf, mesh va contour funksiyalari mavjud. surf funksiyasi uch o'lchovli sirtni rangli ko'rinishda tasvirlaydi. meshgrid funksiyasi ikki o'zgaruvchi uchun koordinatalar to'ri hosil qiladi.",
            "Grafikka sarlavha (title), o'q nomlari (xlabel, ylabel), izoh (legend) va to'r chiziqlari (grid on) qo'shiladi. subplot funksiyasi bitta oynada bir nechta grafikni alohida maydonlarda joylashtirish imkonini beradi. Bu imkoniyatlar tahlil natijalarini ko'rgazmali taqdim etadi.",
        ]),
        ("MATLAB ning muhandislik masalalarida qo'llanilishi", [
            "MATLAB muhandislikda chiziqli algebraik tenglamalar tizimini yechish, differentsial tenglamalar yechish, signal qayta ishlash va boshqaruv tizimlarini modellashtirish uchun keng qo'llaniladi. Chiziqli tenglamalar tizimi A*x = b ko'rinishida x = A\\b buyrug'i bilan bir qadamda yechiladi.",
            "Simulink — MATLAB ning grafik muhiti bo'lib, u dinamik tizimlarni blok-sxemalar yordamida modellashtirish imkonini beradi. Elektr zanjirlar, mexanik tizimlar va boshqaruv tizimlarining o'tkinchi jarayonlari Simulink yordamida vizual tarzda tahlil qilinadi.",
            "MATLAB skriptlarini C/C++ kodiga generatsiya qilish va mustaqil dastur sifatida kompilatsiya qilish mumkin. Bu prototipdan tayyor mahsulotga o'tishni tezlashtiradi va o'rnatilgan tizimlarga joylashtirish imkonini beradi.",
        ]),
    ],
})

TOPICS.append({
    "num": 13,
    "title": "C++ dasturlash tili va muhiti",
    "reja": [
        "C++ tilining tarixi va xususiyatlari.",
        "C++ dasturlash muhitlari va kompilyatorlari.",
        "Birinchi dastur tuzish va kompilyatsiya jarayoni.",
        "C++ tilining qo'llanilish sohalari.",
    ],
    "sections": [
        ("C++ tilining tarixi va xususiyatlari", [
            "C++ dasturlash tili 1979-yilda Byarne Stroustrup tomonidan C tilining kengaytmasi sifatida yaratilgan. Dastlab 'C with Classes' deb nomlangan bu til 1983-yilda C++ nomini oldi. '++' belgisi C tilidagi oshirish (increment) operatorini anglatib, tilning C dan bir qadam oldinda ekanini bildiradi.",
            "C++ tilining asosiy xususiyatlari: yuqori samaradorlik (mashina tiliga yaqin tezlikda ishlaydi), obyektga yo'naltirilgan dasturlashni qo'llab-quvvatlash, past darajali xotira bilan ishlash imkoniyati va kuchli standart kutubxona (STL). Bu xususiyatlar uni universal va kuchli dasturlash tiliga aylantiradi.",
            "C++ tili ISO tomonidan standartlashtirilgan va muntazam yangilanib turadi: C++98, C++03, C++11, C++14, C++17, C++20, C++23. Har bir yangi standart tilga zamonaviy imkoniyatlar qo'shadi: lambda funksiyalar, avtomatik tip aniqlash, oqimli dasturlash vositalari va boshqalar.",
        ]),
        ("C++ dasturlash muhitlari va kompilyatorlari", [
            "C++ da dastur yozish uchun matn muharriri va kompilyator talab qilinadi. Kompilyator manba kodni (source code) mashina tiliga tarjima qiladi. Eng mashhur C++ kompilyatorlari: GCC (GNU Compiler Collection), MSVC (Microsoft Visual C++), va Clang. Ular turli operatsion tizimlarni qo'llab-quvvatlaydi.",
            "Amalda ko'pincha IDE (integratsiyalashgan ishlab chiqish muhiti) dan foydalaniladi: Visual Studio (Windows uchun kuchli muhit), Code::Blocks (bepul va yengil), CLion (professional muhit), Dev-C++ (o'rganish uchun qulay). IDE kompilyatsiya, tuzatish va loyiha boshqaruvini birlashtiradi.",
            "Dasturlash muhitini o'rnatgandan keyin oddiy dastur yozib, kompilyatsiya va bajarish jarayonlarini sinab ko'rish muhim. Bu muhit to'g'ri sozlanganini va dasturchi ish jarayonini tushunganini tekshirishga yordam beradi.",
        ]),
        ("Birinchi dastur tuzish va kompilyatsiya jarayoni", [
            "C++ da an'anaviy birinchi dastur 'Hello, World!' — ekranga salomlashish xabarini chiqaradi. Dastur #include <iostream> direktivasi (kiritish-chiqarish kutubxonasi) va main() funksiyasidan boshlanadi. cout oqimi va << operatori orqali matn ekranga chiqariladi.",
            "Kompilyatsiya jarayoni bir necha bosqichdan iborat: preprosessor (#include va #define direktivalarini qayta ishlaydi), kompilyatsiya (manba kodni obyekt kodga aylantiradi), bog'lash (linker — obyekt fayllar va kutubxonalarni birlashtiradi) va natijada bajariluvchi (executable) fayl hosil bo'ladi.",
            "Dastur yozilgandan so'ng kompilyator xato bersa, sintaksis (yozilish) xatolari mavjud. Kompilyatsiya muvaffaqiyatli bo'lsa-da, dastur noto'g'ri ishlasa, mantiqiy xato bor. Tuzatuvchi (debugger) yordamida dasturni qadamba-qadam bajarib, xatoni topish mumkin.",
        ]),
        ("C++ tilining qo'llanilish sohalari", [
            "C++ tili o'zining yuqori samaradorligi tufayli operatsion tizimlar (Windows, Linux yadro qismlari), o'yin dvigatellari (Unreal Engine), brauzerlar (Chrome, Firefox) va ma'lumotlar bazasi tizimlari (MySQL) yaratishda keng ishlatiladi.",
            "O'rnatilgan tizimlar (embedded systems) va mikrokontrollerlarni dasturlashda C++ muhim o'rin tutadi: tibbiy asboblar, avtomobil elektronikasi, aviatsiya tizimlari va sanoat robotlarida qo'llaniladi. Bu sohalarda dasturning tezligi va resurslardan tejamli foydalanishi hal qiluvchi ahamiyatga ega.",
            "Moliya sohasida yuqori chastotali savdo (HFT) tizimlari, ilmiy hisoblashlarda fizika simulyatsiyalari va sun'iy intellekt kutubxonalarining bazaviy qismlari C++ da yoziladi. Bu tilning universalligi va kuchi uning dolzarbligini uzoq yillar davomida saqlab kelmoqda.",
        ]),
    ],
})

TOPICS.append({
    "num": 14,
    "title": "Dastur elementlari. Alifbosi. Identifikatorlar",
    "reja": [
        "C++ dasturining umumiy tuzilishi.",
        "C++ tilining alifbosi (belgilar to'plami).",
        "Identifikatorlar va ularning nomlash qoidalari.",
        "Kalit so'zlar va izohlar.",
    ],
    "sections": [
        ("C++ dasturining umumiy tuzilishi", [
            "C++ dasturi bir yoki bir nechta fayldan iborat bo'lib, asosiy fayl odatda main() funksiyasini o'z ichiga oladi. Dastur bajarilishi aynan shu funksiyadan boshlanadi. Dastur tuzilishi: preprosessor direktivalar, global e'lonlar, funksiyalar va main() funksiya tanasi.",
            "Preprosessor direktivalar (#include, #define) dastur boshida yoziladi va kompilyatsiyadan oldin qayta ishlanadi. #include <iostream> standart kiritish-chiqarish kutubxonasini ulaydi. Namespace (nomlash sohasi) e'loni — using namespace std; — standart kutubxona elementlarini qisqa yozish imkonini beradi.",
            "Dastur tanasi figurali qavslar { } ichida yoziladi. Har bir buyruq (operator) nuqtali vergul (;) bilan tugaydi. C++ da katta-kichik harflar farqlanadi (case-sensitive): 'Sum' va 'sum' turli identifikatorlardir. Dastur tuzilishini to'g'ri bilish — uning to'g'ri ishlashining asosidir.",
        ]),
        ("C++ tilining alifbosi (belgilar to'plami)", [
            "C++ tilining alifbosi — dastur yozishda ishlatish mumkin bo'lgan belgilar to'plamidir. Unga lotin harflari (A-Z, a-z), raqamlar (0-9), maxsus belgilar va bo'sh joy belgilari kiradi. Kirill harflari faqat satr konstantalar va izohlar ichida ishlatilishi mumkin.",
            "Maxsus belgilar turli maqsadlarda ishlatiladi: arifmetik amallar (+, -, *, /, %), taqqoslash (==, !=, <, >), tayinlash (=), mantiqiy operatorlar (&&, ||, !), qavslar ((), {}, []), ajratuvchilar (;, :, ,) va boshqalar. Har bir belgining aniq vazifasi bor.",
            "Bo'sh joy belgilariga probel, tabulyatsiya va yangi satr belgisi kiradi. Ular dastur o'qilishini yaxshilash uchun ishlatiladi. Kompilyator ortiqcha bo'sh joylarni e'tiborsiz qoldiradi (satr konstantalar ichidagidan tashqari). Dasturni chiroyli formatlash — professional dasturchi odati.",
        ]),
        ("Identifikatorlar va ularning nomlash qoidalari", [
            "Identifikator — bu dasturchi tomonidan o'zgaruvchi, funksiya, klass yoki boshqa elementga beriladigan nomdir. U dasturning 'lug'ati' bo'lib, elementlarni bir-biridan farqlash uchun ishlatiladi.",
            "Identifikator nomlash qoidalari: faqat lotin harflari, raqamlar va pastki chiziq (_) ishlatiladi; raqam bilan boshlanmaydi; bo'sh joy bo'lmaydi; kalit so'zlar nom sifatida ishlatilmaydi; uzunligi cheklanmagan, lekin qisqa va mazmunli nom tanlash tavsiya etiladi.",
            "Yaxshi nomlash amaliyotlari: o'zgaruvchilarga mazmunli nom berish (masalan, radius, count, temperature); camelCase yoki snake_case uslubidan foydalanish; konstantalarga KATTA_HARF ishlatish. Bu dasturni o'qish va tuzatishni osonlashtiradi.",
        ]),
        ("Kalit so'zlar va izohlar", [
            "Kalit so'zlar — C++ tilida maxsus ma'noga ega bo'lgan zaxiralangan so'zlardir. Ularni identifikator sifatida ishlatish mumkin emas. Masalan: int, float, double, char, if, else, for, while, return, class, void, bool, true, false, new, delete kabilar.",
            "C++ standartida 90 dan ortiq kalit so'z mavjud. Har biri tilning muayyan konstruktsiyasini ifodalaydi: ma'lumot turlari (int, double), boshqaruv tuzilmalari (if, for, while), klass yaratish (class, struct) va xotira boshqarish (new, delete).",
            "Izohlar (comments) — dasturga tushuntirish maqsadida yoziladigan matn bo'lib, kompilyator tomonidan e'tiborsiz qoldiriladi. Bir qatorli izoh // belgisidan boshlanadi, ko'p qatorli izoh esa /* va */ orasida yoziladi. Izohlar dasturni boshqalar uchun tushunarliroq qiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 15,
    "title": "Algoritmlar: Chiziqli, tarmoqlanuvchi va takrorlanuvchi",
    "reja": [
        "Chiziqli algoritmlar va ularning xususiyatlari.",
        "Tarmoqlanuvchi algoritmlar va shart tushunchasi.",
        "Takrorlanuvchi (siklik) algoritmlar.",
        "Algoritmlarni blok-sxema va dastur shaklida ifodalash.",
    ],
    "sections": [
        ("Chiziqli algoritmlar va ularning xususiyatlari", [
            "Chiziqli algoritm — buyruqlar ketma-ket, birin-ketin, shartsiz bajarilgan algoritm turidir. Unda hech qanday shart tekshirilmaydi va hech qanday qadam takrorlanmaydi. Har bir buyruq aniq bir marta bajariladi.",
            "Chiziqli algoritmga misol: doira yuzasini hisoblash — avval radius qiymati kiritiladi, keyin S = pi * r * r formulasi bo'yicha hisob-kitob bajariladi, so'ngra natija chiqariladi. Bu uchta buyruq har doim shu tartibda bajariladi.",
            "Blok-sxemada chiziqli algoritm to'g'ri chiziq bo'ylab yuqoridan pastga joylashgan bloklar ketma-ketligi shaklida ifodalanadi. Hayotda ham ko'plab jarayonlar chiziqli: masalan, oshxonada ovqat pishirish retseptidagi qadamlar yoki laboratoriya ishining ketma-ket bosqichlari.",
        ]),
        ("Tarmoqlanuvchi algoritmlar va shart tushunchasi", [
            "Tarmoqlanuvchi algoritm — muayyan shartning bajarilishi yoki bajarilmasligiga qarab turli yo'nalishlar bo'yicha davom etadigan algoritmdir. U 'agar ... bo'lsa ... aks holda ...' mantiqiga asoslanadi.",
            "Shart — bu rost (true) yoki yolg'on (false) qiymat beruvchi mantiqiy ifodadir. Masalan: x > 0, a == b, y != 10. Shartning natijasiga qarab algoritm ikki yoki undan ortiq variantdan birini bajaradi.",
            "Tarmoqlanish to'liq (if-else: ikkala holat uchun amal bor) va to'liqsiz (faqat if: shart rost bo'lgandagina amal bajariladi, aks holda hech narsa qilinmaydi) bo'lishi mumkin. Ichma-ich (nested) tarmoqlanish esa bir nechta shartni ketma-ket tekshirish imkonini beradi.",
        ]),
        ("Takrorlanuvchi (siklik) algoritmlar", [
            "Takrorlanuvchi (siklik) algoritm — muayyan buyruqlar guruhini bir nechta marta qayta bajaruvchi algoritmdir. Takrorlanadigan qism 'sikl tanasi' deb ataladi. Sikl shart bajarilmaguncha yoki belgilangan marta takrorlanadi.",
            "Siklning uchta asosiy turi mavjud: sharti oldin tekshiriladigan sikl (while) — avval shart tekshiriladi, rost bo'lsa bajariladi; sharti keyin tekshiriladigan sikl (do-while) — avval bir marta bajariladi, keyin shart tekshiriladi; parametrli sikl (for) — takrorlanishlar soni oldindan ma'lum.",
            "Sikllardan chiqish sharti to'g'ri belgilanmasa, cheksiz sikl (infinite loop) hosil bo'ladi — dastur to'xtamaydi. Shu sababli sikl o'zgaruvchisining har qadamda o'zgarishini va chiqish shartiga yaqinlashishini ta'minlash zarur.",
        ]),
        ("Algoritmlarni blok-sxema va dastur shaklida ifodalash", [
            "Blok-sxemada tarmoqlanish romb (olmos) shakli bilan ifodalanadi; undan ikkita chiqish yo'nalishi bor: 'ha' (rost) va 'yo'q' (yolg'on). Har bir yo'nalishda mos buyruqlar bloki joylashadi, so'ngra oqim yana birlashadi.",
            "Sikl blok-sxemada shart rombi va orqaga qaytuvchi strelka ko'rinishida ifodalanadi. While sikli uchun strelka rombdan oldin qaytadi, do-while uchun esa sikl tanasidan keyin rombga boradi.",
            "C++ dasturida chiziqli algoritmlar ketma-ket operatorlar sifatida, tarmoqlanish if-else yoki switch operatorlari bilan, sikllar esa for, while va do-while operatorlari bilan yoziladi. Blok-sxemadan dasturga o'tish aniq qoidalarga asoslanadi va mashq qilish orqali o'zlashtiriladi.",
        ]),
    ],
})



TOPICS.append({
    "num": 16,
    "title": "C++ tilida ma'lumotlar va ularning turlari",
    "reja": [
        "Ma'lumot turi (data type) tushunchasi va ahamiyati.",
        "Butun sonli turlar (integer types).",
        "Haqiqiy sonli turlar (floating-point types).",
        "Belgi, mantiqiy va boshqa turlar.",
    ],
    "sections": [
        ("Ma'lumot turi (data type) tushunchasi va ahamiyati", [
            "Ma'lumot turi — o'zgaruvchining xotirada qancha joy egallashi, qanday qiymatlar qabul qilishi va ustida qanday amallar bajarilishi mumkinligini belgilaydigan xarakteristikadir. C++ da har bir o'zgaruvchi e'lon qilinganda uning turi ko'rsatilishi shart.",
            "Ma'lumot turlari kompilyatorga qancha xotira ajratish kerakligini bildiradi. Noto'g'ri tur tanlash xotira isrofiga yoki qiymatning sig'masligiga olib kelishi mumkin. Shu sababli dasturchi masalaga mos turni tanlashni bilishi kerak.",
            "C++ da ma'lumot turlari asosiy (fundamental) va murakkab (compound) turlarga bo'linadi. Asosiy turlarga butun sonlar, haqiqiy sonlar, belgilar va mantiqiy tur kiradi. Murakkab turlarga massivlar, strukturalar, ko'rsatkichlar va klasslar kiradi.",
        ]),
        ("Butun sonli turlar (integer types)", [
            "Butun sonli turlar kasr qismisiz sonlarni saqlash uchun ishlatiladi. Asosiy butun sonli turlar: short (2 bayt, -32768 dan 32767 gacha), int (odatda 4 bayt, taxminan -2 milliard dan +2 milliard gacha), long (4 yoki 8 bayt) va long long (8 bayt).",
            "Har bir butun sonli turning unsigned (ishorasiz, faqat musbat) varianti ham bor. Masalan, unsigned int 0 dan 4294967295 gacha qiymat oladi. Ishorasiz turlar manfiy qiymat kerak bo'lmaganda, masalan, yosh yoki soni kabi hollarda ishlatiladi.",
            "Tur o'lchami platformaga bog'liq bo'lishi mumkin, lekin C++ standarti minimal o'lchamlarni kafolatlaydi: short >= 2 bayt, int >= 2 bayt (amalda 4), long >= 4 bayt, long long >= 8 bayt. sizeof operatori orqali turning aniq o'lchamini bilish mumkin.",
        ]),
        ("Haqiqiy sonli turlar (floating-point types)", [
            "Haqiqiy sonli turlar kasr qismga ega sonlarni ifodalash uchun ishlatiladi. C++ da uchta haqiqiy tur bor: float (4 bayt, 6-7 ta aniq raqam), double (8 bayt, 15-16 ta aniq raqam) va long double (odatda 12 yoki 16 bayt, yanada yuqori aniqlik).",
            "Haqiqiy sonlar kompyuterda suzuvchi nuqta (floating-point) formatida saqlanadi — bu ilmiy yozuv (mantissa * 10^daraja) ga o'xshash. Bu juda katta va juda kichik sonlarni ifodalash imkonini beradi, lekin cheklangan aniqlik mavjud.",
            "Muhim eslatma: haqiqiy sonlar bilan ishlashda yaxlitlash xatolari yuz berishi mumkin. Masalan, 0.1 + 0.2 aynan 0.3 ga teng bo'lmasligi mumkin. Shu sababli haqiqiy sonlarni taqqoslashda aniq tenglik (==) o'rniga taqribiy tenglik (fark < epsilon) ishlatiladi.",
        ]),
        ("Belgi, mantiqiy va boshqa turlar", [
            "char turi bitta belgini saqlash uchun ishlatiladi va 1 bayt joy egallaydi. Belgilar ASCII kodlash jadvali asosida raqamli ko'rinishda saqlanadi: masalan, 'A' = 65, 'a' = 97, '0' = 48. char turi aslida kichik butun son sifatida ham ishlatilishi mumkin.",
            "bool (mantiqiy) turi faqat ikkita qiymat qabul qiladi: true (rost, 1 ga teng) va false (yolg'on, 0 ga teng). U shart ifodalari va mantiqiy operatsiyalarda keng ishlatiladi. Xotirada 1 bayt joy egallaydi.",
            "void turi — qiymat yo'qligini bildiradi. U funksiyalar uchun ishlatiladi: void qaytarmaydigan funksiya hech qanday qiymat bermaydi. wchar_t turi keng belgilar (Unicode) uchun ishlatiladi va 2 yoki 4 bayt joy egallaydi. C++11 dan boshlab auto kalit so'zi turni avtomatik aniqlash imkonini beradi.",
        ]),
    ],
})

TOPICS.append({
    "num": 17,
    "title": "O'zgarmaslar. O'zgaruvchilar. C++ tilida ifodalar. Operatorlar",
    "reja": [
        "O'zgaruvchilarni e'lon qilish va qiymat berish.",
        "O'zgarmaslar (konstantalar) va ularning turlari.",
        "C++ tilida ifodalar va ularni hisoblash tartibi.",
        "Arifmetik, tayinlash va tur o'zgartirish operatorlari.",
    ],
    "sections": [
        ("O'zgaruvchilarni e'lon qilish va qiymat berish", [
            "O'zgaruvchi — xotiradagi nomlangan soha bo'lib, u muayyan turdagi qiymatni saqlaydi va dastur davomida o'zgarishi mumkin. O'zgaruvchi e'lon qilish uchun tur va nom ko'rsatiladi: int x; double salary; char grade; Bu xotiradan joy ajratadi.",
            "O'zgaruvchiga boshlang'ich qiymat berish (initsializatsiya) e'lon bilan birga amalga oshirilishi tavsiya etiladi: int count = 0; double pi = 3.14159; Initsializatsiya qilinmagan o'zgaruvchi noma'lum (axlat) qiymatga ega bo'lishi mumkin, bu xatolarga olib keladi.",
            "Bir satrda bir nechta o'zgaruvchini e'lon qilish mumkin: int a, b, c; yoki int a = 1, b = 2; O'zgaruvchilar dasturda ishlatilishidan oldin e'lon qilinishi shart. C++ da o'zgaruvchini dasturning istalgan joyida e'lon qilish mumkin (C dan farqli).",
        ]),
        ("O'zgarmaslar (konstantalar) va ularning turlari", [
            "O'zgarmas (konstanta) — dastur bajarilishi davomida o'zgarmaydigan qiymatdir. C++ da konstantalar ikki usulda yaratiladi: const kalit so'zi (const double PI = 3.14159265;) va #define direktivasi (#define MAX 100). const usuli afzalroq, chunki u tur xavfsizligini ta'minlaydi.",
            "Literal konstantalar — bu manba kodda to'g'ridan-to'g'ri yozilgan qiymatlar: butun son literallari (42, 0xFF, 077), haqiqiy son literallari (3.14, 2.5e3), belgi literallari ('A', '\\n') va satr literallari (\"Salom\"). Har bir literalning o'z turi bor.",
            "Konstantalardan foydalanish dasturni ishonchli va tushunarli qiladi. Agar biror qiymat dasturning ko'p joyida ishlatsangiz, uni konstanta sifatida bir joyda e'lon qilib, har joyda uning nomini ishlatish oson. Qiymatni o'zgartirish kerak bo'lsa, faqat bir joyda o'zgartiriladi.",
        ]),
        ("C++ tilida ifodalar va ularni hisoblash tartibi", [
            "Ifoda (expression) — operandlar va operatorlardan tashkil topgan konstruksiya bo'lib, hisoblansa natija beradi. Masalan: a + b * c; x > 0; i++; Bu ifodalarning har biri qiymat hosil qiladi.",
            "Ifodani hisoblash tartibi operatorlarning ustuvorligi (priority) va assotsiativligi bilan belgilanadi. Ko'paytirish va bo'lish qo'shish va ayirishdan oldin bajariladi. Teng ustuvorlikdagi operatorlar chapdan o'ngga (yoki o'ngdan chapga, masalan tayinlash) tartibida bajariladi.",
            "Qavslar () hisoblash tartibini o'zgartirish uchun ishlatiladi: (a + b) * c — avval qo'shiladi, keyin ko'paytiriladi. Murakkab ifodalarni qavslar yordamida aniq yozish xatolarni kamaytiradi va kodni tushunarliroq qiladi.",
        ]),
        ("Arifmetik, tayinlash va tur o'zgartirish operatorlari", [
            "Arifmetik operatorlar: + (qo'shish), - (ayirish), * (ko'paytirish), / (bo'lish), % (qoldiqli bo'lish). Muhim: butun sonlarni bo'lganda natija ham butun son bo'ladi (7/2 = 3, qoldiq tashlanadi). Haqiqiy natija uchun operandlardan biri haqiqiy bo'lishi kerak (7.0/2 = 3.5).",
            "Tayinlash operatori (=) o'ngdagi ifoda qiymatini chapdagi o'zgaruvchiga beradi. Qisqartirilgan tayinlash operatorlari: +=, -=, *=, /=, %= (masalan, x += 5 bu x = x + 5 bilan teng). Inkrement (++) va dekrement (--) operatorlari qiymatni 1 ga oshiradi yoki kamaytiradi.",
            "Tur o'zgartirish (type casting) — bir turdagi qiymatni boshqa turga aylantirish. Yashirin (implicit) o'zgartirish kompilyator tomonidan avtomatik bajariladi (masalan, int dan double ga). Aniq (explicit) o'zgartirish dasturchi tomonidan ko'rsatiladi: (double)x yoki static_cast<double>(x).",
        ]),
    ],
})

TOPICS.append({
    "num": 18,
    "title": "C++ tilining asosiy elementlari",
    "reja": [
        "Dastur tuzilishi va main funksiya.",
        "Kiritish va chiqarish oqimlari (cin, cout).",
        "Standart kutubxonalar va ularni ulash.",
        "Nomlash sohalari (namespace) va ko'rinish sohasi (scope).",
    ],
    "sections": [
        ("Dastur tuzilishi va main funksiya", [
            "Har bir C++ dasturining bajarilishi main() funksiyasidan boshlanadi. Bu funksiya dasturda aynan bitta bo'lishi shart. Uning ikki turi mavjud: int main() va int main(int argc, char* argv[]) — ikkinchisi buyruq satridan argumentlar olish uchun ishlatiladi.",
            "main funksiyasi int turida qiymat qaytaradi: return 0 — dastur muvaffaqiyatli tugadi, noldan farqli qiymat esa xatolikni bildiradi. Funksiya tanasi figurali qavslar ichida yoziladi va bir nechta operatorlardan iborat bo'ladi.",
            "Dastur odatda quyidagi tartibda tuziladi: 1) preprosessor direktivalar (#include); 2) using namespace std; e'loni; 3) global o'zgaruvchi va funksiya prototipalari (kerak bo'lsa); 4) main() funksiya; 5) boshqa funksiyalarning ta'riflari. Bu tartib dasturni o'qishni osonlashtiradi.",
        ]),
        ("Kiritish va chiqarish oqimlari (cin, cout)", [
            "C++ da standart kiritish-chiqarish <iostream> kutubxonasi orqali amalga oshiriladi. cout oqimi va << (chiqarish) operatori ma'lumotni ekranga chiqaradi: cout << \"Natija: \" << x << endl; Bu bir nechta elementni ketma-ket chiqarish imkonini beradi.",
            "cin oqimi va >> (kiritish) operatori klaviaturadan ma'lumot o'qiydi: cin >> x; Bir nechta qiymat ketma-ket o'qilishi mumkin: cin >> a >> b; cin bo'sh joylarni ajratuvchi sifatida qabul qiladi va mos turdagi qiymatni o'zgaruvchiga yozadi.",
            "endl manipulyatori yangi satrga o'tkazadi va bufer tozalaydi. '\\n' belgisi ham yangi satr hosil qiladi, lekin buferni tozalamaydi (tezroq ishlaydi). setw, setprecision kabi manipulyatorlar (<iomanip> kutubxonasi) chiqarishni formatlash uchun ishlatiladi.",
        ]),
        ("Standart kutubxonalar va ularni ulash", [
            "C++ standart kutubxonasi (STL — Standard Template Library) keng imkoniyatlar beradi: konteynerlar (vector, list, map), algoritmlar (sort, find, count), kiritish-chiqarish, matematika va boshqalar. Ular #include direktivasi orqali ulanadi.",
            "Eng ko'p ishlatiladigan kutubxonalar: <iostream> — kiritish-chiqarish; <cmath> — matematik funksiyalar (sqrt, pow, sin, cos); <string> — satrlar bilan ishlash; <vector> — dinamik massiv; <algorithm> — saralash va qidiruv; <fstream> — fayllar bilan ishlash.",
            "Kutubxonani ulash: #include <kutubxona_nomi> — standart kutubxonalar uchun burchakli qavslar, #include \"mening_faylim.h\" — foydalanuvchi fayllari uchun qo'shtirnoq ishlatiladi. Bu dasturning imkoniyatlarini kengaytiradi va tayyor yechimlardan foydalanish imkonini beradi.",
        ]),
        ("Nomlash sohalari (namespace) va ko'rinish sohasi (scope)", [
            "Namespace — nom to'qnashuvlarini oldini olish uchun mo'ljallangan mexanizmdir. std nomlar sohasi C++ standart kutubxonasidagi barcha elementlarni o'z ichiga oladi. using namespace std; e'loni har safar std:: yozmaslik uchun qulaylik yaratadi.",
            "O'zgaruvchining ko'rinish sohasi (scope) — u foydalanish mumkin bo'lgan dastur qismidir. Lokal o'zgaruvchilar faqat e'lon qilingan blok (figurali qavslar) ichida ko'rinadi. Global o'zgaruvchilar esa dasturning barcha qismida foydalanish mumkin.",
            "Bir xil nomli lokal va global o'zgaruvchi bo'lganda, lokal o'zgaruvchi ustunlik qiladi. Globaliga murojaat uchun :: (scope resolution) operatori ishlatiladi. O'zgaruvchining ko'rinish sohasini to'g'ri boshqarish xatolarni kamaytiradi va dasturni xavfsizroq qiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 19,
    "title": "C++ tilida operatorlar",
    "reja": [
        "Operatorlar tushunchasi va tasnifi.",
        "Taqqoslash va mantiqiy operatorlar.",
        "Bitli (bit) operatorlar.",
        "Operatorlar ustuvorligi va assotsiativligi.",
    ],
    "sections": [
        ("Operatorlar tushunchasi va tasnifi", [
            "Operator — bu operandlar ustida muayyan amalni bajaruvchi belgi yoki belgilar guruhidir. C++ da operatorlar arifmetik (+, -, *, /, %), tayinlash (=, +=, -=), taqqoslash (==, !=, <, >), mantiqiy (&&, ||, !), bitli (&, |, ^, ~, <<, >>) va boshqa turlarga bo'linadi.",
            "Operatorlar operandlar soniga ko'ra unar (bitta operand: ++x, !a), binar (ikkita operand: a + b, x == y) va ternar (uchta operand: shart ? ifoda1 : ifoda2) bo'ladi. Ternar operator if-else ning qisqa shakli sifatida qulay.",
            "C++ da operatorlarni qayta yuklash (operator overloading) imkoniyati mavjud — bu foydalanuvchi tomonidan yaratilgan turlar (klasslar) uchun standart operatorlarga yangi ma'no berish imkonini beradi. Masalan, ikkita vektor obyektini + operatori bilan qo'shish mumkin.",
        ]),
        ("Taqqoslash va mantiqiy operatorlar", [
            "Taqqoslash operatorlari ikkita qiymatni solishtiradi va bool (true yoki false) natija qaytaradi. Operatorlar: == (teng), != (teng emas), < (kichik), > (katta), <= (kichik yoki teng), >= (katta yoki teng). Ular asosan shart ifodalarida (if, while) ishlatiladi.",
            "Mantiqiy operatorlar bool qiymatlar ustida amal bajaradi: && (VA — ikkala operand rost bo'lgandagina rost), || (YOKI — kamida biri rost bo'lsa rost), ! (INKOR — qiymatni teskarisiga o'zgartiradi). Ular murakkab shartlar tuzishda ishlatiladi.",
            "Muhim xususiyat — qisqa tutashuv (short-circuit evaluation): && da birinchi operand false bo'lsa, ikkinchisi tekshirilmaydi; || da birinchisi true bo'lsa, ikkinchisi tekshirilmaydi. Bu xususiyat samaradorlikni oshiradi va ba'zan xatolardan himoya qiladi.",
        ]),
        ("Bitli (bit) operatorlar", [
            "Bitli operatorlar sonlarning ikkilik (binary) ko'rinishidagi alohida bitlari ustida amal bajaradi. & (bitli VA), | (bitli YOKI), ^ (bitli ISTISNO YOKI — XOR), ~ (bitli INKOR), << (chapga surish), >> (o'ngga surish).",
            "Chapga surish (<<) sonni n bit chapga suradi, bu 2^n ga ko'paytirishga teng. O'ngga surish (>>) esa 2^n ga butun bo'lishga teng. Masalan: 5 << 1 = 10 (5ni 2ga ko'paytirish), 8 >> 2 = 2 (8ni 4ga bo'lish).",
            "Bitli operatorlar past darajali dasturlashda: bayroqlar (flags) bilan ishlash, maxsus registrlar boshqarish, ma'lumotlarni shifrlash va siqish algoritmlarida keng qo'llaniladi. Masalan, muayyan bitni yoqish: flags |= (1 << n); tekshirish: if (flags & (1 << n)).",
        ]),
        ("Operatorlar ustuvorligi va assotsiativligi", [
            "Operatorlar ustuvorligi (priority/precedence) — bir ifodada bir nechta operator bo'lganda qaysi biri birinchi bajarilishini belgilaydi. Masalan, * va / ning ustuvorligi + va - dan yuqori: 2 + 3 * 4 = 14 (avval ko'paytirish). Qavslar ustuvorlikni o'zgartiradi.",
            "C++ da 15 dan ortiq ustuvorlik darajasi bor. Eng yuqori: qavslar (), massiv indeksi []; o'rtacha: arifmetik, taqqoslash, mantiqiy; eng past: tayinlash (=) va vergul (,) operatorlari. To'liq jadvalni eslab qolish shart emas, lekin asosiy tartibni bilish muhim.",
            "Assotsiativlik — bir xil ustuvorlikdagi operatorlar qaysi tartibda bajarilishini ko'rsatadi. Ko'pchilik operatorlar chapdan o'ngga (left-to-right) assotsiativdir: a - b - c = (a-b) - c. Tayinlash va unar operatorlar o'ngdan chapga: a = b = c bu a = (b = c) degan ma'no.",
        ]),
    ],
})

TOPICS.append({
    "num": 20,
    "title": "C++ tilida tarmoqlanuvchi jarayonlarga dasturlar tuzish",
    "reja": [
        "if operatori va uning sintaksisi.",
        "if-else va ichma-ich if tuzilmalari.",
        "switch-case operatori.",
        "Tarmoqlanuvchi jarayonlarga amaliy misollar.",
    ],
    "sections": [
        ("if operatori va uning sintaksisi", [
            "if operatori shartli bajarilishni ta'minlaydi: agar shart rost bo'lsa, blok ichidagi buyruqlar bajariladi; aks holda o'tkazib yuboriladi. Sintaksisi: if (shart) { buyruqlar; } — figurali qavslar ichida bir yoki bir nechta buyruq yoziladi.",
            "Shart qismi doimo qavslar ichida yoziladi va mantiqiy ifoda bo'lishi kerak. C++ da noldan farqli har qanday qiymat rost (true), nol esa yolg'on (false) hisoblanadi. Masalan: if (x) — x nolga teng bo'lmasa bajariladi.",
            "Misol: sonning musbatligini tekshirish. if (x > 0) { cout << \"Musbat\" << endl; } — bu dastur x musbat bo'lgandagina xabar chiqaradi. Bitta buyruq bo'lsa qavslarni tushirish mumkin, lekin har doim qavs ishlatish xatolardan saqlaydi.",
        ]),
        ("if-else va ichma-ich if tuzilmalari", [
            "if-else tuzilmasi ikkita variant uchun mo'ljallangan: shart rost bo'lsa birinchi blok, yolg'on bo'lsa ikkinchi blok bajariladi. Sintaksisi: if (shart) { blok1; } else { blok2; } — bu to'liq tarmoqlanish.",
            "Ichma-ich if (nested if) — if bloki ichida yana if operatori joylashishi. else-if zanjiri bir nechta shartni ketma-ket tekshiradi: if (shart1) {...} else if (shart2) {...} else if (shart3) {...} else {...}. Birinchi rost shart topilsa, tegishli blok bajariladi va qolganlari o'tkazib yuboriladi.",
            "Misol: baho tizimi. if (ball >= 86) cout << \"A'lo\"; else if (ball >= 71) cout << \"Yaxshi\"; else if (ball >= 56) cout << \"Qoniqarli\"; else cout << \"Qoniqarsiz\"; — ball qiymatiga qarab natija chiqaradi.",
        ]),
        ("switch-case operatori", [
            "switch operatori bitta ifodaning qiymatiga qarab ko'p variantdan birini tanlash uchun ishlatiladi. U ko'p else-if zanjiri o'rniga qulayroq variant. Sintaksisi: switch (ifoda) { case qiymat1: buyruqlar; break; case qiymat2: buyruqlar; break; default: buyruqlar; }",
            "switch ifodasi butun yoki belgi (char) turida bo'lishi kerak. Har bir case yorlig'i aniq bir qiymatni ifodalaydi. break operatori kerakli blokdan keyin switch dan chiqish uchun ishlatiladi — u bo'lmasa keyingi case ga o'tib ketadi (fall-through).",
            "default bo'limi hech bir case mos kelmagan holda bajariladi va ixtiyoriy. Misol: hafta kunlari. switch(kun) { case 1: cout << \"Dushanba\"; break; case 2: cout << \"Seshanba\"; break; ... default: cout << \"Noto'g'ri\"; }",
        ]),
        ("Tarmoqlanuvchi jarayonlarga amaliy misollar", [
            "Misol 1: Ikki sondan kattasini topish. int a, b; cin >> a >> b; if (a > b) cout << a; else cout << b; — bu sodda tarmoqlanish misoli bo'lib, ikkita qiymatni solishtirish va kattasini aniqlash vazifasini hal qiladi.",
            "Misol 2: Kvadrat tenglama yechish (ax^2 + bx + c = 0). Avval diskriminant D = b*b - 4*a*c hisoblanadi. if (D > 0) — ikkita ildiz; else if (D == 0) — bitta ildiz; else — haqiqiy ildiz yo'q. Har bir holat uchun mos formula ishlatiladi.",
            "Misol 3: Kalkulyator (switch bilan). Foydalanuvchidan ikki son va amal belgisi (+, -, *, /) so'raladi. switch(amal) { case '+': natija = a+b; break; case '-': natija = a-b; break; ... } — amal belgisiga qarab tegishli hisob-kitob bajariladi.",
        ]),
    ],
})



TOPICS.append({
    "num": 21,
    "title": "Tarmoqlanuvchi jarayonlarga dasturlar tuzish",
    "reja": [
        "Murakkab shartli ifodalar tuzish.",
        "Tarmoqlanishning amaliy qo'llanilishi.",
        "Ko'p variantli tanlov masalalari.",
        "Tarmoqlanuvchi dasturlardagi tipik xatolar va ularni bartaraf etish.",
    ],
    "sections": [
        ("Murakkab shartli ifodalar tuzish", [
            "Murakkab shart — bir nechta oddiy shartlarning mantiqiy operatorlar (&&, ||, !) yordamida birlashtirilgan ifodasidir. Masalan: if (x > 0 && x < 100) — x 0 va 100 orasida ekanligini tekshiradi. Bu ikki shartning bir vaqtda bajarilishini talab qiladi.",
            "Mantiqiy operatorlarni to'g'ri qo'llash muhim: && — ikkala shart ham rost bo'lishi kerak; || — kamida bittasi rost bo'lsa yetarli; ! — shartni teskarisiga o'giradi. Qavslar murakkab shartlarda tartibni aniq qiladi: if ((a > b && c < d) || e == 0).",
            "Amaliy masala: yil kabisa (leap year) ekanligini aniqlash. Shart: (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0). Bu murakkab mantiqiy ifoda uchta shartni birlashtiradi va to'g'ri natija beradi.",
        ]),
        ("Tarmoqlanishning amaliy qo'llanilishi", [
            "Tarmoqlanish haqiqiy dasturlarda keng qo'llaniladi: foydalanuvchi kiritgan ma'lumotni tekshirish (validatsiya), xatoliklarni qayta ishlash, turli vaziyatlarga turli javob berish va ruxsatlarni boshqarish bunga misol bo'ladi.",
            "Ma'lumotlarni tekshirish misoli: foydalanuvchi yoshini kiritganda, if (yosh < 0 || yosh > 150) { cout << \"Xato!\"; } else { ... } — bu noto'g'ri qiymatlarni filtrlab, dasturning barqaror ishlashini ta'minlaydi.",
            "Menyu tizimi misoli: dastur foydalanuvchiga bir nechta tanlov taklif qiladi (1-Hisoblash, 2-Grafik, 3-Chiqish), keyin switch yoki if-else zanjiri orqali mos amal bajariladi. Bu interaktiv dasturlarning asosiy tuzilmasi.",
        ]),
        ("Ko'p variantli tanlov masalalari", [
            "Ba'zi masalalar ikki emas, balki ko'plab variantlardan birini tanlashni talab qiladi. Masalan, oy raqamiga qarab oy nomini chiqarish — 12 ta variant. Buni else-if zanjiri yoki switch operatori bilan yechish mumkin.",
            "enum (sanash) turlari ko'p variantli tanlashda qulaylik yaratadi: enum Ranglar {QIZIL, YASHIL, KO'K}; switch(rang) { case QIZIL: ...; case YASHIL: ...; } — bu kodni o'qishni osonlashtiradi va xatolarni kamaytiradi.",
            "Ko'p variantli masalaga yana bir misol: transport vositasining turini aniqlash. switch(tur) { case 'A': cout << \"Avtomobil\"; break; case 'V': cout << \"Velosiped\"; break; ... } — foydalanuvchi kiritgan belgiga mos nomni chiqaradi.",
        ]),
        ("Tarmoqlanuvchi dasturlardagi tipik xatolar va ularni bartaraf etish", [
            "Eng keng tarqalgan xato — taqqoslash (==) o'rniga tayinlash (=) operatorini ishlatish: if (x = 5) har doim rost bo'ladi, chunki x ga 5 qiymati beriladi. To'g'ri variant: if (x == 5). Ba'zi kompilyatorlar bu haqda ogohlantiradi.",
            "Yana bir tipik xato — else ning noto'g'ri if ga birikishi (dangling else). Ichma-ich if-else da else eng yaqin if ga tegishli. Noaniqlikni oldini olish uchun har doim figurali qavslar {} ishlatish kerak, hatto bitta operator bo'lsa ham.",
            "Switch da break operatorini unutish — keyingi case blokiga 'tushib ketish' (fall-through) xatosiga olib keladi. Shuningdek, butun diapazondagi qiymatlarni tekshirishda switch o'rniga if-else ishlatish to'g'riroq, chunki switch faqat alohida qiymatlarni tekshiradi.",
        ]),
    ],
})

TOPICS.append({
    "num": 22,
    "title": "Shartli o'tish operatori. U yordamida dasturlar tuzish",
    "reja": [
        "Shartli (ternar) operator va uning sintaksisi.",
        "Shartli operatorning if-else bilan taqqoslanishi.",
        "Shartli operator yordamida amaliy masalalar yechish.",
        "Ichma-ich shartli operatorlar va ularning qo'llanilishi.",
    ],
    "sections": [
        ("Shartli (ternar) operator va uning sintaksisi", [
            "Shartli operator (ternar operator) — C++ dagi yagona uchta operandli operatordir. Uning sintaksisi: shart ? ifoda1 : ifoda2. Agar shart rost bo'lsa ifoda1 ning qiymati qaytariladi, yolg'on bo'lsa — ifoda2 ning qiymati.",
            "Misol: int katta = (a > b) ? a : b; — bu a va b dan kattasini katta o'zgaruvchisiga beradi. Agar a > b rost bo'lsa, a olinadi, aks holda b olinadi. Bu bitta satrda if-else ning ishini bajaradi.",
            "Ternar operator ifoda sifatida ishlaydi, ya'ni u qiymat qaytaradi. Bu uni cout ichida, funksiya argumenti sifatida va boshqa ifodalar ichida to'g'ridan-to'g'ri ishlatish imkonini beradi: cout << (x >= 0 ? \"Musbat\" : \"Manfiy\");",
        ]),
        ("Shartli operatorning if-else bilan taqqoslanishi", [
            "Shartli operator if-else ning qisqa shakli bo'lib, sodda ikkita variantdan birini tanlash uchun qulay. Ammo u murakkab, ko'p qatorli amallar uchun mos emas — bunday hollarda if-else o'qishga qulayroq.",
            "if-else dan farqi: ternar operator ifoda hosil qiladi (uni o'zgaruvchiga berish mumkin), if-else esa operator (buyruq) hisoblanadi. Masalan: int min = (a < b) ? a : b; — buni if-else bilan yozish uchun avval o'zgaruvchini e'lon qilib, keyin if-else ichida qiymat berish kerak.",
            "Tavsiya: oddiy, bir qatorli tanlovlarda ternar operatorni, murakkab mantiq va ko'p qatorli bloklarda if-else ni ishlatish maqsadga muvofiq. Ternar operatorni haddan tashqari ko'p ishlatish kodni o'qishni qiyinlashtirishi mumkin.",
        ]),
        ("Shartli operator yordamida amaliy masalalar yechish", [
            "Masala 1: Sonning juft yoki toqligini aniqlash. string natija = (n % 2 == 0) ? \"Juft\" : \"Toq\"; cout << natija; — bu qisqa va aniq yechim.",
            "Masala 2: Absolyut qiymatni hisoblash. double abs_x = (x >= 0) ? x : -x; — manfiy bo'lsa ishorasini o'zgartiradi, musbat bo'lsa o'zini qaytaradi. Bu cmath kutubxonasidagi fabs() funksiyasi bilan bir xil natija beradi.",
            "Masala 3: Eng kichik uchta sonni topish. int min = (a < b) ? ((a < c) ? a : c) : ((b < c) ? b : c); — bu ichma-ich ternar operator bilan uchta qiymatdan eng kichigini aniqlaydi, lekin bunday murakkab ifoda o'rniga if-else o'qishga qulayroq bo'lishi mumkin.",
        ]),
        ("Ichma-ich shartli operatorlar va ularning qo'llanilishi", [
            "Ichma-ich (nested) ternar operator — bir ternar ifodaning ichida boshqa ternar ifoda joylashishidir. Bu ko'p variantli tanlov imkonini beradi: natija = (x > 0) ? \"Musbat\" : (x < 0) ? \"Manfiy\" : \"Nol\"; — uchta holatni bitta ifodada hal qiladi.",
            "Ichma-ich ternar operator baho berish uchun: char baho = (ball >= 86) ? 'A' : (ball >= 71) ? 'B' : (ball >= 56) ? 'C' : 'D'; — ball qiymatiga qarab mos harfli baho qaytaradi.",
            "Ammo ikki va undan ortiq darajali ichma-ich ternar operatorlar kodni o'qishni ancha qiyinlashtiradi. Dasturlash amaliyotida ikkitadan chuqurroq ichma-ich ternar operatordan qochish va bunday hollarda if-else yoki switch ishlatish tavsiya etiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 23,
    "title": "C++ tilida takrorlanuvchi jarayonlarga dasturlar tuzish",
    "reja": [
        "while sikl operatori va uning ishlash prinsipi.",
        "do-while sikl operatori.",
        "for sikl operatori va uning tuzilishi.",
        "Sikllar yordamida amaliy masalalar yechish.",
    ],
    "sections": [
        ("while sikl operatori va uning ishlash prinsipi", [
            "while operatori sharti oldin tekshiriladigan sikl bo'lib, sikl tanasi shart rost bo'lgan holda takrorlanadi. Sintaksisi: while (shart) { buyruqlar; }. Shart har takrorlanish boshida tekshiriladi — yolg'on bo'lsa, sikl tanasi bajarilmaydi va sikl tugaydi.",
            "Agar shart dastlab yolg'on bo'lsa, while sikl tanasi bir marta ham bajarilmaydi. Bu uni do-while dan farqlaydi. while sikli takrorlanishlar soni oldindan noma'lum bo'lgan holatlarda qo'l keladi: masalan, foydalanuvchi to'g'ri qiymat kiritguncha takrorlash.",
            "Misol: 1 dan 10 gacha sonlar yig'indisini hisoblash. int i = 1, sum = 0; while (i <= 10) { sum += i; i++; } — har qadamda i qo'shiladi va oshiriladi. i 11 bo'lganda shart yolg'on bo'lib, sikl tugaydi. Natija: sum = 55.",
        ]),
        ("do-while sikl operatori", [
            "do-while operatori sharti keyin tekshiriladigan sikl bo'lib, sikl tanasi kamida bir marta bajariladi. Sintaksisi: do { buyruqlar; } while (shart); — avval buyruqlar bajariladi, keyin shart tekshiriladi. Shart rost bo'lsa yana takrorlanadi.",
            "do-while ning asosiy farqi — u kamida bir marta bajarilishni kafolatlaydi. Bu foydalanuvchidan kiritish so'raydigan dasturlarda qulay: avval so'rov chiqariladi, keyin kiritilgan qiymat tekshiriladi va noto'g'ri bo'lsa qayta so'raladi.",
            "Misol: musbat son kiritishni so'rash. int n; do { cout << \"Musbat son kiriting: \"; cin >> n; } while (n <= 0); — foydalanuvchi musbat son kiritguncha takrorlanadi. Birinchi marta ham albatta so'rov chiqadi.",
        ]),
        ("for sikl operatori va uning tuzilishi", [
            "for operatori takrorlanishlar soni oldindan ma'lum bo'lganda eng qulay. Sintaksisi: for (boshlang'ich; shart; qadam) { buyruqlar; }. boshlang'ich — sikl o'zgaruvchisini e'lon va initsializatsiya; shart — har takrorlanish boshida tekshiriladi; qadam — har takrorlanish oxirida bajariladi.",
            "Misol: 1 dan n gacha sonlarning faktorialini hisoblash. long fact = 1; for (int i = 1; i <= n; i++) { fact *= i; } — i 1 dan n gacha oshadi, har qadamda fact ga i ko'paytiriladi.",
            "for siklining barcha uchta qismi ixtiyoriy: for (;;) cheksiz sikl hosil qiladi; boshlang'ichda bir nechta o'zgaruvchi e'lon qilish mumkin: for (int i=0, j=10; i<j; i++, j--). C++11 dan boshlab range-based for: for (int x : massiv) — massiv elementlarini birma-bir oladi.",
        ]),
        ("Sikllar yordamida amaliy masalalar yechish", [
            "Masala 1: n ta sonning o'rta arifmetigini hisoblash. double sum = 0; for (int i = 0; i < n; i++) { double x; cin >> x; sum += x; } cout << sum / n; — har bir kiritilgan son yig'indiga qo'shiladi, oxirida n ga bo'linadi.",
            "Masala 2: Sonning raqamlar yig'indisini topish. int s = 0; while (n > 0) { s += n % 10; n /= 10; } — har qadamda oxirgi raqam ajratiladi (n%10) va yig'indiga qo'shiladi, keyin son 10 ga bo'linadi. Son 0 bo'lganda sikl tugaydi.",
            "Masala 3: Fibonachchi ketma-ketligining dastlabki n hadini chiqarish. int a=0, b=1; for (int i=0; i<n; i++) { cout << a << \" \"; int temp = a + b; a = b; b = temp; } — har qadamda navbatdagi had hisoblanadi.",
        ]),
    ],
})

TOPICS.append({
    "num": 24,
    "title": "Takrorlanuvchi jarayonlarga dasturlar tuzish",
    "reja": [
        "Yig'indi va ko'paytma hisoblash sikllari.",
        "Ichma-ich (nested) sikllar.",
        "Siklda sanash va qidirish masalalari.",
        "Sikllar samaradorligi va optimallashtirish.",
    ],
    "sections": [
        ("Yig'indi va ko'paytma hisoblash sikllari", [
            "Ko'plab matematik masalalar ketma-ket elementlar yig'indisi yoki ko'paytmasini hisoblashni talab qiladi. Yig'indi uchun boshlang'ich qiymat 0, ko'paytma uchun esa 1 ga teng bo'lishi kerak. Har bir takrorlanishda navbatdagi element qo'shiladi yoki ko'paytiriladi.",
            "Misol: S = 1 + 1/2 + 1/3 + ... + 1/n. double S = 0; for (int i = 1; i <= n; i++) { S += 1.0 / i; } — e'tibor bering: 1.0 ishlatilgan, chunki 1/i butun bo'lish bo'lib 0 beradi. Haqiqiy natija uchun kamida bitta operand haqiqiy bo'lishi kerak.",
            "Ko'paytma misoli: n! (n faktorial) = 1 * 2 * 3 * ... * n. long P = 1; for (int i = 1; i <= n; i++) { P *= i; } — har qadamda P ga navbatdagi i ko'paytiriladi. Katta n qiymatlarda to'lib ketish (overflow) xavfi borligini yodda tutish kerak.",
        ]),
        ("Ichma-ich (nested) sikllar", [
            "Ichma-ich sikl — bir sikl tanasi ichida boshqa sikl joylashishi. Tashqi sikl har bir takrorlanishida ichki sikl to'liq bajariladi. Masalan, tashqi sikl 5 marta, ichki sikl 3 marta takrorlansa, jami 15 ta takrorlanish bo'ladi.",
            "Misol: ko'paytirish jadvali. for (int i=1; i<=9; i++) { for (int j=1; j<=9; j++) { cout << i*j << \"\\t\"; } cout << endl; } — tashqi sikl satrlarni, ichki sikl ustunlarni boshqaradi. Har bir satrda 9 ta ko'paytma chiqariladi.",
            "Ichma-ich sikllar ikki o'lchovli massivlar, matritsalar va jadvallar bilan ishlashda keng qo'llaniladi. Ammo uch va undan ortiq darajali ichma-ich sikllar dasturning tezligini sezilarli pasaytirishi mumkin, shuning uchun ularni ehtiyotkorlik bilan ishlatish kerak.",
        ]),
        ("Siklda sanash va qidirish masalalari", [
            "Sanash masalasi — berilgan to'plamda muayyan shartga javob beradigan elementlar sonini aniqlash. Misol: massivda nechta musbat son borligini hisoblash. int count = 0; for (int i=0; i<n; i++) { if (a[i] > 0) count++; } — har safar shart bajarilganda hisoblagich oshadi.",
            "Qidirish masalasi — berilgan shartga mos birinchi (yoki barcha) elementni topish. Misol: massivda birinchi manfiy elementni topish. for (int i=0; i<n; i++) { if (a[i] < 0) { cout << a[i]; break; } } — break topilgandan keyin siklni tugatadi.",
            "Eng katta va eng kichik elementni topish ham sikl yordamida bajariladi. int max = a[0]; for (int i=1; i<n; i++) { if (a[i] > max) max = a[i]; } — birinchi element boshlang'ich maximum sifatida olinadi, keyin har bir element u bilan taqqoslanadi.",
        ]),
        ("Sikllar samaradorligi va optimallashtirish", [
            "Siklning samaradorligi — u necha marta takrorlanishi va har bir takrorlanishda qancha ish bajarilishi bilan belgilanadi. O'zgarmas hisob-kitoblarni sikl tashqarisiga chiqarish, keraksiz takrorlanishlarni kamaytirish samaradorlikni oshiradi.",
            "Erta chiqish (early exit) strategiyasi: qidiruv masalalarida element topilgach break yordamida siklni tugatish. Bu o'rtacha holatda takrorlanishlar sonini ikki barobar kamaytirishi mumkin.",
            "Algoritmik murakkablik: bir martalik sikl O(n), ichma-ich ikki sikl O(n^2), uch sikl O(n^3) murakkablikka ega. n katta bo'lganda bu farq juda sezilarli. Masalan, n=1000 da O(n^2) = 1 000 000 amal, O(n^3) = 1 000 000 000 amal talab qiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 25,
    "title": "Sharti oldin beriladigan va sharti keyin beriladigan sikl operatorlari",
    "reja": [
        "while siklining batafsil tahlili va ishlash mexanizmi.",
        "do-while siklining batafsil tahlili va ishlash mexanizmi.",
        "while va do-while o'rtasidagi farqlar va tanlash mezonlari.",
        "Amaliy masalalar: while va do-while qo'llanilishi.",
    ],
    "sections": [
        ("while siklining batafsil tahlili va ishlash mexanizmi", [
            "while sikli 'sharti oldin tekshiriladigan sikl' (precondition loop) deb ataladi. Ishlash tartibi: 1) shart tekshiriladi; 2) agar rost — sikl tanasi bajariladi va 1-qadamga qaytiladi; 3) agar yolg'on — sikl tugaydi va keyingi operator bajariladi.",
            "while siklida uchta muhim element bo'lishi kerak: boshlang'ich qiymat (sikldan oldin o'zgaruvchi initsializatsiyasi), shart (sikl davom etish sharti) va o'zgarish (sikl tanasida o'zgaruvchining o'zgarishi). Ulardan birortasi yo'q bo'lsa, cheksiz sikl yoki xato yuzaga keladi.",
            "while sikli ko'pincha: fayl oxirigacha o'qish (while (!fin.eof())), foydalanuvchi 'Chiqish' tanlaguncha ishlash, yoki aniqlik (epsilon) ga erishilguncha hisoblash kabi oldindan necha marta takrorlanishi noma'lum holatlarda ishlatiladi.",
        ]),
        ("do-while siklining batafsil tahlili va ishlash mexanizmi", [
            "do-while sikli 'sharti keyin tekshiriladigan sikl' (postcondition loop) deb ataladi. Ishlash tartibi: 1) sikl tanasi bajariladi; 2) shart tekshiriladi; 3) agar rost — 1-qadamga qaytiladi; 4) agar yolg'on — sikl tugaydi.",
            "do-while ning muhim xususiyati — sikl tanasi kamida bir marta albatta bajariladi, shart faqat keyin tekshiriladi. Bu xususiyat avval biror amalni bajarib, keyin davom ettirish kerakligini tekshirish lozim bo'lgan holatlarda foydali.",
            "do-while sintaksisida while dan farqli ravishda oxirida nuqtali vergul (;) qo'yiladi: do { ... } while (shart); — bu nuqtali vergulni unutish kompilyatsiya xatosiga olib keladi. Bu kichik, lekin muhim sintaktik farq.",
        ]),
        ("while va do-while o'rtasidagi farqlar va tanlash mezonlari", [
            "Asosiy farq: while da shart OLDIN tekshiriladi — sikl tanasi 0 marta ham bajarilishi mumkin. do-while da shart KEYIN tekshiriladi — sikl tanasi kamida 1 marta bajariladi. Bu farq dastur mantiqiga ta'sir qiladi.",
            "while ni tanlash kerak: agar shartni avval tekshirish mantiqan to'g'ri bo'lsa. Masalan, ro'yxat bo'sh bo'lishi mumkin bo'lganda elementlarni qayta ishlash — bo'sh bo'lsa hech narsa qilinmasligi kerak.",
            "do-while ni tanlash kerak: agar biror amalni avval bajarib, keyin shart tekshirish kerak bo'lsa. Masalan, menyu ko'rsatish (avval menyu chiqadi, keyin tanlov tekshiriladi) yoki ma'lumot kiritishni so'rash (avval so'rov, keyin tekshirish) hollari.",
        ]),
        ("Amaliy masalalar: while va do-while qo'llanilishi", [
            "while misoli: Evklid algoritmi (EKUB topish). while (b != 0) { int temp = b; b = a % b; a = temp; } cout << a; — bu ikki sonning eng katta umumiy bo'luvchisini topadi. b nol bo'lganda sikl tugaydi va a da javob qoladi.",
            "do-while misoli: Sonning raqamlar sonini hisoblash. int count = 0; do { n /= 10; count++; } while (n != 0); — do-while bu yerda muhim, chunki n=0 bo'lganda ham javob 1 bo'lishi kerak (0 soni 1 ta raqamdan iborat). while ishlatilsa n=0 da natija 0 bo'lardi.",
            "Kombinatsiyalangan misol: Dastur foydalanuvchidan takrorlashni so'raydi. char javob; do { /* asosiy ish */ cout << \"Davom etasizmi? (h/y): \"; cin >> javob; } while (javob == 'h' || javob == 'H'); — foydalanuvchi 'h' kiritsa davom etadi.",
        ]),
    ],
})



TOPICS.append({
    "num": 26,
    "title": "Parametrli sikl operatorlari",
    "reja": [
        "for sikl operatorining tuzilishi va ishlash mexanizmi.",
        "Sikl parametrining turli variantlari.",
        "for siklida massivlar va ketma-ketliklar bilan ishlash.",
        "Amaliy masalalar: for sikli qo'llanilishi.",
    ],
    "sections": [
        ("for sikl operatorining tuzilishi va ishlash mexanizmi", [
            "for sikli uchta qismdan iborat: for (initsializatsiya; shart; qadam) { tana; }. Initsializatsiya faqat bir marta — siklning boshida bajariladi. Shart har takrorlanish boshida tekshiriladi. Qadam har takrorlanish oxirida bajariladi.",
            "Bajarilish tartibi: 1) initsializatsiya; 2) shart tekshiriladi — yolg'on bo'lsa sikl tugaydi; 3) shart rost — tana bajariladi; 4) qadam bajariladi; 5) 2-bosqichga qaytiladi. Bu tartibni tushunish sikl bilan ishlashning asosidir.",
            "for sikli while siklining maxsus, qulayroq shakli hisoblanadi. Quyidagi ikkita konstruksiya ekvivalent: for (int i=0; i<n; i++) {...} va int i=0; while (i<n) {...; i++;} — lekin for siklida barcha boshqaruv elementlari bir joyda yoziladi, bu esa o'qishni osonlashtiradi.",
        ]),
        ("Sikl parametrining turli variantlari", [
            "Sikl o'zgaruvchisi o'sishi shart emas — kamayishi ham mumkin: for (int i=10; i>=1; i--) — 10 dan 1 gacha kamayish tartibida. Qadam ixtiyoriy: for (int i=0; i<100; i+=5) — 5 tadan oshadi. Haqiqiy parametr ham mumkin: for (double x=0; x<=1; x+=0.1).",
            "for siklining bir nechta o'zgaruvchili ko'rinishi: for (int i=0, j=n-1; i<j; i++, j--) — ikkita o'zgaruvchi bir vaqtda boshqariladi. Bu massivning ikkala uchidan markazga qarab harakatlanishda foydali (masalan, massivni teskarisiga aylantirish).",
            "C++11 dagi range-based for: for (int elem : arr) { cout << elem; } — massiv yoki konteynerning har bir elementini navbatma-navbat oladi. Bu indeks bilan ishlash kerak bo'lmaganda kodni soddalashtiradi va xatolarni kamaytiradi.",
        ]),
        ("for siklida massivlar va ketma-ketliklar bilan ishlash", [
            "Massiv elementlarini o'qish: for (int i=0; i<n; i++) { cin >> a[i]; } — har bir element ketma-ket o'qiladi. Chiqarish: for (int i=0; i<n; i++) { cout << a[i] << \" \"; } — elementlar orasiga bo'sh joy qo'yib chiqariladi.",
            "Massiv elementlari yig'indisi: int sum=0; for (int i=0; i<n; i++) sum += a[i]; Eng katta element: int max=a[0]; for (int i=1; i<n; i++) if (a[i]>max) max=a[i]; — bu massivlar bilan ishlashning eng asosiy operatsiyalari.",
            "Ketma-ketliklar (seriyalar) bilan ishlash: for sikli yordamida arifmetik va geometrik progressiyalar, Fibonachchi ketma-ketligi, tub sonlar va boshqa matematik ketma-ketliklarni generatsiya qilish va ular ustida amallar bajarish mumkin.",
        ]),
        ("Amaliy masalalar: for sikli qo'llanilishi", [
            "Masala 1: Sonning tub yoki tub emasligini aniqlash. bool tub = true; for (int i=2; i*i<=n; i++) { if (n%i==0) { tub=false; break; } } — n ni 2 dan sqrt(n) gacha sonlarga bo'lib ko'ramiz, bo'linsa tub emas.",
            "Masala 2: Sonni ikkilik (binary) sanoq tizimiga aylantirish. string natija = \"\"; for (int temp=n; temp>0; temp/=2) { natija = char('0'+temp%2) + natija; } — har qadamda oxirgi bit ajratiladi va satr boshiga qo'shiladi.",
            "Masala 3: Yulduzchalardan uchburchak chizish. for (int i=1; i<=n; i++) { for (int j=1; j<=i; j++) cout << \"*\"; cout << endl; } — tashqi sikl satrlar, ichki sikl har satrdagi yulduzchalar sonini boshqaradi. Natija: 1-satrda 1 ta, 2-satrda 2 ta, ... n-satrda n ta yulduzcha.",
        ]),
    ],
})

TOPICS.append({
    "num": 27,
    "title": "Break, continue va goto operatorlari. Ular yordamida dasturlar tuzish",
    "reja": [
        "break operatori va uning qo'llanilishi.",
        "continue operatori va uning qo'llanilishi.",
        "goto operatori va uning xususiyatlari.",
        "break, continue va goto bilan amaliy misollar.",
    ],
    "sections": [
        ("break operatori va uning qo'llanilishi", [
            "break operatori sikl (for, while, do-while) yoki switch operatoridan zudlik bilan chiqish uchun ishlatiladi. U bajarilganda, eng yaqin (ichki) sikl yoki switch to'liq to'xtatiladi va undan keyingi operator bajariladi.",
            "break ko'pincha qidiruv masalalarida ishlatiladi: kerakli element topilgach siklni davom ettirishning keragi yo'q. Misol: massivda manfiy element borligini tekshirish. for (int i=0; i<n; i++) { if (a[i]<0) { bor=true; break; } } — birinchi manfiy topilsa chiqadi.",
            "Ichma-ich sikllarda break faqat ENG ICHKI siklni tugatadi. Tashqi sikldan chiqish uchun bayroq o'zgaruvchi (flag) yoki goto ishlatish kerak. Masalan: bool topildi=false; for (...) { for (...) { if (shart) { topildi=true; break; } } if (topildi) break; }",
        ]),
        ("continue operatori va uning qo'llanilishi", [
            "continue operatori joriy takrorlanishning qolgan qismini o'tkazib yuboradi va siklni KEYINGI takrorlanishdan davom ettiradi. for siklida continue dan keyin qadam qismi (i++) bajariladi; while/do-while da esa shart tekshiriladi.",
            "Misol: faqat musbat sonlarni yig'ish, manfiy sonlarni o'tkazib yuborish. for (int i=0; i<n; i++) { if (a[i] < 0) continue; sum += a[i]; } — manfiy element uchrasa, sum += qatori o'tkazib yuboriladi va keyingi i ga o'tiladi.",
            "continue murakkab sikl tanalarida foydali: agar biror shart bajarilmasa, qolgan ko'p qatorli kodni o'tkazish kerak bo'lganda, if ichiga joylashtirib qo'yish o'rniga boshida continue yozish kodni tekisroq (kamroq ichma-ich) qiladi.",
        ]),
        ("goto operatori va uning xususiyatlari", [
            "goto operatori dastur bajarilishini belgilangan yorliqqa (label) shartsiz o'tkazadi. Sintaksisi: goto yorliq_nomi; ... yorliq_nomi: operator; Yorliq — identifikator va undan keyin ikki nuqta (:) dan iborat.",
            "goto zamonaviy dasturlashda deyarli ishlatilmaydi, chunki u dastur mantiqini tushunishni qiyinlashtiradi va 'spagetti kod' (chalkash, o'qib bo'lmaydigan kod) ga olib keladi. Strukturali dasturlash tamoyillari goto siz ham istalgan algoritmni ifodalash mumkinligini isbotlaydi.",
            "goto ning yagona qabul qilingan qo'llanilishi — ichma-ich ko'p darajali sikllardan tashqariga bir vaqtda chiqish. Bu holda goto break + bayroqdan soddaroq bo'lishi mumkin: for (...) { for (...) { if (shart) goto tashqari; } } tashqari: ... Boshqa hollarda goto dan qochish kerak.",
        ]),
        ("break, continue va goto bilan amaliy misollar", [
            "Masala 1 (break): Foydalanuvchi 0 kiritguncha sonlar yig'indisini hisoblash. while (true) { cin >> x; if (x == 0) break; sum += x; } — cheksiz sikl ichida 0 kirsa break chiqaradi. Bu 'sentinel value' (signal qiymati) texnikasi deb ataladi.",
            "Masala 2 (continue): 1 dan 100 gacha sonlar ichida 3 ga bo'linmaydiganlarning yig'indisi. for (int i=1; i<=100; i++) { if (i%3==0) continue; sum += i; } — 3 ga bo'linadiganlar o'tkazib yuboriladi, qolganlari yig'iladi.",
            "Masala 3 (ichma-ich sikldan chiqish): Ikki o'lchovli massivda qiymatni qidirish. for (int i=0; i<m; i++) { for (int j=0; j<n; j++) { if (mat[i][j]==key) { cout<<i<<\",\"<<j; goto topildi; } } } topildi: ... — topilgach ikkala sikldan ham chiqiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 28,
    "title": "For, Break, continue va goto operatorlari",
    "reja": [
        "for siklining kengaytirilgan imkoniyatlari.",
        "break va continue ning for siklida ishlashi.",
        "Sikllarni boshqarish strategiyalari.",
        "Kompleks masalalar yechishda sikllarni boshqarish.",
    ],
    "sections": [
        ("for siklining kengaytirilgan imkoniyatlari", [
            "for siklining ixtiyoriy qismlarini tushirib qoldirish mumkin. for (;;) — barcha qismlar tushirilgan, bu cheksiz sikl hosil qiladi (break bilan boshqariladi). for (;shart;) — while ga ekvivalent. Bu moslashuvchanlik turli vaziyatlarda qo'l keladi.",
            "for siklida initsializatsiya qismida turli xil operatsiyalar bajarilishi mumkin: funksiya chaqirish, bir nechta o'zgaruvchini e'lon qilish. Qadam qismida ham bir nechta ifoda vergul bilan ajratilishi mumkin: for (int i=0, j=n; i<n; i++, j--). Bu ikki yo'nalishda harakatlanish uchun qulay.",
            "C++ 17 da if va switch operatorlarida ham initsializatsiya kiritish imkoniyati paydo bo'ldi. Bu bilan birga for sikli zamonaviy C++ da lambdalar va algoritmlar bilan integratsiyalashgan holda ishlatilishi ham keng tarqalgan.",
        ]),
        ("break va continue ning for siklida ishlashi", [
            "for siklida break bajarilganda: initsializatsiya, shart, qadam — barchasini qoldirib, sikldan to'liq chiqiladi. Keyingi satrdan for dan keyin yozilgan operator bajariladi.",
            "for siklida continue bajarilganda: sikl tanasining qolgan qismi o'tkazib yuboriladi, lekin QADAM qismi (masalan, i++) BAJARILADI va keyin shart qayta tekshiriladi. Bu while dagi continue dan farqli — while da qadam alohida yozilmaganligi uchun uni o'tkazib yuborish mumkin.",
            "Muhim farq misoli: for (int i=0; i<10; i++) { if (i==5) continue; cout<<i; } — 5 ni o'tkazib, 0123456789 dan 012346789 chiqaradi. i++ bajariladi. while da xuddi shunday qilsangiz, i++ ni continue dan oldin yozmasangiz cheksiz sikl bo'ladi.",
        ]),
        ("Sikllarni boshqarish strategiyalari", [
            "Sentinel (qorovul) qiymati strategiyasi: sikl maxsus qiymat kiritilguncha davom etadi. while (true) { cin>>x; if (x==sentinel) break; process(x); } — bu ma'lumotlar soni oldindan noma'lum bo'lganda ishlatiladi.",
            "Bayroq (flag) o'zgaruvchi strategiyasi: shart bajarilganini belgilab, sikldan keyin tekshirish. bool topildi=false; for (...) { if (shart) { topildi=true; break; } } if (topildi) { ... } — bu break dan keyingi harakatni boshqarish uchun foydali.",
            "Hisoblagich strategiyasi: sikl ichida shartga mos elementlar soni hisoblanadi. int count=0; for (...) { if (shart) count++; } — sikldan keyin count qiymati orqali xulosalar chiqariladi. Bu uchta strategiya sikl dasturlarning ko'pchiligini qamrab oladi.",
        ]),
        ("Kompleks masalalar yechishda sikllarni boshqarish", [
            "Masala 1: Matritsada saddle point (egar nuqta) topish. Tashqi sikl satrlarni ko'radi, ichki sikl har satrda minimumni topadi, keyin bu minimum o'z ustunida maksimum ekanligini tekshiradi. break va bayroqlar birgalikda ishlatiladi.",
            "Masala 2: Oddiy sonlarni Eratosfen elagi usuli bilan topish. Massivda barcha elementlar true, keyin 2 dan boshlab har bir sonning karralilari false qilinadi: for (int i=2; i*i<=n; i++) { if (sieve[i]) for (int j=i*i; j<=n; j+=i) sieve[j]=false; } — ichma-ich sikl va continue qo'llaniladi.",
            "Masala 3: Matn ichidan so'zlarni ajratish. for siklida belgilar ko'riladi, bo'sh joy uchrasa continue, aks holda so'z boshlanadi va keyingi bo'sh joygacha yig'iladi. Bu parsing (tahlil qilish) masalalarining oddiy namunasi bo'lib, break va continue birgalikda ishlatiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 29,
    "title": "Mustaqil ta'lim: C++ tilida massivlar",
    "reja": [
        "Massiv tushunchasi va e'lon qilish.",
        "Massiv elementlariga murojaat va ular ustida amallar.",
        "Massivlarni funksiyalarga uzatish.",
        "Massivlar bilan ishlashda xatolar va ularni oldini olish.",
    ],
    "sections": [
        ("Massiv tushunchasi va e'lon qilish", [
            "Massiv — bir xil turdagi elementlar to'plamini yagona nom ostida saqlaydigan ma'lumotlar tuzilmasidir. Har bir element indeks (tartib raqami) orqali aniqlanadi. C++ da massiv indekslari 0 dan boshlanadi: birinchi element a[0], ikkinchisi a[1] va hokazo.",
            "Massivni e'lon qilish: tur nomi[o'lcham]; Masalan: int a[10]; — 10 ta butun sondan iborat massiv. O'lcham konstanta ifoda bo'lishi kerak (standart C++ da). Initsializatsiya: int a[5] = {1, 2, 3, 4, 5}; yoki int a[] = {1, 2, 3}; — o'lcham avtomatik aniqlanadi.",
            "Massiv e'lon qilinganda, kompilyator xotirada ketma-ket joylashgan, ko'rsatilgan o'lchamdagi joyni ajratadi. Masalan, int a[5] uchun 5*4=20 bayt ajratiladi (int = 4 bayt). Massiv elementlari xotirada bir-birining yonida joylashadi, bu esa ularga tez murojaat qilish imkonini beradi.",
        ]),
        ("Massiv elementlariga murojaat va ular ustida amallar", [
            "Massiv elementiga murojaat uchun indeks operatori [] ishlatiladi: a[i] — i-indeksli element. Indeks 0 dan (o'lcham-1) gacha bo'lishi kerak. C++ indeks chegarasini tekshirmaydi — chegaradan chiqish dasturning noto'g'ri ishlashiga yoki xatoga olib keladi.",
            "Massivni to'ldirish: for (int i=0; i<n; i++) cin >> a[i]; Chiqarish: for (int i=0; i<n; i++) cout << a[i] << \" \"; Yig'indi: int s=0; for (int i=0; i<n; i++) s += a[i]; Bu eng asosiy operatsiyalardir.",
            "Massiv elementlarini saralash (sorting) eng keng tarqalgan amallardan biri. Oddiy usul — tanlash orqali saralash: har qadamda minimum topiladi va joriy element bilan almashtiriladi. STL da tayyor sort funksiyasi bor: sort(a, a+n); — bu tezroq va qulay.",
        ]),
        ("Massivlarni funksiyalarga uzatish", [
            "C++ da massiv funksiyaga ko'rsatkich sifatida uzatiladi — nusxasi yaratilmaydi. Funksiya prototipi: void print(int arr[], int size); yoki void print(int* arr, int size); — o'lchamni ham alohida parametr sifatida uzatish kerak, chunki funksiya massiv o'lchamini bilmaydi.",
            "Funksiya ichida massiv elementlarini o'zgartirish asl massivga ta'sir qiladi (chunki manzil uzatilgan). Bu xususiyat massivni qayta ishlash uchun qulay, lekin ehtiyotkorlik talab qiladi — istamagan o'zgarishlar yuzaga kelishi mumkin.",
            "const kalit so'zi orqali massivni himoyalash mumkin: void print(const int arr[], int size); — bu funksiya ichida massiv elementlarini o'zgartirishga ruxsat bermaydi. Bu xavfsiz dasturlash tamoyiliga mos keladi.",
        ]),
        ("Massivlar bilan ishlashda xatolar va ularni oldini olish", [
            "Eng keng tarqalgan xato — indeks chegarasidan chiqish (out-of-bounds). Masalan, int a[5] da a[5] ga murojaat — bu xotiraning ruxsat etilmagan qismiga kirish. C++ bu xatoni kompilyatsiya vaqtida aniqlamaydi, dastur ishlaganda kutilmagan natija beradi.",
            "Yana bir xato — massiv o'lchamini noto'g'ri hisoblash. Masalan, n ta element uchun sikl i<n emas i<=n deb yozilsa, ortiqcha bitta element qaraladi (n-indeks, mavjud emas). Bu 'off-by-one' xatosi deb ataladi.",
            "Oldini olish usullari: massiv o'rniga vector<int> ishlatish (o'lchami dinamik, at() metodi chegarani tekshiradi); doimiy o'lchamlar uchun const int SIZE = 100; ishlatish; siklda doimo i < size sharti bilan ishlash. Zamonaviy C++ da oddiy massivlar o'rniga std::array yoki std::vector tavsiya etiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 30,
    "title": "Bir o'lchovli massivlar. Ko'p o'lchovli massivlar. Dinamik massivlar",
    "reja": [
        "Bir o'lchovli massivlar bilan amaliy ishlash.",
        "Ko'p o'lchovli (ikki o'lchovli) massivlar.",
        "Dinamik massivlar va xotira boshqarish.",
        "std::vector — zamonaviy alternativa.",
    ],
    "sections": [
        ("Bir o'lchovli massivlar bilan amaliy ishlash", [
            "Bir o'lchovli massiv — elementlar chiziqli ketma-ketlikda joylashgan tuzilma. Eng ko'p uchraydigan amallar: elementlarni kiritish/chiqarish, yig'indi/o'rta hisoblash, max/min topish, saralash, qidirish va massivni aylantirish (reverse).",
            "Massivni teskarisiga aylantirish: for (int i=0; i<n/2; i++) { int temp=a[i]; a[i]=a[n-1-i]; a[n-1-i]=temp; } — boshi va oxiridagi elementlar almashtiriladi, markaz tomon harakatlaniladi. Bu oddiy, lekin ko'p ishlatiladigan algoritmdir.",
            "Chiziqli qidirish: for (int i=0; i<n; i++) { if (a[i]==key) return i; } return -1; — har bir element tekshiriladi, topilsa indeksi qaytariladi, topilmasa -1. Tartiblangan massivda ikkilik qidirish (binary search) tezroq ishlaydi — O(log n).",
        ]),
        ("Ko'p o'lchovli (ikki o'lchovli) massivlar", [
            "Ikki o'lchovli massiv — jadval (matritsa) ko'rinishidagi ma'lumotlar tuzilmasi. E'lon: int mat[m][n]; bu m ta satr va n ta ustundan iborat matritsa. Elementga murojaat: mat[i][j] — i-satr, j-ustundagi element (ikkala indeks 0 dan boshlanadi).",
            "Ikki o'lchovli massivni to'ldirish va chiqarish ichma-ich sikllar bilan bajariladi: for (int i=0; i<m; i++) for (int j=0; j<n; j++) cin >> mat[i][j]; Matritsa elementlari xotirada satr bo'yicha ketma-ket joylashadi (row-major order).",
            "Matritsa amallari: satr/ustun yig'indilari, bosh diagonal elementlari (mat[i][i]), transponerlash (satr va ustunni almashtirish), ikkita matritsani qo'shish va ko'paytirish. Bu amallar chiziqli algebra va texnik hisob-kitoblarda keng qo'llaniladi.",
        ]),
        ("Dinamik massivlar va xotira boshqarish", [
            "Statik massivning o'lchami kompilyatsiya vaqtida ma'lum bo'lishi kerak. Dinamik massiv esa dastur ishlash vaqtida (runtime) kerakli o'lchamda yaratiladi. C++ da new operatori orqali: int* a = new int[n]; — n ta elementli massiv xotirada yaratiladi.",
            "Dinamik massiv bilan ishlash oddiygidek: a[0], a[1], ... — indekslar bilan murojaat qilinadi. Ammo ish tugagach, xotirani qaytarish SHART: delete[] a; — aks holda xotira oqishi (memory leak) yuzaga keladi. Bu C++ da muhim javobgarlik.",
            "Ikki o'lchovli dinamik massiv: int** mat = new int*[m]; for (int i=0; i<m; i++) mat[i] = new int[n]; — avval satrlar uchun ko'rsatkichlar massivi, keyin har bir satr uchun alohida massiv yaratiladi. O'chirish ham teskari tartibda: avval har bir satr, keyin ko'rsatkichlar massivi.",
        ]),
        ("std::vector — zamonaviy alternativa", [
            "std::vector — C++ standart kutubxonasidagi dinamik massiv bo'lib, o'lchamini avtomatik boshqaradi. E'lon: vector<int> v(n); yoki vector<int> v = {1,2,3}; Element qo'shish: v.push_back(x); — massiv oxiriga element qo'shadi, o'lcham avtomatik oshadi.",
            "vector ning afzalliklari: o'lcham dinamik (istalgan vaqtda oshadi/kamayadi), xotira avtomatik boshqariladi (delete kerak emas), at() metodi indeks chegarasini tekshiradi, size() metodi joriy o'lchamni qaytaradi, STL algoritmlari bilan ishlaydi.",
            "vector bilan ishlash: for (int i=0; i<v.size(); i++) cout << v[i]; yoki for (int x : v) cout << x; Saralash: sort(v.begin(), v.end()); Zamonaviy C++ da oddiy massivlar (C-style arrays) o'rniga vector ishlatish kuchli tavsiya etiladi — u xavfsizroq, qulayroq va moslashuvchan.",
        ]),
    ],
})



# ==================================================================
# FAYLNI GENERATSIYA QILISH
# ==================================================================
if __name__ == "__main__":
    write_docx("/projects/sandbox/Konspekt_30_mavzu.docx")

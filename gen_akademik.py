# -*- coding: utf-8 -*-
"""
Akademik yozuv - 10 mavzu konspekt generatori.
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
        '<dc:title>Akademik yozuv - Konspekt</dc:title>'
        '<dc:creator>Konspekt generatori</dc:creator>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
        '</cp:coreProperties>' % (now, now))

APP_PROPS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    '<Application>Konspekt generatori</Application></Properties>')

def write_docx(path="Akademik_yozuv_10_mavzu.docx"):
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

TOPICS.append({
    "num": 1,
    "title": "Akademik yozuvga kirish. Ilmiy matnning tuzilishi",
    "reja": [
        "Akademik yozuv tushunchasi va uning maqsadi.",
        "Ilmiy matn va kundalik matn o'rtasidagi farqlar.",
        "Ilmiy matnning umumiy tuzilishi va komponentlari.",
        "Akademik yozuvning asosiy tamoyillari va talablari.",
    ],
    "sections": [
        ("Akademik yozuv tushunchasi va uning maqsadi", [
            "Akademik yozuv — bu ilmiy bilimlarni yaratish, taqdim etish va tarqatish maqsadida qo'llaniladigan maxsus yozma nutq shaklidir. U universitetlar, ilmiy muassasalar va tadqiqot markazlarida qo'llaniladigan rasmiy yozuv uslubi hisoblanadi. Akademik yozuvning asosiy maqsadi — yangi bilimlarni aniq, mantiqiy va asoslangan holda ifoda etishdir.",
            "Akademik yozuv nafaqat ilmiy maqolalar yozish, balki referat, kurs ishi, bitiruv malakaviy ishi, dissertatsiya va ilmiy hisobotlar tayyorlashni ham o'z ichiga oladi. U tadqiqotchi fikrlarini tizimli va izchil tarzda ifodalash uchun zarur vosita sanaladi.",
            "Akademik yozuvni egallash talabalarga o'z fikrlarini aniq bayon etish, tanqidiy fikrlash, manbalar bilan ishlash va ilmiy jamoatchilik bilan muloqot qilish ko'nikmalarini beradi. Bu ko'nikmalar nafaqat ilm-fanda, balki kasbiy faoliyatda ham muhim ahamiyat kasb etadi.",
        ]),
        ("Ilmiy matn va kundalik matn o'rtasidagi farqlar", [
            "Ilmiy matn kundalik matndan bir qator jihatlari bilan farqlanadi: u rasmiy uslubda yoziladi, shaxsiy fikr o'rniga dalillarga tayanadi, aniq atamalar qo'llaniladi va mantiqiy tuzilmaga ega. Kundalik matndagi emotsional ifodalar, so'zlashuv uslubi va noaniq iboralar ilmiy matnga xos emas.",
            "Ilmiy matnda har bir da'vo dalillar bilan asoslanishi shart: statistik ma'lumotlar, eksperimental natijalar, avvalgi tadqiqotlarga havolalar va mantiqiy xulosalar. Kundalik matndagi umumiy fikrlar va asossiz bayonotlar ilmiy matnda qabul qilinmaydi.",
            "Ilmiy matn obyektiv bo'lishi kerak — muallif shaxsiy his-tuyg'ularini emas, balki faktlar va tahlilga asoslangan xulosalarni taqdim etadi. Matndagi har bir tushuncha aniq ta'riflanadi va bir ma'noda qo'llaniladi. Bu ilmiy muloqotning aniqligi va ishonchliligini ta'minlaydi.",
        ]),
        ("Ilmiy matnning umumiy tuzilishi va komponentlari", [
            "Ilmiy matnning klassik tuzilishi IMRaD modeli deb ataladi: Introduction (kirish), Methods (usullar), Results (natijalar) va Discussion (muhokama). Bu tuzilma tadqiqot jarayonini mantiqiy izchillikda taqdim etish imkonini beradi.",
            "Kirish qismida tadqiqot muammosi, uning dolzarbligi, mavjud adabiyotlar sharhi va tadqiqot maqsadi bayon etiladi. Usullar qismida tadqiqot qanday o'tkazilgani batafsil yoritiladi. Natijalar qismida olingan ma'lumotlar taqdim etiladi, muhokama qismida esa natijalar talqin qilinadi.",
            "Bundan tashqari ilmiy matnga annotatsiya (qisqacha mazmun), kalit so'zlar, xulosa va foydalanilgan adabiyotlar ro'yxati ham kiradi. Har bir komponent o'z vazifasini bajarib, yaxlit ilmiy asarni tashkil etadi.",
        ]),
        ("Akademik yozuvning asosiy tamoyillari va talablari", [
            "Akademik yozuvning asosiy tamoyillari: aniqlik (har bir so'z va jumla bir ma'noli bo'lishi), ixchamlik (ortiqcha so'zlarsiz, lo'nda bayon), mantiqiylik (fikrlar izchil, bir-biridan kelib chiquvchi), obyektivlik (shaxsiy his-tuyg'ularsiz) va asoslilik (har bir da'vo dalil bilan tasdiqlangan).",
            "Akademik yozuvda rasmiy uslubga rioya qilish talab etiladi: qisqartmalardan qochish (to'liq shaklda yozish), so'zlashuv iboralarini ishlatmaslik, birinchi shaxsdan yozmaslik (u o'rniga passiv konstruksiyalar) va terminologiyadan izchil foydalanish.",
            "Plagiatdan qochish — akademik yozuvning eng muhim talabi. Boshqa mualliflar fikrlari va ma'lumotlari albatta havola (sitata, parafraz) bilan ko'rsatilishi kerak. Plagiat akademik vijdon buzilishi hisoblanadi va jiddiy oqibatlarga olib keladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 2,
    "title": "Ilmiy mavzuni tanlash va ma'lumot yig'ish",
    "reja": [
        "Ilmiy mavzuni tanlash tamoyillari va mezonlari.",
        "Tadqiqot muammosini shakllantirish.",
        "Ma'lumot manbalari va ularni qidirish usullari.",
        "Ma'lumotlarni baholash va tizimlashtirish.",
    ],
    "sections": [
        ("Ilmiy mavzuni tanlash tamoyillari va mezonlari", [
            "Ilmiy mavzuni tanlash — tadqiqotning birinchi va eng muhim bosqichi. To'g'ri tanlangan mavzu tadqiqotchini ilhomlantiradi va uzoq muddatli izlanishga turtki beradi. Mavzu tanlashda shaxsiy qiziqish, sohaning dolzarbligi va tadqiqot imkoniyatlari hisobga olinadi.",
            "Mavzu tanlash mezonlari: dolzarblik (hozirgi kundagi ahamiyati), yangilik (avvalgi ishlardan farqi), amalga oshirilishi mumkinligi (resurslar yetarliligi), ilmiy ahamiyati (nazariy yoki amaliy hissasi) va tadqiqot doirasining aniqligi (juda keng yoki juda tor bo'lmasligi).",
            "Mavzuni aniqlashtirishda keng sohadan boshlab, bosqichma-bosqich toraytiriladi. Masalan: tilshunoslik — pragmatika — iltimos ifodalash usullari — o'zbek tilidagi iltimos ifodalash strategiyalari. Bunda mavzu aniq, o'lchanadigan va bajarilishi mumkin darajaga yetadi.",
        ]),
        ("Tadqiqot muammosini shakllantirish", [
            "Tadqiqot muammosi — mavjud bilimlardagi bo'shliq yoki yechilmagan savol bo'lib, tadqiqot aynan shu muammoni hal qilishga qaratiladi. Muammoni aniqlash uchun mavjud adabiyotlar chuqur o'rganiladi va nima ma'lum, nima noma'lum ekanligi aniqlanadi.",
            "Tadqiqot muammosidan tadqiqot savoli va gipoteza kelib chiqadi. Tadqiqot savoli aniq, o'lchanadigan va javob topilishi mumkin bo'lishi kerak. Gipoteza esa tadqiqot savoliga taxminiy javob bo'lib, u tekshirilishi va tasdiqlanishi yoki rad etilishi mumkin.",
            "Tadqiqot maqsadi va vazifalari muammo asosida shakllantiriladi. Maqsad — tadqiqotning umumiy yo'nalishi; vazifalar — maqsadga erishish uchun bajariladigan aniq qadamlar. Ular bir-biri bilan mantiqiy bog'liq va izchil bo'lishi kerak.",
        ]),
        ("Ma'lumot manbalari va ularni qidirish usullari", [
            "Ilmiy ma'lumot manbalari birlamchi (original tadqiqotlar: maqolalar, dissertatsiyalar, hisobotlar) va ikkilamchi (sharh maqolalar, darsliklar, ensiklopediyalar) turlarga bo'linadi. Birlamchi manbalar ishonchlilik jihatdan ustuvor hisoblanadi.",
            "Elektron ilmiy bazalar: Google Scholar (bepul, keng qamrovli), Scopus va Web of Science (impakt-faktorli jurnallar), PubMed (tibbiyot), JSTOR (ijtimoiy fanlar), ResearchGate va Academia.edu (tadqiqotchilar tarmog'i). Ular millionlab ilmiy ishlarni o'z ichiga oladi.",
            "Qidirish strategiyasi: kalit so'zlarni aniqlash, sinonim va atamalarni ishlatish, Boolean operatorlari (AND, OR, NOT) qo'llash, natijalarni sanaga va manbaga ko'ra filtrlash. Qor to'pi usuli — bitta yaxshi maqolaning adabiyotlar ro'yxatidan boshqa tegishli manbalarni topish.",
        ]),
        ("Ma'lumotlarni baholash va tizimlashtirish", [
            "Manbani baholash mezonlari: muallif malakasi (ilmiy darajasi, soha mutaxassisi), nashr manbasi (retsenziyalangan jurnalmi), sanasi (dolzarbmi), metodologiyasi (ishonchli usullar qo'llanganmi) va havolalar soni (boshqa olimlar tomonidan qancha iqtibos qilingan).",
            "Ma'lumotlarni tizimlashtirish usullari: kartoteka tuzish (har bir manba uchun alohida yozuv), tematik guruhlash (mavzular bo'yicha tartiblash), jadval tuzish (mualliflar, yillar, xulosalar taqqoslash) va kontseptual xarita yaratish (tushunchalar orasidagi bog'lanishlarni ko'rsatish).",
            "Referens menejerlari — manbalarni saqlash va boshqarish dasturlari: Zotero (bepul), Mendeley, EndNote. Ular manbalarni saqlaydi, avtomatik sitata yaratadi va adabiyotlar ro'yxatini to'g'ri formatda shakllantiradi. Bu tadqiqotchining vaqtini sezilarli tejaydi.",
        ]),
    ],
})

TOPICS.append({
    "num": 3,
    "title": "Ilmiy maqola va referat tuzilishi",
    "reja": [
        "Ilmiy maqolaning tuzilishi va yozish bosqichlari.",
        "Referatning tuzilishi va rasmiylashtirilishi.",
        "Annotatsiya va xulosa yozish qoidalari.",
        "Adabiyotlar ro'yxatini to'g'ri shakllantirish.",
    ],
    "sections": [
        ("Ilmiy maqolaning tuzilishi va yozish bosqichlari", [
            "Ilmiy maqola quyidagi tarkibiy qismlardan iborat: sarlavha, mualliflar ma'lumoti, annotatsiya, kalit so'zlar, kirish, asosiy qism (nazariy asos, metodologiya, natijalar, muhokama), xulosa va adabiyotlar ro'yxati. Har bir qismning o'z vazifasi va talablari mavjud.",
            "Sarlavha qisqa (10-15 so'z), aniq va tadqiqot mazmunini to'liq aks ettiruvchi bo'lishi kerak. Annotatsiya (150-250 so'z) maqolaning qisqacha mazmunidir: muammo, usul, natija va xulosa. Kalit so'zlar (4-6 ta) maqolani qidiruv tizimlarida topilishini ta'minlaydi.",
            "Maqola yozish bosqichlari: 1) tadqiqot o'tkazish, 2) natijalarni tahlil qilish, 3) maqola rejasini tuzish, 4) qoralama (draft) yozish, 5) qayta ko'rib chiqish va tahrirlash, 6) formatlashtirish (jurnal talablariga moslashtirish), 7) jurnalga yuborish. Qayta tahrirlash sifat kafolati.",
        ]),
        ("Referatning tuzilishi va rasmiylashtirilishi", [
            "Referat — muayyan mavzu bo'yicha mavjud adabiyotlarni o'rganib, ularning qisqacha mazmunini bayon etuvchi ilmiy ishdir. U mustaqil tadqiqot emas, balki manbalar tahlili va umumlashtirilmasi hisoblanadi. Referat talabalarning ilmiy ish bilan tanishishning birinchi bosqichi sifatida muhim.",
            "Referatning tuzilmasi: titul varag'i (mavzu, muallif, rahbar), mundarija, kirish (mavzuning dolzarbligi, maqsad va vazifalar), asosiy qism (boblar va paragraflar), xulosa (xulosalar va takliflar) va foydalanilgan adabiyotlar ro'yxati.",
            "Referatni rasmiylashtirish talablari: A4 format, Times New Roman 14pt, 1.5 interval, chapdan 3 sm, o'ngdan 1.5 sm, yuqori va pastdan 2 sm maydon. Sahifalar raqamlanadi (titul sahifasiz). Hajmi odatda 15-25 sahifa. Har bir bob yangi sahifadan boshlanadi.",
        ]),
        ("Annotatsiya va xulosa yozish qoidalari", [
            "Annotatsiya (abstract) — ilmiy ishning qisqacha mazmuni bo'lib, u o'quvchiga ishni to'liq o'qimasdan asosiy mazmunini tushunish imkonini beradi. Annotatsiya mustaqil matn sifatida tushunilishi kerak: muammo, usul, asosiy natijalar va xulosa 150-300 so'z ichida bayon etiladi.",
            "Xulosa (conclusion) — ilmiy ishning yakuniy qismi bo'lib, unda tadqiqot natijalari umumlashtiriladi, tadqiqot savollariga javob beriladi va asosiy xulosalar shakllantiriladi. Xulosa yangi ma'lumot kiritilmaydi, faqat ilgari bayon etilgan natijalar qisqacha takrorlanadi.",
            "Yaxshi annotatsiya va xulosaning xususiyatlari: ixchamlik (faqat muhim ma'lumotlar), to'liqlik (asosiy fikrlar qamrab olingan), mustaqillik (boshqa qismlarsiz tushuniladi) va aniqlik (noaniq iboralar yo'q). Ular ko'pincha ishning eng ko'p o'qiladigan qismlari.",
        ]),
        ("Adabiyotlar ro'yxatini to'g'ri shakllantirish", [
            "Adabiyotlar ro'yxati — ilmiy ishda foydalanilgan barcha manbalarning to'liq bibliografik ma'lumotlari. Har bir matnda ishlangan havola ro'yxatda aks etishi va har bir ro'yxatdagi manba matnda havolasi bo'lishi kerak.",
            "Bibliografik tavsif elementlari: muallif(lar) ismi, nashr yili, asar nomi, jurnalnomi (maqola uchun), nashriyot, sahifalar soni. Formatlar turlicha: APA (ijtimoiy fanlar), MLA (gumanitar fanlar), Chicago, Harvard va GOST (MDH mamlakatlari).",
            "APA formati misoli: Karimov, A. B. (2023). Akademik yozuv asoslari. Toshkent: Fan nashriyoti. Maqola uchun: Rahimova, D. S. (2024). Ilmiy uslub xususiyatlari. Tilshunoslik jurnali, 5(2), 34-45. Izchillik muhim — bir format butun ish davomida saqlanadi.",
        ]),
    ],
})

TOPICS.append({
    "num": 4,
    "title": "Matnni ma'noviy qismlarga bo'lish",
    "reja": [
        "Matnni bo'laklash tushunchasi va zaruriyati.",
        "Paragraf tuzilishi va uning vazifalari.",
        "Matnda mantiqiy bog'lanish va kogerensiya.",
        "Ma'noviy qismlar orasidagi o'tish usullari.",
    ],
    "sections": [
        ("Matnni bo'laklash tushunchasi va zaruriyati", [
            "Matnni ma'noviy qismlarga bo'lish — bu yirik matnni mantiqiy jihatdan tugallangan kichikroq bo'laklarga ajratish jarayoni. Bu bo'laklash o'quvchiga matnni osonroq idrok etish, asosiy fikrlarni ajratib ko'rish va matn mantiqini kuzatish imkonini beradi.",
            "Ilmiy matnda bo'laklash bir nechta darajada amalga oshiriladi: butun ish qismlarga (kirish, asosiy qism, xulosa), asosiy qism boblarga, boblar paragraflarga, paragraflar esa jumlalarga bo'linadi. Har bir daraja o'z mantiqiy yaxlitligiga ega.",
            "To'g'ri bo'laklanmagan matn o'quvchini charchatadi, asosiy fikrlar yo'qoladi va matn tushunarsiz bo'lib qoladi. Shu sababli matn strukturasini oldindan rejalashtirish va har bir qismning vazifasini aniq belgilash muhim ahamiyatga ega.",
        ]),
        ("Paragraf tuzilishi va uning vazifalari", [
            "Paragraf — bitta asosiy fikrni rivojlantiruvchi jumlalar guruhi. Har bir paragrafda bitta markaziy fikr (topic sentence) bo'lishi va qolgan jumlalar uni qo'llab-quvvatlashi, dalillashi yoki tushuntirishi kerak.",
            "Paragrafning klassik tuzilmasi: tema jumlasi (asosiy fikrni bildiradi) — rivojlantirish (dalillar, misollar, tushuntirish) — yakunlovchi jumla (fikrni xulosalab, keyingi paragrafga o'tish). Bu tuzilma paragrafning mantiqiy yaxlitligini ta'minlaydi.",
            "Paragraf uzunligi muvozanatli bo'lishi kerak: juda qisqa (1-2 jumla) paragraflar fikrni yetarli rivojlantirmaydi; juda uzun (sahifadan ortiq) paragraflar o'quvchini charchatadi. Ilmiy matn uchun 5-8 jumla (100-200 so'z) optimal hajm hisoblanadi.",
        ]),
        ("Matnda mantiqiy bog'lanish va kogerensiya", [
            "Kogerensiya — matnning mantiqiy yaxlitligi va izchilligi. Kogerent matndagi barcha qismlar bir-biri bilan mantiqiy bog'liq, har bir jumla oldingi fikrdan kelib chiqadi va keyingisiga asos yaratadi. Bu o'quvchiga muallifning fikr yo'lini osongina kuzatish imkonini beradi.",
            "Koheziya — matn qismlarini grammatik va leksik vositalar orqali bog'lash. Koheziya vositalari: ko'rsatish olmoshlari (bu, shu, u), bog'lovchilar (shu sababli, shuning uchun), leksik takror (kalit so'zlarni takrorlash) va sinonimlar (bir xil tushunchani turli so'zlar bilan ifodalash).",
            "Ilmiy matnda mantiqiy bog'lanish ayniqsa muhim: o'quvchi har bir yangi fikrning oldingi bilan qanday bog'liqligini tushunishi kerak. Buni ta'minlash uchun o'tish so'zlari, aniq havolalar va mantiqiy ketma-ketlikni saqlab yozish zarur.",
        ]),
        ("Ma'noviy qismlar orasidagi o'tish usullari", [
            "O'tish so'zlari va iboralari turli mantiqiy munosabatlarni ifodalaydi: qo'shish (bundan tashqari, shuningdek), qarama-qarshi qo'yish (biroq, aksincha, shunga qaramay), sabab-oqibat (shu sababli, natijada, demak), vaqt (avvalo, keyin, nihoyat) va misol keltirish (masalan, jumladan).",
            "Boblar va katta qismlar orasidagi o'tish uchun alohida o'tish paragraflari yoziladi. Bunday paragraf avvalgi qismda nimalar bayon etilganini qisqacha eslatib, keyingi qismda nimalar muhokama qilinishini oldindan aytib beradi.",
            "Samarali o'tishning muhim sharti — har bir yangi qism avvalgisidan mantiqan kelib chiqishi. O'quvchi hech qachon savol bermasligi kerak: nima uchun muallif bu yerda bu haqida gapirmoqda? Har bir o'tish tabiiy va asosli bo'lishi lozim.",
        ]),
    ],
})

TOPICS.append({
    "num": 5,
    "title": "Ilmiy uslub grammatikasi",
    "reja": [
        "Ilmiy uslubning grammatik xususiyatlari.",
        "Ilmiy matndagi gap tuzilishi va uning xususiyatlari.",
        "Fe'l shakllari va nisbatlarining qo'llanilishi.",
        "Atamalar va ilmiy leksikaning o'ziga xosligi.",
    ],
    "sections": [
        ("Ilmiy uslubning grammatik xususiyatlari", [
            "Ilmiy uslub grammatikasi kundalik nutqdan sezilarli farqlanadi. Uning asosiy xususiyatlari: murakkab gap tuzilmalari, passiv nisbatning ko'p qo'llanilishi, nominalizatsiya (fe'l o'rniga ot ishlatish), shartli va maqsad ergash gaplar hamda ilmiy atamalarning izchil qo'llanilishidir.",
            "Ilmiy matndagi gaplar odatda kundalik nutqdagiga qaraganda uzunroq va murakkab tuzilishga ega. Bu murakkablik muallifning fikrni aniq va to'liq ifodalash zaruratidan kelib chiqadi: shartlar, chegaralar va munosabatlar bir gapda ifodalanadi.",
            "Ilmiy uslubda shaxssizlik muhim xususiyat: men o'rniga tadqiqot, tahlil, kuzatish kabi otlar subjekt sifatida ishlatiladi. Bu matnga obyektivlik beradi va diqqatni muallifdan tadqiqot predmetiga yo'naltiradi.",
        ]),
        ("Ilmiy matndagi gap tuzilishi va uning xususiyatlari", [
            "Ilmiy matndagi gap turlari: darak gap (asosiy tur — ma'lumot beradi), so'roq gap (kamdan-kam, ritorik savollar uchun), buyruq gap (deyarli ishlatilmaydi). Ilmiy matn asosan darak gaplardan iborat bo'lib, ular mazmunni obyektiv bayon etadi.",
            "Qo'shma gaplar ilmiy matndagi asosiy grammatik birlik: boglovchili qo'shma gap (sabab-oqibat: chunki, shu sababli; qarshi qo'yish: ammo, biroq; qo'shish: va, hamda) va ergashgan qo'shma gap (aniqlovchi, to'ldiruvchi, hol ergash gaplar).",
            "Gapning ma'lum-noma'lum tuzilishi: ilmiy matnda yangi axborot gapning oxiriga, ma'lum axborot esa boshiga joylashtiriladi. Bu o'quvchini ma'lum ma'lumotdan yangi ma'lumotga olib boradi va matn oqimini tabiiy qiladi.",
        ]),
        ("Fe'l shakllari va nisbatlarining qo'llanilishi", [
            "Ilmiy matnda vaqt shakllari: hozirgi zamon (umumiy haqiqatlar va xulosalar uchun: suv 100 darajada qaynaydi), o'tgan zamon (tadqiqot natijalari va boshqalar ishlari uchun: tajriba o'tkazildi, Karimov ta'kidlagan) va kelasi zamon (kamdan-kam, rejalar uchun).",
            "Passiv nisbat ilmiy matnda keng qo'llaniladi: tajriba o'tkazildi (kim tomonidan — muhim emas), natijalar tahlil qilindi, so'rovnoma o'tkazildi. Passiv nisbat muallifdan emas, balki jarayondan, natijadan xabar beradi va matnga obyektiv ohang beradi.",
            "Modal fe'llar ehtimollik va ishonch darajasini ifodalaydi: mumkin (ehtimollik), kerak (zaruriyat), lozim (tavsiya). Ilmiy matnda ehtiyotkorlik muhim: to'g'ridan-to'g'ri da'vo o'rniga yumshoq ifodalash (hedging): bu natija ... bo'lishi mumkin, ehtimol ... ta'sir ko'rsatadi.",
        ]),
        ("Atamalar va ilmiy leksikaning o'ziga xosligi", [
            "Ilmiy atama — muayyan sohadagi aniq tushunchani ifodalovchi so'z yoki iboradir. Atamalar bir ma'noli va barcha mutaxassislar uchun bir xil tushuniladi. Masalan: morfema (tilshunoslikda), fotosintez (biologiyada), diskriminant (matematikada).",
            "Ilmiy matnning leksik tarkibi uch qatlamdan iborat: umumiste'mol so'zlar (barcha matnlarda uchraydi), umumilmiy so'zlar (tadqiqot, tahlil, natija, metod — barcha fanlarda) va maxsus atamalar (faqat muayyan sohada qo'llaniladi).",
            "Ilmiy matnda so'z tanlashda aniqlik va izchillik muhim: bir tushuncha uchun butun matn davomida bitta atama ishlatiladi (sinonimdan qochiladi); yangi atama kiritilganda u albatta ta'riflanadi; umumiy so'zlar o'rniga aniq atamalar afzal ko'riladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 6,
    "title": "Tavsif va tahlil metodikalari",
    "reja": [
        "Ilmiy tavsif tushunchasi va uning turlari.",
        "Tahlil metodikasi va uning bosqichlari.",
        "Taqqoslash va klassifikatsiya usullari.",
        "Ilmiy matndagi misollar va dalillar keltirish qoidalari.",
    ],
    "sections": [
        ("Ilmiy tavsif tushunchasi va uning turlari", [
            "Ilmiy tavsif — tadqiqot obyektini uning xususiyatlari, tarkibi, tuzilishi va funksiyalari bo'yicha batafsil yoritish jarayoni. Tavsif o'quvchiga obyektni to'liq tasavvur qilish imkonini beradi. U tahlilning dastlabki bosqichi sifatida muhim ahamiyatga ega.",
            "Ilmiy tavsif turlari: strukturaviy tavsif (obyektning tuzilishi, qismlari), funksional tavsif (vazifasi, ishlash prinsipi), genetik tavsif (kelib chiqishi, rivojlanish tarixi) va taqqosiy tavsif (boshqa obyektlar bilan farqlari va o'xshashliklari).",
            "Tavsif yozish qoidalari: aniq va to'liq bo'lish, ortiqcha shaxsiy baho bermaslik, mantiqiy tartibda bayon etish (masalan, umumiydan xususiyga yoki qismlardan yaxlitga) va o'quvchi uchun tushunarli til ishlatish.",
        ]),
        ("Tahlil metodikasi va uning bosqichlari", [
            "Tahlil (analiz) — butunni qismlarga ajratib, har bir qismni alohida o'rganish va ular orasidagi munosabatlarni aniqlash jarayoni. Ilmiy matndagi tahlil dalillar asosida xulosalar chiqarish, tendensiyalarni aniqlash va qonuniyatlarni topish maqsadida bajariladi.",
            "Tahlilning asosiy bosqichlari: 1) tahlil obyektini aniqlash, 2) mezonlarni belgilash, 3) obyektni qismlarga ajratish, 4) har bir qismni alohida o'rganish, 5) qismlar orasidagi munosabatlarni aniqlash, 6) xulosalar shakllantirish.",
            "Tahlil turlari: miqdoriy tahlil (statistik ma'lumotlar asosida, raqamlar bilan), sifat tahlili (ma'no, mazmun, xususiyatlar asosida), SWOT tahlili (kuchli-zaif tomonlar, imkoniyat-tahdidlar), kontent-tahlil (matn mazmunini tizimli o'rganish) va diskurs-tahlil (til va ijtimoiy kontekst).",
        ]),
        ("Taqqoslash va klassifikatsiya usullari", [
            "Taqqoslash (komparativ usul) — ikki yoki undan ortiq obyektni muayyan mezonlar bo'yicha solishtirish. Taqqoslash ikki yo'nalishda amalga oshiriladi: o'xshashliklarni topish (umumiy jihatlar) va farqlarni aniqlash (o'ziga xos jihatlar). Natijalar ko'pincha jadval shaklida taqdim etiladi.",
            "Klassifikatsiya — obyektlarni muayyan mezon asosida guruhlarga ajratish. Klassifikatsiya mantiqiy bo'lishi kerak: bir mezon asosida (masalan, hajmiga ko'ra), to'liq qamrovli (barcha elementlar kiritilgan), bir-birini istisno etuvchi (har bir element faqat bitta guruhga tegishli).",
            "Ilmiy matndagi taqqoslash va klassifikatsiya uchun signal so'zlar: taqqoslash uchun — o'xshash, shunga yaqin, xuddi shunday, farqli o'laroq, aksincha; klassifikatsiya uchun — turlarga bo'linadi, guruhlanadi, quyidagi kategoriyalar ajratiladi.",
        ]),
        ("Ilmiy matndagi misollar va dalillar keltirish qoidalari", [
            "Misol — umumiy qoidani aniq holat bilan tushuntiradigan ko'rgazmali material. Ilmiy matndagi misollar aniq, tegishli va ishonchli bo'lishi kerak. Ular fikrni kuchaytiradi va o'quvchiga mavhum tushunchani yaxshiroq anglash imkonini beradi.",
            "Dalil — da'voni tasdiqlovchi yoki rad etuvchi fakt, statistika yoki mantiqiy asoslash. Dalil turlari: empirik (tajriba natijalari, statistika), mantiqiy (fikr yuritish, sillogizm), avtoritar (mutaxassis fikri, havola) va analogiya (o'xshash holatga tayanish).",
            "Misollar va dalillarni keltirish uchun signal iboralar: masalan, jumladan, xususan, bunga misol sifatida, bu fikrni quyidagi dalil tasdiqlaydi, statistik ma'lumotlarga ko'ra, tadqiqot natijalari shuni ko'rsatadiki. Bu iboralar o'quvchini navbatdagi dalilga tayyorlaydi.",
        ]),
    ],
})

TOPICS.append({
    "num": 7,
    "title": "Nutq uslublari",
    "reja": [
        "Nutq uslubi tushunchasi va uning turlari.",
        "Ilmiy nutq uslubining xususiyatlari.",
        "Rasmiy (ish yuritish) uslubi.",
        "Badiiy va publitsistik uslublar bilan taqqoslash.",
    ],
    "sections": [
        ("Nutq uslubi tushunchasi va uning turlari", [
            "Nutq uslubi — muayyan muloqot sohasida qo'llaniladigan til vositalarining tanlash va uyushtirish usuli. Nutq uslublari so'zlashuvchi maqsadi, auditoriyasi va muloqot holatiga qarab farqlanadi. Tilshunoslikda beshta asosiy funksional uslub ajratiladi.",
            "Nutq uslublarining beshtaligi: ilmiy uslub (ilm-fan sohasi), rasmiy uslub (ish yuritish, qonunchilik), publitsistik uslub (ommaviy axborot vositalari), badiiy uslub (adabiyot) va so'zlashuv uslubi (kundalik muloqot). Har bir uslubning o'z leksik, grammatik va kompozitsion xususiyatlari bor.",
            "Uslub tanlash muloqot vaziyatiga bog'liq: kim bilan gaplashilmoqda (auditoriya), qayerda (rasmiy yoki norasmiy muhit), nima maqsadda (ma'lumot berish, ishontirish, ta'sir ko'rsatish) va qaysi shaklda (og'zaki yoki yozma). Noto'g'ri uslub tanlash muloqot buzilishiga olib keladi.",
        ]),
        ("Ilmiy nutq uslubining xususiyatlari", [
            "Ilmiy uslubning leksik xususiyatlari: maxsus atamalar, umumilmiy so'zlar (tadqiqot, metod, natija), mavhum otlar ko'pligi, emotsional so'zlarning yo'qligi va qat'iy bir ma'nolilik. Ilmiy matndagi har bir so'z aniq vazifa bajaradi va almashtirib bo'lmaydi.",
            "Ilmiy uslubning grammatik xususiyatlari: murakkab gap tuzilmalari, passiv konstruksiyalar, sababiy va shartli ergash gaplar, raqamlar va formulalar, sanash va tasniflovchi konstruksiyalar. Gap uzunligi kundalik nutqqa qaraganda sezilarli yuqori.",
            "Ilmiy uslubning kompozitsion xususiyatlari: qat'iy mantiqiy tuzilma, dalillarga asoslangan bayon, oldindan belgilangan struktura (IMRaD), obyektiv ton va shaxssiz bayon. Bu xususiyatlar ilmiy muloqotning aniqligini va ishonchliligini ta'minlaydi.",
        ]),
        ("Rasmiy (ish yuritish) uslubi", [
            "Rasmiy uslub — davlat boshqaruvi, huquq va ish yuritish sohasida qo'llaniladigan nutq uslubi. Uning hujjatlari: qonunlar, farmoyishlar, buyruqlar, shartnomalar, arizalar, ma'lumotnomalar va rasmiy xatlar. Bu uslub aniqlik va bir ma'nolilik talab etadi.",
            "Rasmiy uslubning xususiyatlari: qat'iy standart shakllar (shablon), maxsus klishire iboralar (shu munosabat bilan, yuqoridagiga asosan), oldindan belgilangan tuzilma, shaxssiz bayon va huquqiy terminologiya. Ijodiy yondashuv va shaxsiy uslub bu yerda o'rinsiz.",
            "Ilmiy va rasmiy uslub o'rtasidagi umumiylik: ikkalasi ham rasmiy, aniq va obyektiv. Farqi: ilmiy uslub dalillaydi va tushuntiradi, rasmiy uslub esa buyuradi va tartibga soladi. Ilmiy matnda shaxsiy fikr (asoslangan) bo'lishi mumkin, rasmiy matnda esa yo'q.",
        ]),
        ("Badiiy va publitsistik uslublar bilan taqqoslash", [
            "Badiiy uslub — adabiyotda qo'llaniladigan, estetik ta'sir ko'rsatishga yo'naltirilgan nutq uslubi. U ko'chimlar (metafora, sifatlash, mubolag'a), emotsional leksika, subyektiv bayon va ijodiy erkinlik bilan xarakterlanadi. Ilmiy uslubda bularning barchasi nomaqbul.",
            "Publitsistik uslub — ommaviy axborot vositalarida qo'llaniladigan, axborot berish va jamoatchilik fikrini shakllantirishga qaratilgan uslub. U ilmiy va badiiy uslub elementlarini birlashtiradi: faktlarga tayanadi (ilmiydek), lekin emotsional ta'sir ham ko'rsatadi (badiiydek).",
            "Taqqoslash jadvali: Ilmiy — obyektiv, atamali, dalilga asoslangan. Badiiy — subyektiv, obrazli, hissiyotga asoslangan. Publitsistik — o'rtacha, faktli, ta'sirga yo'naltirilgan. Rasmiy — standart, qat'iy, huquqiy. So'zlashuv — erkin, emotsional, situativ. Akademik yozuvda ilmiy uslub qo'llaniladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 8,
    "title": "Publitsistik uslub va o'zak uslublari",
    "reja": [
        "Publitsistik uslubning mohiyati va vazifalari.",
        "Publitsistik uslubning o'zak turlari (janrlar).",
        "Publitsistik matndagi til vositalari.",
        "Publitsistik va ilmiy uslubning qiyosiy tahlili.",
    ],
    "sections": [
        ("Publitsistik uslubning mohiyati va vazifalari", [
            "Publitsistik uslub — ommaviy axborot vositalari (gazeta, jurnal, televideniye, radio, internet) sohasida qo'llaniladigan funksional uslub. Uning ikki asosiy vazifasi bor: axborot berish (voqealar haqida xabar qilish) va ta'sir ko'rsatish (jamoatchilik fikrini shakllantirish).",
            "Publitsistik uslubning xususiyatlari: dolzarblik (hozirgi kundagi masalalar), ommaboplik (keng auditoriyaga tushunarli), emotsionallik (hissiy ta'sir), da'vatkorlik (harakatga undash) va standart bilan ekspressivlikning uyg'unligi.",
            "Publitsistik uslub ilmiy matndan farqli ravishda muallif pozitsiyasini ochiq ifodalaydi, emotsional leksikadan foydalanadi va o'quvchini ishontirishga intiladi. Biroq ishonchli publitsistika ham faktlarga tayanadi va asossiz da'volardan qochadi.",
        ]),
        ("Publitsistik uslubning o'zak turlari (janrlar)", [
            "Axborot janrlari: xabar (qisqa, faktli), reportaj (voqea joyidan jonli bayon), intervyu (savol-javob), hisobot (rasmiy tadbirdan). Bu janrlar asosan axborot berish vazifasini bajaradi — obyektiv, qisqa, aniq.",
            "Tahliliy janrlar: maqola (muammoni chuqur tahlil qilish), sharh (voqea yoki hodisaga baho berish), obzor (bir nechta voqeani qiyoslash). Bu janrlarda muallif fikri, tahlili va xulosasi muhim o'rin tutadi.",
            "Badiiy-publitsistik janrlar: ocherk (inson yoki voqea haqida badiiy bayon), esse (erkin shakldagi fikr yuritish), feleton (hajviy tanqid), pamflet (keskin siyosiy tanqid). Bu janrlarda badiiy tasvir vositalari keng qo'llaniladi.",
        ]),
        ("Publitsistik matndagi til vositalari", [
            "Leksik vositalar: ijtimoiy-siyosiy terminologiya, baholovchi leksika (ijobiy yoki salbiy baho), klishire iboralar (qayd etish lozimki, e'tiborga molik), neologizmlar va chet so'zlar. Bu vositalar matnning publitsistik xarakterini belgilaydi.",
            "Grammatik vositalar: ritorik savollar (o'quvchini fikrlashga undash), undov gaplar (emotsional kuchaytirish), qisqa gaplar (ta'sir kuchini oshirish) va takrorlar (muhim fikrni ta'kidlash). Ular matnning ekspressivligini oshiradi.",
            "Kompozitsion vositalar: sarlavhaning jozibadorligi (o'quvchini jalb qilish), lid (dastlabki abzats — eng muhim ma'lumot), teskari piramida printsipi (muhimdan kamroq muhimga) va samarali yakunlash (o'quvchini fikrlashga yoki harakatga undash).",
        ]),
        ("Publitsistik va ilmiy uslubning qiyosiy tahlili", [
            "Umumiy jihatlar: ikkala uslub ham yozma shaklda namoyon bo'ladi, faktlarga tayanadi, mantiqiy tuzilmaga ega va muayyan maqsadga yo'naltirilgan. Har ikkisida ham muallif o'z pozitsiyasini ifodalaydi (ilmiydagi ehtiyotkorroq, publitsistikdagi ochiqroq).",
            "Farqlar: ilmiy uslub — obyektiv, shaxssiz, ehtiyotkor, tor auditoriyaga; publitsistik — subyektiv, shaxsiy, qat'iy, keng auditoriyaga. Ilmiy matn isbotlaydi, publitsistik matn ishontiradi. Ilmiy matn o'quvchini o'ylashga, publitsistik matn harakatga undaydi.",
            "Akademik yozuvda publitsistik uslub nomaqbul: emotsional ifodalar, mubolag'a, asossiz da'volar va subyektiv baholar ilmiy matnda ishlatilmaydi. Biroq ilmiy-ommabop maqolalarda publitsistik elementlar qo'llanilishi mumkin — keng auditoriyaga tushunarli qilish maqsadida.",
        ]),
    ],
})

TOPICS.append({
    "num": 9,
    "title": "Ma'ruza tuzilmasi",
    "reja": [
        "Ma'ruza tushunchasi va uning akademik muhitdagi o'rni.",
        "Ma'ruzaning tuzilmaviy qismlari.",
        "Ma'ruza matnini tayyorlash va yozish.",
        "Ma'ruzani samarali taqdim etish ko'nikmalari.",
    ],
    "sections": [
        ("Ma'ruza tushunchasi va uning akademik muhitdagi o'rni", [
            "Ma'ruza (doklad) — muayyan mavzu bo'yicha auditoriya oldida og'zaki yoki yozma shaklda taqdim etiladigan ilmiy xabar. Akademik muhitda ma'ruza konferensiyalar, seminarlar, himoyalar va dars mashg'ulotlarida taqdim etiladi. U tadqiqot natijalarini jamoatchilikka yetkazishning muhim shakli.",
            "Ma'ruza ilmiy maqoladan farq qiladi: u og'zaki taqdimot uchun mo'ljallangan, vaqt bilan chegaralangan (odatda 10-20 daqiqa), vizual materiallar bilan qo'llab-quvvatlanadi va auditoriya bilan bevosita muloqotni nazarda tutadi.",
            "Ma'ruzaning muvaffaqiyati nafaqat mazmunning sifatiga, balki taqdim etish mahoratiga ham bog'liq. Yaxshi ma'ruza tushuntirishga emas, muloqotga asoslangan: ma'ruzachi auditoriyani jalb qiladi, savollar kutadi va javob beradi.",
        ]),
        ("Ma'ruzaning tuzilmaviy qismlari", [
            "Ma'ruza uchta asosiy qismdan iborat: kirish (introduction), asosiy qism (body) va xulosa (conclusion). Kirish — auditoriya e'tiborini jalb qilish, mavzuni tanishtirish va ma'ruza rejasini aytish. U umumiy vaqtning 10-15 foizini tashkil etadi.",
            "Asosiy qism — ma'ruzaning eng katta qismi (70-80%). U 2-4 ta asosiy fikrdan iborat bo'ladi, har bir fikr dalillar, misollar va vizual materiallar bilan qo'llab-quvvatlanadi. Fikrlar orasida aniq o'tishlar bo'lishi muhim.",
            "Xulosa — ma'ruzaning yakuniy qismi (10-15%). Unda asosiy fikrlar qisqacha takrorlanadi, xulosa shakllantiriladi va auditoriyaga yakuniy xabar beriladi. Ko'pincha qo'shimcha tadqiqot yo'nalishlari yoki amaliy takliflar ham aytiladi. Savol-javob bo'limi odatda xulosadan keyin o'tkaziladi.",
        ]),
        ("Ma'ruza matnini tayyorlash va yozish", [
            "Ma'ruza matni va ilmiy maqola matni orasida muhim farq bor: ma'ruza matni og'zaki nutq uchun mo'ljallangan. Shuning uchun gaplar qisqaroq, tuzilma soddaroq va takrorlar ko'proq bo'lishi tabiiy. Auditoriya matni qayta o'qiy olmaydi — birinchi marta eshitib tushunishi kerak.",
            "Ma'ruza matni tayyorlash bosqichlari: 1) maqsadni aniqlash (nima aytmoqchiman), 2) auditoriyani tahlil qilish (kimga aytmoqchiman), 3) asosiy fikrlarni tanlash (3-4 ta), 4) rejani tuzish, 5) matnni yozish, 6) vizual materiallarni tayyorlash, 7) mashq qilish.",
            "Vaqt boshqaruvi muhim: 10 daqiqalik ma'ruza uchun taxminan 1200-1500 so'z, 20 daqiqalik uchun 2500-3000 so'z. Mashq qilib, vaqtni o'lchash zarur. Vizual materiallar (slaydlar) uchun qoida: bitta slayd — bitta fikr, bitta daqiqa.",
        ]),
        ("Ma'ruzani samarali taqdim etish ko'nikmalari", [
            "Og'zaki taqdim etish ko'nikmalari: aniq va baland ovozda gapirish, auditoriya bilan ko'z aloqasi (ko'z kontakti), tabiiy imo-ishoralar, ovoz tonini o'zgartirish (monotonlikdan qochish) va pauzalardan samarali foydalanish (muhim fikrlarni ta'kidlash uchun).",
            "Vizual materiallar bilan ishlash: slaydlardagi matnni o'qimaslik (ular qo'llab-quvvatlovchi material, asosiy matn emas), slaydlarga ortiqcha matn joylashtirmaslik, grafik va jadvallarni tushuntirish va texnik muammolarga tayyor bo'lish (zaxira nusxa).",
            "Auditoriya bilan muloqot: savollarni kutish va ularga tayyorgarlik ko'rish, qiyin savolga xotirjamlik bilan javob berish, bilmaydigan narsani tan olish va keyinroq javob berishga va'da berish. Muvaffaqiyatli ma'ruzachi auditoriyani faol ishtirokchi sifatida jalb qiladi.",
        ]),
    ],
})

TOPICS.append({
    "num": 10,
    "title": "Ilmiy ishni yozish",
    "reja": [
        "Ilmiy ish yozish jarayoni va uning bosqichlari.",
        "Ilmiy ishda havola berish va sitata keltirish qoidalari.",
        "Ilmiy ishni tahrirlash va takomillashtirish.",
        "Ilmiy ishni rasmiylashtirish va topshirish.",
    ],
    "sections": [
        ("Ilmiy ish yozish jarayoni va uning bosqichlari", [
            "Ilmiy ish yozish chiziqli emas, balki tsiklik jarayondir: rejalashtirish, yozish, qayta ko'rish va tahrirlash bosqichlari bir necha marta takrorlanadi. Har bir qayta ko'rishda ish sifati yaxshilanadi. Birinchi qoralama (draft) mukammal bo'lishi shart emas — muhimi boshlash.",
            "Yozish bosqichlari: 1) mavzuni tanlash va aniqlash, 2) adabiyotlarni o'rganish, 3) tadqiqot savolini shakllantirish, 4) tadqiqotni o'tkazish, 5) batafsil rejani tuzish, 6) birinchi qoralamani yozish, 7) qayta ko'rish va tahrirlash, 8) rasmiylashtirish. Har bir bosqich muhim.",
            "Yozish strategiyalari: yuqoridan pastga (avval reja, keyin to'ldirish), pastdan yuqoriga (avval detalllar, keyin tuzilma), markazdan chetga (avval asosiy qism, keyin kirish va xulosa). Ko'pchilik mutaxassislar kirish qismini eng oxirida yozishni tavsiya etadi — chunki u butun ish mazmunini aks ettirishi kerak.",
        ]),
        ("Ilmiy ishda havola berish va sitata keltirish qoidalari", [
            "Havola (iqtibos) — boshqa muallifning fikri yoki ma'lumotini o'z ishingizda ko'rsatish. Havola berish uchta maqsadga xizmat qiladi: original muallifga hurmat ko'rsatish, da'volarni dalillash va o'quvchiga asl manbani topish imkonini berish.",
            "Havola berish usullari: to'g'ridan-to'g'ri sitata (muallifning so'zlarini aynan keltirish, qo'shtirnoq ichida), parafraz (muallifning fikrini o'z so'zlaringiz bilan ifoda etish, havola bilan) va umumlashtirish (manbaning asosiy g'oyasini qisqacha bayon etish). Barcha usullarda havola ko'rsatilishi shart.",
            "Havola formatlari: matn ichidagi havola (Karimov, 2023, 45-bet) va izoh shaklida (pastki izoh yoki oxiridagi izoh). Format tanlash jurnal yoki muassasa talablariga bog'liq. Muhimi — butun ish davomida bir xil format izchil qo'llanilishi.",
        ]),
        ("Ilmiy ishni tahrirlash va takomillashtirish", [
            "Tahrirlash uch darajada amalga oshiriladi: tuzilmaviy tahrir (mantiqiy izchillik, qismlar tartibi, fikr rivojlanishi), paragraf darajasidagi tahrir (har bir paragrafning yaxlitligi, o'tishlar, dalillar yetarliligi) va jumla darajasidagi tahrir (grammatika, uslub, so'z tanlash, tinish belgilari).",
            "O'z-o'zini tahrirlash usullari: yozgandan keyin biroz dam olish (kamida bir kun), matnni ovoz chiqarib o'qish (noqulay joylar seziladi), qog'ozda chop etib o'qish (ekranda ko'rinmagan xatolar ko'rinadi) va har bir paragrafni alohida tekshirish.",
            "Tashqi fikr olish: ilmiy rahbar bilan maslahatlashish, hamkasblardan fikr so'rash (peer review), yozuv markazidan yordam olish. Tanqidiy fikrlarni shaxsiy qabul qilmaslik va ulardan ish sifatini yaxshilash uchun foydalanish muhim.",
        ]),
        ("Ilmiy ishni rasmiylashtirish va topshirish", [
            "Rasmiylashtirish talablari: belgilangan format (shrift, hajm, interval, maydonlar), titul varag'i, mundarija, kirish, asosiy qism (boblar), xulosa, adabiyotlar ro'yxati va ilovalar (kerak bo'lsa). Har bir muassasaning o'z talablari borligini tekshirish zarur.",
            "Tekshirish ro'yxati (checklist): titul varag'i to'g'rimi, mundarija sahifa raqamlariga mosmi, barcha havolalar ro'yxatda bormi, jadval va rasmlar raqamlanganmi, imlo tekshirilganmi, hajm talabga mosmi va format talablari bajarilganmi.",
            "Topshirishdan oldin: plagiat tekshiruvidan o'tkazish (Turnitin, StrikePlagiarism, Antiplagiat), formatni yakuniy tekshirish, PDF formatda saqlash (formatni saqlab qolish uchun) va zaxira nusxa yaratish. Muddatdan oldin topshirish — kutilmagan muammolarni hal qilish uchun vaqt qoldiradi.",
        ]),
    ],
})

# ===================== ISHGA TUSHIRISH =====================
if __name__ == "__main__":
    write_docx("/projects/sandbox/konspekt-30-mavzu/Akademik_yozuv_10_mavzu.docx")

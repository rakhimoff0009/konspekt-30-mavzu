#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mustaqil ish: Nutqini o'stirish metodikasi fanining ilmiy-nazariy asoslari
Target: 9 A4 pages, Times New Roman 14pt, 1.5 line spacing
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT = "/projects/sandbox/konspekt-30-mavzu/Nutq_metodikasi_mustaqil_ish.docx"

PLAN = [
    "Kirish",
    "Nutqini o'stirish metodikasi fanining predmeti va vazifalari",
    "Nutqini o'stirish metodikasining ilmiy asoslari",
    "Nutqini o'stirish metodikasining nazariy manbalari",
    "Bola nutqining rivojlanish bosqichlari",
    "Nutqini o'stirish metodikasining boshqa fanlar bilan aloqasi",
    "Nutqini o'stirish metodikasida qo'llaniladigan tadqiqot usullari",
    "Nutqini o'stirish metodikasining amaliy ahamiyati",
    "Xulosa va foydalanilgan adabiyotlar",
]

SECTIONS = [
    # 1. Kirish
    [
        "Nutq — insonning eng muhim muloqot vositasi bo'lib, u tafakkur, "
        "bilim olish va ijtimoiy hayotda faoliyat ko'rsatishning asosiy "
        "qurollaridan biridir. Bola hayotining dastlabki yillarida nutqning "
        "shakllanishi va rivojlanishi uning keyingi hayotidagi muvaffaqiyatlar "
        "uchun mustahkam poydevor yaratadi. Aynan shu sababli maktabgacha "
        "yoshdagi bolalarning nutqini o'stirish pedagogika fanining eng muhim "
        "yo'nalishlaridan biri sifatida e'tirof etiladi.",

        "Nutqini o'stirish metodikasi — bu bolalar nutqini shakllantirish va "
        "rivojlantirishning maqsadlari, mazmuni, usullari va vositalarini "
        "o'rganuvchi fan sohasi hisoblanadi. U pedagogika, psixologiya, "
        "tilshunoslik va fiziologiya fanlarining kesishuvida joylashgan bo'lib, "
        "har bir fanning ilmiy yutuqlaridan foydalanadi. Bu fan bolalarning "
        "nutqiy rivojlanishini ilmiy asosda tashkil etish imkonini beradi.",

        "Mazkur mustaqil ishda nutqini o'stirish metodikasi fanining ilmiy "
        "va nazariy asoslari batafsil yoritiladi. Ishda fanning predmeti, "
        "vazifalari, ilmiy manbalari, bola nutqining rivojlanish bosqichlari, "
        "boshqa fanlar bilan aloqasi va tadqiqot usullari tahlil qilinadi. "
        "Mavzuning dolzarbligi shundaki, bugungi kunda bolalar nutqining "
        "rivojlanishida turli muammolar kuzatilmoqda va ularni hal etish "
        "uchun ilmiy asoslangan yondashuvlar zarur.",
    ],

    # 2. Predmeti va vazifalari
    [
        "Nutqini o'stirish metodikasi fanining predmeti — bu maktabgacha "
        "yoshdagi bolalarning nutqiy malakalarini shakllantirish va "
        "rivojlantirish jarayonining qonuniyatlari, mazmuni, shakllari va "
        "usullarini o'rganishdir. Fan bolaning tug'ilgan kunidan boshlab "
        "maktabga borgunga qadar bo'lgan davrdagi nutqiy rivojlanishini "
        "qamrab oladi. U bolaning tovush talaffuzidan tortib, bog'lanishli "
        "nutq ko'nikmalarini shakllantirishgacha bo'lgan barcha jihatlarni "
        "o'z ichiga oladi.",

        "Nutqini o'stirish metodikasining asosiy vazifalari quyidagilardan "
        "iborat. Birinchi vazifa — bolalarning nutqiy rivojlanishi "
        "qonuniyatlarini o'rganish va ularni ta'limiy jarayonda hisobga "
        "olish. Ikkinchi vazifa — nutqni rivojlantirishning samarali "
        "usullari va vositalarini ishlab chiqish. Uchinchi vazifa — "
        "nutqiy mashg'ulotlarning mazmuni va tizimini aniqlash.",

        "To'rtinchi vazifa — har bir yosh bosqichiga mos nutqiy "
        "ko'nikmalar standartlarini belgilash. Beshinchi vazifa — "
        "nutqiy rivojlanishni baholash mezonlari va diagnostika "
        "usullarini ishlab chiqish. Oltinchi vazifa — nutqi rivojlanishida "
        "muammolari bo'lgan bolalar bilan ishlash yo'llarini aniqlash. "
        "Yettinchi vazifa — oila va maktabgacha ta'lim muassasasi "
        "hamkorligida bola nutqini rivojlantirish tizimini yaratish.",

        "Nutqini o'stirish metodikasi fani bir nechta muhim savollarga "
        "javob beradi: nimani o'rgatish kerak (mazmun), qanday o'rgatish "
        "kerak (usullar va shakllar), nima uchun aynan shunday o'rgatish "
        "kerak (ilmiy asos), qachon o'rgatish kerak (yosh bosqichlari) "
        "va qanday natijaga erishish kerak (standartlar). Bu savollar "
        "fanning asosiy kategoriyalarini tashkil etadi va ta'limiy "
        "jarayonni ilmiy asosda tashkil etish imkonini beradi.",
    ],

    # 3. Ilmiy asoslari
    [
        "Nutqini o'stirish metodikasining ilmiy asoslari bir nechta "
        "fundamental fan yutuqlariga tayanadi. Birinchi va eng muhim "
        "ilmiy asos — tilshunoslik (lingvistika). Tilshunoslik fani "
        "tilning tuzilishi, qonuniyatlari, fonetik, leksik, grammatik "
        "va sintaktik tizimlari haqida bilim beradi. Nutqini o'stirish "
        "metodikasi tilshunoslikning til haqidagi nazariy bilimlarini "
        "amaliy pedagogik faoliyatda qo'llaydi.",

        "Tilshunoslikning fonetika bo'limi bolaning tovush talaffuzini "
        "o'rgatish uchun asos bo'ladi. Leksikologiya bo'limi lug'at "
        "boyligini oshirish yo'llarini belgilaydi. Grammatika bo'limi "
        "gap tuzish va morfologik ko'nikmalarni shakllantirish usullarini "
        "asoslaydi. Nutq madaniyati bo'limi bolaga to'g'ri, aniq va "
        "ifodali gapirish ko'nikmalarini o'rgatish yo'llarini ko'rsatadi.",

        "Ikkinchi muhim ilmiy asos — psixologiya fani. Psixologiya "
        "nutqning psixik jarayon sifatidagi xususiyatlarini, nutq va "
        "tafakkur orasidagi aloqani, nutqning shakllanish mexanizmlarini "
        "o'rganadi. L.S. Vigotskiyning nutq va tafakkur munosabatlari "
        "haqidagi nazariyasi, A.N. Leontyevning faoliyat nazariyasi, "
        "A.R. Luriyaning neyrolingvistik tadqiqotlari nutqini o'stirish "
        "metodikasi uchun muhim nazariy asos bo'lib xizmat qiladi.",

        "Uchinchi ilmiy asos — fiziologiya va neyrofiziologiya fanlari. "
        "I.P. Pavlovning oliy nerv faoliyati haqidagi ta'limoti, ikkinchi "
        "signal tizimi to'g'risidagi nazariyasi nutqining fiziologik "
        "mexanizmlarini tushunishga yordam beradi. Bosh miya po'stlog'ining "
        "nutq markazlari — Broka va Vernike markazlari — bolaning "
        "nutqiy rivojlanishi qonuniyatlarini tushuntiradi. Bu bilimlar "
        "nutq buzilishlarini aniqlash va bartaraf etishda muhim rol o'ynaydi.",

        "To'rtinchi ilmiy asos — pedagogika fani. Didaktika tamoyillari, "
        "ta'lim qonuniyatlari, tarbiya nazariyasi nutqini o'stirish "
        "metodikasining pedagogik asosini tashkil etadi. Bolaga nutqni "
        "o'rgatishda ko'rgazmalilik, izchillik, tizimlilik, onglilik, "
        "individual yondashuv kabi didaktik tamoyillarga amal qilinadi. "
        "Shuningdek, ta'limning yosh xususiyatlariga mosligi tamoyili "
        "nutqiy mashg'ulotlar mazmunini belgilashda asosiy mezon bo'ladi.",
    ],

    # 4. Nazariy manbalari
    [
        "Nutqini o'stirish metodikasining nazariy manbalari bir nechta "
        "muhim ilmiy yo'nalishlarni o'z ichiga oladi. Birinchi manba — "
        "bolalar nutqini o'rganish tarixi. Bolalar nutqini ilmiy jihatdan "
        "o'rganish XIX asr oxirlarida boshlangan. Dastlab ota-onalarning "
        "kuzatuv kundaliklari, keyinchalik maxsus eksperimental tadqiqotlar "
        "orqali bolalar nutqining rivojlanish qonuniyatlari aniqlangan.",

        "K.D. Ushinskiy rus pedagogikasida bolalar nutqini o'stirish "
        "metodikasining asoschisi hisoblanadi. U ona tilini o'rgatishning "
        "muhimligini, ko'rgazmali ta'lim orqali nutq rivojlantirishni, "
        "xalq og'zaki ijodi vositalaridan foydalanishni asoslagan. "
        "Ye.I. Tixeyeva maktabgacha yoshdagi bolalar nutqini "
        "rivojlantirishning izchil metodikasini yaratgan birinchi olima "
        "bo'lib, uning ishlari bugungi kungacha ahamiyatini yo'qotmagan.",

        "A.M. Leushina bog'lanishli nutqni rivojlantirish metodikasi "
        "ustida katta ilmiy ish olib borgan. U bolaning situativ nutqdan "
        "kontekstli nutqga o'tish qonuniyatlarini aniqlagan. O.S. Ushakova "
        "maktabgacha yoshdagi bolalar nutqini rivojlantirishning yaxlit "
        "tizimini yaratgan va nutqiy tarbiyaning mazmuni, usullari va "
        "shakllarini ishlab chiqqan. Bu olimlarning ilmiy merosi "
        "bugungi metodikaning nazariy poydevori hisoblanadi.",

        "O'zbek pedagogikasida ham nutqini o'stirish metodikasi sohasida "
        "muhim tadqiqotlar olib borilgan. M. Sodiqova, Sh. Shoumarova, "
        "D. Raximova va boshqa olimlar o'zbek bolalarining nutqiy "
        "rivojlanish xususiyatlarini o'rganib, milliy til xususiyatlarini "
        "hisobga olgan metodikalar ishlab chiqishgan. Bu tadqiqotlar "
        "O'zbekiston maktabgacha ta'lim tizimida nutqiy tarbiya "
        "mazmunini belgilashda muhim manba bo'lib xizmat qilmoqda.",

        "Zamonaviy davrda nutqini o'stirish metodikasining nazariy "
        "manbalari yanada kengaymoqda. Psixolingvistika, neyrolingvistika, "
        "kognitiv tilshunoslik, pragmatika kabi yangi fan yo'nalishlari "
        "bolalar nutqini o'rganishga yangi yondashuvlar olib kelmoqda. "
        "Xususan, bola nutqining kommunikativ funksiyasi, pragmatik "
        "ko'nikmalar, metalingvistik ongli nutq haqidagi tadqiqotlar "
        "metodikaning nazariy bazasini boyitmoqda.",
    ],

    # 5. Nutq rivojlanish bosqichlari
    [
        "Bola nutqining rivojlanishi muayyan bosqichlar orqali amalga "
        "oshadi. Birinchi bosqich — tayyorlov davri (0-1 yosh). Bu davrda "
        "bola atrofdagi nutqni idrok etishni o'rganadi. U tovushlar "
        "farqiga boradi, kattalarning intonatsiyasini tushunadi, o'z "
        "ovozi bilan turli tovushlar chiqaradi. Gukking (2-3 oylik), "
        "babbling yoki babillash (5-6 oylik) kabi nutqgacha bo'lgan "
        "bosqichlar kuzatiladi. 10-12 oylikda birinchi so'zlar paydo bo'ladi.",

        "Ikkinchi bosqich — erta nutq davri (1-3 yosh). Bu davrda bolaning "
        "lug'ati tez sur'atda o'sadi. 1 yoshda 10-15 so'z bo'lsa, 2 yoshga "
        "kelib 200-300 so'zga, 3 yoshga kelib 1000-1500 so'zga yetadi. "
        "Bola avval bir so'zli gaplar (golofrazlar), keyin ikki so'zli "
        "gaplar, so'ng sodda gaplar tuzishni o'rganadi. Bu davrda "
        "grammatik tuzilmalar shakllanishi boshlanadi va bola oddiy "
        "savollarga javob bera oladi.",

        "Uchinchi bosqich — maktabgacha yosh nutqi (3-5 yosh). Bu davrda "
        "bolaning nutqi sifat jihatidan keskin o'zgaradi. Lug'at boyligi "
        "2000-3000 so'zga yetadi, gap tuzilishi murakkablashadi, ergash "
        "gapli qo'shma gaplar paydo bo'ladi. Bola voqealarni hikoya qilib "
        "bera oladi, savollarga batafsil javob beradi, oddiy munozaraga "
        "kirishadi. Tovush talaffuzi ko'p jihatdan to'g'rilanadi, lekin "
        "ba'zi tovushlar (r, l, sh) hali qiyinchilik tug'dirishi mumkin.",

        "To'rtinchi bosqich — katta maktabgacha yosh nutqi (5-7 yosh). "
        "Bu bosqichda nutq rivojlanishining asosiy ko'rsatkichlari "
        "shakllanadi. Lug'at boyligi 3500-4000 so'zga yetadi, bola "
        "ko'chma ma'noli so'zlarni tushunadi, frazeologizmlarni "
        "ishlata boshlaydi. Bog'lanishli nutq to'liq shakllanadi — "
        "bola hikoya tuzadi, ertak aytib beradi, o'z fikrini asoslab "
        "gapiradi. Tovush talaffuzi to'liq normaga keladi.",

        "Har bir bosqichda nutqning turli tomonlari — fonetik, leksik, "
        "grammatik va bog'lanishli nutq — parallel ravishda rivojlanadi. "
        "Nutqini o'stirish metodikasi har bir bosqichning xususiyatlarini "
        "hisobga olib, ta'limiy jarayonning mazmuni va usullarini "
        "belgilaydi. Agar biror bosqichda nutqiy rivojlanish kechiksa, "
        "bu keyingi bosqichlarga ham salbiy ta'sir ko'rsatadi. Shu "
        "sababli har bir bosqichda o'z vaqtida pedagogik ta'sir "
        "ko'rsatish muhim ahamiyatga ega.",
    ],

    # 6. Boshqa fanlar bilan aloqasi
    [
        "Nutqini o'stirish metodikasi mustaqil fan sifatida boshqa ko'plab "
        "fanlar bilan chambarchas bog'liq. Bu aloqadorlik fanning "
        "interdistsiplinar xarakterini belgilaydi. Avvalo, u tilshunoslik "
        "fani bilan uzviy bog'liq. Tilshunoslikning fonetika, leksikologiya, "
        "grammatika, uslubiyat va nutq madaniyati bo'limlari nutqini "
        "o'stirish metodikasining lingvistik asosini tashkil etadi.",

        "Psixologiya fani bilan aloqasi ham juda mustahkam. Umumiy "
        "psixologiya nutq jarayonining mexanizmlarini, bolalar psixologiyasi "
        "esa turli yosh davrlarida nutqning shakllanish xususiyatlarini "
        "tushuntirib beradi. Pedagogik psixologiya ta'lim jarayonida "
        "nutqiy ko'nikmalarning shakllanishi qonuniyatlarini o'rgatadi. "
        "Psixolingvistika esa nutq ishlab chiqarish va nutqni idrok etish "
        "modellarini taqdim etadi.",

        "Anatomiya va fiziologiya fanlari nutq a'zolarining tuzilishi, "
        "artikulyatsion apparatning ishlash mexanizmi, eshitish "
        "analizatorining funksiyasi haqida ma'lumot beradi. Bu bilimlar "
        "bolaning tovush talaffuzini shakllantirish, nutq nafasini "
        "rivojlantirish va nutq buzilishlarini aniqlashda zarurdir. "
        "Neyrofiziologiya bosh miyaning nutq markazlari faoliyatini "
        "tushuntirib, nutqiy rivojlanish kechikishlarini baholashga "
        "yordam beradi.",

        "Logopediya fani bilan aloqasi alohida ahamiyatga ega. Logopediya "
        "nutq buzilishlarini aniqlash, bartaraf etish va oldini olish "
        "bilan shug'ullanadi. Nutqini o'stirish metodikasi esa sog'lom "
        "bolalarning nutqiy rivojlanishini ta'minlaydi. Bu ikki fan "
        "birgalikda bolaning nutqiy kamolotini ta'minlashga xizmat qiladi. "
        "Tarbiyachi logopedik bilimlardan xabardor bo'lishi va zarurat "
        "tug'ilganda mutaxassisga yo'naltira olishi kerak.",

        "Maktabgacha pedagogika va maktabgacha ta'lim nazariyasi bilan "
        "ham bevosita bog'liq. Nutq rivojlantirish maktabgacha ta'limning "
        "ajralmas qismi bo'lib, u jismoniy, aqliy, ahloqiy va estetik "
        "tarbiya bilan uzviy bog'langan. Bola jismoniy mashqlar "
        "vaqtida nafas olish va artikulyatsion apparatni mashq qildiradi, "
        "aqliy tarbiya jarayonida lug'ati boyiydi, ahloqiy suhbatlarda "
        "bog'lanishli nutqi rivojlanadi, estetik tarbiyada badiiy so'z "
        "san'atini o'rganadi.",
    ],

    # 7. Tadqiqot usullari
    [
        "Nutqini o'stirish metodikasida ilmiy tadqiqotlar olib borish "
        "uchun turli usullardan foydalaniladi. Bu usullar pedagogik, "
        "psixologik va lingvistik tadqiqot usullarining sintezi "
        "hisoblanadi. Birinchi va eng keng tarqalgan usul — kuzatish "
        "usuli. Tarbiyachi yoki tadqiqotchi bolaning nutqiy faoliyatini "
        "tabiiy sharoitda kuzatib, nutqiy rivojlanish darajasi, "
        "xususiyatlari va muammolarini aniqlaydi.",

        "Kuzatish usuli tashqi va ichki turlariga bo'linadi. Tashqi "
        "kuzatishda tadqiqotchi bola faoliyatiga aralashmasdan uning "
        "nutqini kuzatadi va qayd etadi. Ishtirokchi kuzatishda esa "
        "tadqiqotchi bola bilan muloqotga kirishib, uning nutqiy "
        "imkoniyatlarini ochishga harakat qiladi. Kuzatish natijalari "
        "maxsus protokollarda, kundaliklarda yoki texnik vositalar "
        "yordamida qayd etiladi.",

        "Ikkinchi usul — eksperiment. Pedagogik eksperiment nutqini "
        "o'stirish metodikasida eng muhim tadqiqot usuli hisoblanadi. "
        "U aniqlash eksperimenti va shakllantirish eksperimentiga "
        "bo'linadi. Aniqlash eksperimentida bolalarning mavjud nutqiy "
        "rivojlanish darajasi o'lchanadi. Shakllantirish eksperimentida "
        "esa yangi metod yoki vosita sinab ko'riladi va uning "
        "samaradorligi baholanadi.",

        "Uchinchi usul — suhbat va intervyu. Tarbiyachi bola bilan "
        "maxsus tuzilgan savollar asosida individual suhbat o'tkazib, "
        "uning nutqiy rivojlanishini baholaydi. Bu usul bolaning lug'at "
        "boyligini, grammatik to'g'ri gapirishini, bog'lanishli nutq "
        "darajasini aniqlashda samarali. Shuningdek, ota-onalar bilan "
        "suhbat bolaning oiladagi nutqiy muhiti haqida ma'lumot beradi.",

        "To'rtinchi usul — test va diagnostika usullari. Maxsus "
        "ishlab chiqilgan testlar yordamida bolaning nutqiy "
        "rivojlanishi standart me'yorlarga muvofiqlik darajasi "
        "aniqlanadi. Beshinchi usul — bolalar nutqi namunalarini "
        "tahlil qilish. Bolaning erkin nutqi yozib olinib, fonetik, "
        "leksik, grammatik va pragmatik jihatdan tahlil qilinadi. "
        "Oltinchi usul — hujjatlar tahlili. Tarbiyachilarning ish "
        "rejalari, mashg'ulot ishlanmalari va bolalar ishlari tahlil "
        "qilinib, nutqiy tarbiya holati baholanadi.",
    ],

    # 8. Amaliy ahamiyati
    [
        "Nutqini o'stirish metodikasi fanining amaliy ahamiyati juda "
        "keng qamrovli bo'lib, u bir nechta muhim yo'nalishlarni qamrab "
        "oladi. Birinchidan, bu fan maktabgacha ta'lim muassasalarida "
        "bolalarning nutqiy rivojlanishini ilmiy asosda tashkil etish "
        "imkonini beradi. Tarbiyachi fanning nazariy bilimlarini "
        "egallagan holda, mashg'ulotlarni samarali rejalashtiradi, "
        "to'g'ri usullar tanlaydi va bolaning nutqiy kamolotiga "
        "maqsadli ta'sir ko'rsatadi.",

        "Ikkinchidan, nutqini o'stirish metodikasi nutqiy rivojlanishda "
        "muammolari bo'lgan bolalarni erta aniqlash va ularga o'z vaqtida "
        "yordam ko'rsatish imkonini beradi. Tarbiyachi fanning diagnostika "
        "usullarini egallagan holda, bolaning nutqiy rivojlanishi me'yorga "
        "mos kelmasligi alomatlarini payqab, ota-onalarga va "
        "mutaxassislarga o'z vaqtida murojaat qilishni maslahat bera "
        "oladi.",

        "Uchinchidan, bu fan nutqiy mashg'ulotlarning samarali shakllari "
        "va usullarini ishlab chiqish uchun ilmiy asos yaratadi. Har "
        "bir yosh guruhiga mos, bolalarning qiziqishlarini hisobga "
        "olgan, o'yin elementlari bilan boyitilgan mashg'ulotlar "
        "yaratiladi. Bu mashg'ulotlar bolaning nutqiy ko'nikmalarini "
        "tabiiy va quvonchli tarzda rivojlantirishga xizmat qiladi.",

        "To'rtinchidan, nutqini o'stirish metodikasi oilada bola "
        "nutqini rivojlantirish bo'yicha ilmiy asoslangan tavsiyalar "
        "ishlab chiqadi. Ota-onalar uchun bolaning yosh xususiyatlariga "
        "mos nutqiy o'yinlar, mashqlar va faoliyat turlari tavsiya "
        "etiladi. Bu oila va maktabgacha ta'lim muassasasi o'rtasidagi "
        "hamkorlikni mustahkamlaydi va bola nutqining uzluksiz "
        "rivojlanishini ta'minlaydi.",

        "Beshinchidan, bu fan maktabga tayyorlov dasturlari uchun "
        "muhim ma'lumot manbai hisoblanadi. Bolaning maktabga "
        "tayyorligi ko'p jihatdan uning nutqiy rivojlanish darajasiga "
        "bog'liq. Nutqini o'stirish metodikasi bolada maktab ta'limiga "
        "zarur nutqiy ko'nikmalar — savodxonlikka tayyorgarlik, "
        "monologik nutq, hikoya tuzish, qayta hikoyalash, munozara "
        "yuritish qobiliyatlarini shakllantirish yo'llarini belgilaydi.",

        "Oltinchidan, nutqini o'stirish metodikasi tarbiyachilarning "
        "kasbiy tayyorgarligini oshirishga xizmat qiladi. Pedagogika "
        "oliy ta'lim muassasalarida ushbu fan alohida o'quv kursi "
        "sifatida o'qitiladi. Kelajakdagi tarbiyachilar nutqiy "
        "tarbiyaning nazariy asoslari va amaliy usullarini puxta "
        "o'rganadilar va amaliyotda samarali qo'llay oladilar.",
    ],

    # 9. Xulosa
    [
        "Yuqorida bayon etilgan ma'lumotlar shundan dalolat beradiki, "
        "nutqini o'stirish metodikasi — bu mustaqil pedagogik fan bo'lib, "
        "u tilshunoslik, psixologiya, fiziologiya va pedagogika "
        "fanlarining ilmiy yutuqlariga tayanadi. Fanning predmeti — "
        "maktabgacha yoshdagi bolalarning nutqiy rivojlanish "
        "qonuniyatlarini o'rganish va nutqiy tarbiyaning samarali "
        "usullarini ishlab chiqishdir.",

        "Nutqini o'stirish metodikasining ilmiy-nazariy asoslari boy "
        "va ko'p qirrali bo'lib, ular ko'p asrlik pedagogik tajriba, "
        "psixologik va lingvistik tadqiqotlar natijasida shakllangan. "
        "K.D. Ushinskiy, Ye.I. Tixeyeva, A.M. Leushina, O.S. Ushakova "
        "va boshqa ko'plab olimlarning ilmiy merosi fanning nazariy "
        "poydevorini tashkil etadi.",

        "Bola nutqining rivojlanishi muayyan bosqichlar orqali amalga "
        "oshib, har bir bosqichda nutqning fonetik, leksik, grammatik "
        "va bog'lanishli tomonlari parallel rivojlanadi. Tarbiyachi bu "
        "bosqichlarni yaxshi bilishi va har bir bola uchun individual "
        "yondashuvni amalga oshirishi muhim. Shu tariqa bolaning nutqiy "
        "kamoloti to'liq ta'minlanadi.",

        "Xulosa qilib aytganda, nutqini o'stirish metodikasi fanining "
        "ilmiy-nazariy asoslarini chuqur o'rganish har bir tarbiyachi "
        "uchun zaruratdir. Bu bilimlar tarbiyachiga bolalar nutqini "
        "samarali rivojlantirish, nutqiy muammolarni o'z vaqtida "
        "aniqlash va professional faoliyat yuritish imkonini beradi. "
        "Fanning amaliy ahamiyati keng bo'lib, u maktabgacha ta'lim "
        "sifatini oshirishga bevosita xizmat qiladi.",

        "Foydalanilgan adabiyotlar:",
        "1. O'zbekiston Respublikasining \"Maktabgacha ta'lim va tarbiya "
        "to'g'risida\"gi Qonuni. — Toshkent, 2020.",
        "2. Maktabgacha ta'limning Davlat o'quv dasturi \"Ilk qadam\". — "
        "Toshkent, 2018.",
        "3. Sodiqova M. \"Bolalar nutqini o'stirish metodikasi\". — "
        "Toshkent, 2019.",
        "4. Shoumarova M. \"Maktabgacha yoshdagi bolalar nutqini "
        "rivojlantirish\". — Toshkent, 2021.",
        "5. Ushakova O.S. \"Razvitiye rechi doshkolnikov\". — Moskva, 2017.",
        "6. Alekseyeva M.M., Yashina B.I. \"Metodika razvitiya rechi\". — "
        "Moskva, 2018.",
    ],
]


# ---------- DOCX XML BUILDER ----------

def make_para(text, bold=False, size=28, align=None, before=120, after=120):
    align_xml = f'<w:jc w:val="{align}"/>' if align else ""
    bold_xml = '<w:b/>' if bold else ""
    safe = escape(text)
    return (
        '<w:p><w:pPr>'
        f'<w:spacing w:before="{before}" w:after="{after}" w:line="360" w:lineRule="auto"/>'
        f'{align_xml}'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>{bold_xml}</w:rPr>'
        '</w:pPr><w:r><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>{bold_xml}</w:rPr>'
        f'<w:t xml:space="preserve">{safe}</w:t></w:r></w:p>'
    )

def make_body_para(text):
    safe = escape(text)
    return (
        '<w:p><w:pPr>'
        '<w:spacing w:before="0" w:after="120" w:line="360" w:lineRule="auto"/>'
        '<w:ind w:firstLine="709"/>'
        '<w:jc w:val="both"/>'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        '</w:pPr><w:r><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        f'<w:t xml:space="preserve">{safe}</w:t></w:r></w:p>'
    )

def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


# Build body
body = []

# TITLE PAGE
for _ in range(4):
    body.append(make_para("", size=28))
body.append(make_para("O'ZBEKISTON RESPUBLIKASI", bold=True, size=28, align="center"))
body.append(make_para("MAKTABGACHA VA MAKTAB TA'LIMI VAZIRLIGI", bold=True, size=28, align="center"))
body.append(make_para("", size=24))
body.append(make_para("MUSTAQIL ISH", bold=True, size=40, align="center"))
body.append(make_para("", size=24))
body.append(make_para("Mavzu:", bold=True, size=28, align="center"))
body.append(make_para(
    "Nutqini o'stirish metodikasi fanining ilmiy-nazariy asoslari",
    bold=True, size=28, align="center"))
for _ in range(6):
    body.append(make_para("", size=24))
body.append(make_para("Bajardi: ______________________", size=28, align="right"))
body.append(make_para("Tekshirdi: ______________________", size=28, align="right"))
for _ in range(2):
    body.append(make_para("", size=24))
body.append(make_para("Toshkent — 2026", bold=True, size=28, align="center"))
body.append(page_break())

# PLAN PAGE
body.append(make_para("REJA", bold=True, size=32, align="center"))
body.append(make_para("", size=20))
for i, item in enumerate(PLAN, 1):
    body.append(make_para(f"{i}. {item}", size=28))
body.append(page_break())

# CONTENT PAGES
for idx, (heading, paragraphs) in enumerate(zip(PLAN, SECTIONS)):
    body.append(make_para(f"{idx + 1}. {heading.upper()}", bold=True, size=30, align="center"))
    body.append(make_para("", size=18))
    for p in paragraphs:
        body.append(make_body_para(p))
    if idx < len(SECTIONS) - 1:
        body.append(page_break())

sectPr = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" '
    'w:header="708" w:footer="708" w:gutter="0"/>'
    '<w:cols w:space="708"/>'
    '<w:docGrid w:linePitch="360"/>'
    '</w:sectPr>'
)

document_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>' + ''.join(body) + sectPr + '</w:body></w:document>'
)

content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

document_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
<w:sz w:val="28"/><w:szCs w:val="28"/>
<w:lang w:val="uz-UZ"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr>
<w:spacing w:line="360" w:lineRule="auto"/>
</w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
<w:name w:val="Normal"/><w:qFormat/>
</w:style></w:styles>'''

core_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:title>Mustaqil ish - Nutqini o stirish metodikasi</dc:title>
<dc:creator>Talaba</dc:creator>
</cp:coreProperties>'''

app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>Microsoft Office Word</Application>
</Properties>'''

# Write DOCX
os.makedirs(os.path.dirname(OUT), exist_ok=True)
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', content_types_xml)
    z.writestr('_rels/.rels', rels_xml)
    z.writestr('word/_rels/document.xml.rels', document_rels_xml)
    z.writestr('word/document.xml', document_xml)
    z.writestr('word/styles.xml', styles_xml)
    z.writestr('docProps/core.xml', core_xml)
    z.writestr('docProps/app.xml', app_xml)

print(f"Created: {OUT}")
print(f"Size: {os.path.getsize(OUT)} bytes")

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mustaqil ish: Bolalarni nutqini o'stirish fanining maqsadi va vazifasi
Target: 9 A4 pages, Times New Roman 14pt, 1.5 line spacing
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT = "/projects/sandbox/konspekt-30-mavzu/Bolalar_nutqi_maqsad_vazifa.docx"

PLAN = [
    "Kirish",
    "Bolalarni nutqini o'stirish fanining mohiyati",
    "Bolalarni nutqini o'stirish fanining maqsadi",
    "Bolalarni nutqini o'stirish fanining vazifalari",
    "Nutqiy tarbiyaning asosiy yo'nalishlari",
    "Nutqini o'stirish fanining maktabgacha ta'limdagi o'rni",
    "Nutqini o'stirish maqsad va vazifalarining yosh guruhlariga bog'liqligi",
    "Nutqini o'stirish fanining zamonaviy tendensiyalari",
    "Xulosa va foydalanilgan adabiyotlar",
]

SECTIONS = [
    # 1. Kirish
    [
        "Nutq insonni boshqa tirik mavjudotlardan ajratib turadigan eng muhim "
        "xususiyatlardan biridir. Nutq orqali inson o'z fikrlarini ifodalaydi, "
        "atrofdagilar bilan muloqot qiladi, bilim oladi va ijtimoiy hayotda "
        "faol ishtirok etadi. Bolaning nutqiy rivojlanishi uning butun shaxsiy "
        "kamolotining asosiy ko'rsatkichlaridan biri hisoblanadi. Shu sababli "
        "maktabgacha ta'lim tizimida bolalarning nutqini o'stirish eng muhim "
        "pedagogik vazifalardan biri sifatida e'tirof etiladi.",

        "Bolalarni nutqini o'stirish fani — bu maktabgacha yoshdagi bolalarning "
        "nutqiy malakalarini shakllantirish va rivojlantirishning ilmiy "
        "asoslarini, maqsadlarini, vazifalarini, usullarini va vositalarini "
        "o'rganuvchi pedagogik fan sohasidir. Bu fan tarbiyachilarga bolalarning "
        "nutqiy kamolotini ilmiy asosda tashkil etish, samarali usullar qo'llash "
        "va har bir bolaning nutqiy ehtiyojlariga mos yondashish imkonini beradi.",

        "Mazkur mustaqil ishda bolalarni nutqini o'stirish fanining maqsadi va "
        "vazifalari batafsil yoritiladi. Ishda fanning mohiyati, asosiy maqsadi, "
        "vazifalari tizimi, nutqiy tarbiyaning yo'nalishlari, maktabgacha "
        "ta'limdagi o'rni va zamonaviy tendensiyalari tahlil qilinadi. Mavzuning "
        "dolzarbligi shundaki, bugungi kunda bolalarning nutqiy rivojlanishiga "
        "bo'lgan talab ortmoqda va bu sohada aniq maqsadlar va vazifalar "
        "belgilash zaruriyati kuchaymoqda.",
    ],

    # 2. Fanning mohiyati
    [
        "Bolalarni nutqini o'stirish fani pedagogika fanlarining muhim "
        "tarmog'i bo'lib, u maktabgacha yoshdagi bolalarning og'zaki nutqini "
        "shakllantirish va rivojlantirish jarayonini ilmiy jihatdan o'rganadi. "
        "Fanning mohiyati shundaki, u bolaning nutqiy rivojlanishini tasodifiy "
        "jarayonga emas, balki maqsadli, rejalashtirilgan va ilmiy asoslangan "
        "pedagogik faoliyatga aylantiradi.",

        "Bu fan tilshunoslik, psixologiya, pedagogika va fiziologiya "
        "fanlarining tutashgan nuqtasida joylashgan bo'lib, har bir fanning "
        "ilmiy yutuqlaridan foydalanadi. Tilshunoslikdan til tizimi va nutq "
        "qoidalari, psixologiyadan nutqning shakllanish mexanizmlari, "
        "pedagogikadan ta'lim tamoyillari va usullari, fiziologiyadan nutq "
        "a'zolarining rivojlanish xususiyatlari olinadi va pedagogik "
        "amaliyotga tatbiq etiladi.",

        "Fanning o'rganish ob'ekti — maktabgacha yoshdagi bolalarning nutqiy "
        "rivojlanish jarayonidir. Fanning predmeti esa shu jarayonni boshqarish, "
        "ya'ni bolalar nutqini maqsadli ravishda shakllantirish va "
        "rivojlantirishning pedagogik shart-sharoitlari, mazmuni, usullari va "
        "shakllari hisoblanadi. Fan bolaning tug'ilganidan maktabga borguncha "
        "bo'lgan barcha nutqiy rivojlanish bosqichlarini qamrab oladi.",

        "Bolalarni nutqini o'stirish fanining mohiyatini tushunish uchun nutq "
        "va til tushunchalarini farqlash lozim. Til — bu ijtimoiy hodisa bo'lib, "
        "muayyan qoidalar tizimidan iborat. Nutq esa tilning individual "
        "qo'llanilishi, ya'ni muayyan shaxsning muloqot jarayonidagi faoliyatidir. "
        "Fan bolaga tilni o'rgatish orqali uning nutqiy faoliyatini "
        "shakllantiradi — ya'ni bolani to'g'ri, ravon, ifodali va mazmunli "
        "gapira olishga o'rgatadi.",
    ],

    # 3. Maqsadi
    [
        "Bolalarni nutqini o'stirish fanining bosh maqsadi — maktabgacha "
        "yoshdagi bolalarda ona tilida erkin muloqot qilish qobiliyatini "
        "shakllantirish, ularning nutqiy malakalarini tizimli ravishda "
        "rivojlantirish va maktab ta'limiga tayyorlashdir. Bu maqsad bolaning "
        "butun maktabgacha davridagi nutqiy tarbiyasining umumiy yo'nalishini "
        "belgilaydi.",

        "Fanning maqsadi bir nechta muhim jihatlarni o'z ichiga oladi. "
        "Birinchi jihat — bolada to'g'ri nutq ko'nikmalarini shakllantirish. "
        "Bu degani, bola ona tilidagi barcha tovushlarni to'g'ri talaffuz "
        "qilishi, so'zlarni grammatik jihatdan to'g'ri qo'llashi, gaplarni "
        "mantiqiy to'g'ri tuzishi va o'z fikrini aniq ifodalashi kerak.",

        "Ikkinchi jihat — bolada bog'lanishli nutqni rivojlantirish. "
        "Bog'lanishli nutq deganda bolaning o'z fikrini izchil, mantiqiy va "
        "mazmunan to'liq ifodalay olish qobiliyati tushuniladi. Bola o'z "
        "boshidan o'tgan voqealarni hikoya qilib bera olishi, ertak aytishi, "
        "rasm bo'yicha hikoya tuza olishi va suhbatda faol ishtirok eta olishi "
        "kerak. Bu ko'nikma maktab ta'limida muvaffaqiyatning asosi hisoblanadi.",

        "Uchinchi jihat — bolada nutqiy muloqot madaniyatini tarbiyalash. "
        "Bola nafaqat to'g'ri gapirishni, balki madaniyatli muloqot olib "
        "borishni ham o'rganishi kerak. Bu suhbatdoshni diqqat bilan "
        "tinglash, so'zini bo'lmaslik, iltimos va minnatdorchilik bildirish, "
        "murojaat qilish odoblari va nutq etiketini o'z ichiga oladi.",

        "To'rtinchi jihat — bolada badiiy so'zga qiziqish uyg'otish. "
        "Bola adabiy asarlarni tinglash va tushunishga, she'rlarni yodlashga, "
        "ertaklarni qayta hikoya qilishga va badiiy nutq vositalarini "
        "o'zlashtirishga qiziqa boshlashi kerak. Bu bolaning nutqini "
        "boyitadi, estetik didini shakllantiradi va kitobxonlik "
        "madaniyatining asosini yaratadi. Fanning umumiy maqsadi bolani "
        "har tomonlama nutqiy jihatdan maktabga tayyor holga keltirishdir.",
    ],

    # 4. Vazifalari
    [
        "Bolalarni nutqini o'stirish fanining vazifalari maqsaddan kelib "
        "chiqqan holda aniq va tizimli tarzda belgilangan. Birinchi vazifa — "
        "tovush madaniyatini tarbiyalash. Bu vazifa bolaning nutq tovushlarini "
        "to'g'ri talaffuz qilishi, nutq nafasini boshqarishi, ovoz balandligi "
        "va tezligini muvofiqlashtirishi, intonatsiyadan to'g'ri foydalanishi "
        "ko'nikmalarini shakllantirishni o'z ichiga oladi. Har bir yosh "
        "bosqichida bu vazifa alohida mazmun kasb etadi.",

        "Ikkinchi vazifa — lug'at boyligini oshirish. Bu vazifa bolaning "
        "faol va passiv lug'atini muntazam kengaytirib borishni ko'zda "
        "tutadi. Bola yangi so'zlarni o'zlashtirishi, ularning ma'nosini "
        "tushunishi va nutqida to'g'ri qo'llashi kerak. Lug'at ishi "
        "bolaning bilim doirasi kengayishi bilan parallel olib boriladi — "
        "yangi predmetlar, hodisalar, harakatlar va sifatlar nomlari "
        "o'rgatiladi.",

        "Uchinchi vazifa — grammatik to'g'ri nutqni shakllantirish. "
        "Bu vazifa bolaga so'z yasash, so'z o'zgartirish, gap tuzish "
        "qoidalarini amaliy ko'nikmalar orqali o'rgatishni nazarda tutadi. "
        "Bola otlarni to'g'ri turlantirishni, fe'llarni tuslantirishni, "
        "sifat va ravishlardan foydalanishni, sodda va qo'shma gaplar "
        "tuzishni o'rganadi. Grammatik ko'nikmalar maxsus mashqlar va "
        "didaktik o'yinlar orqali shakllantiriladi.",

        "To'rtinchi vazifa — bog'lanishli nutqni rivojlantirish. Bu "
        "fanning eng murakkab va muhim vazifasi hisoblanadi. Bog'lanishli "
        "nutk ikki turga bo'linadi: dialogik (suhbat) va monologik (hikoya). "
        "Bola savollarga javob berish, suhbat yuritish, o'z fikrini bayon "
        "etish, voqealarni izchil hikoya qilish, tasvirlash va mulohaza "
        "yuritish ko'nikmalarini egallashi kerak.",

        "Beshinchi vazifa — badiiy adabiyot bilan tanishtirish. Bu vazifa "
        "bolani badiiy so'z san'ati bilan tanishtirishni, adabiy asarlarni "
        "tushunish va baholash ko'nikmalarini, she'rlarni ifodali aytish "
        "malakasini va ertaklarni dramatizatsiya qilish qobiliyatini "
        "shakllantirishni nazarda tutadi. Oltinchi vazifa — savodxonlikka "
        "tayyorlash. Maktabga tayyorlov guruhida bolaga tovush tahlili, "
        "harf bilan tanishish va dastlabki o'qish ko'nikmalarining "
        "asoslari o'rgatiladi.",
    ],

    # 5. Nutqiy tarbiyaning yo'nalishlari
    [
        "Nutqiy tarbiya bir nechta asosiy yo'nalishlar bo'yicha olib "
        "boriladi. Birinchi yo'nalish — fonetik yo'nalish. Bu yo'nalish "
        "bolaning nutq tovushlarini to'g'ri talaffuz etishini ta'minlashga "
        "qaratilgan. Har bir yosh davrida boladan muayyan tovushlarni "
        "egallashi kutiladi. Masalan, 3 yoshda oddiy tovushlar (m, n, p, b, t, d), "
        "4-5 yoshda murakkabroq tovushlar (s, z, sh, ch), 5-6 yoshda eng "
        "murakkab tovushlar (r, l) to'liq shakllanishi kerak.",

        "Fonetik yo'nalishda artikulyatsion gimnastika, nafas mashqlari, "
        "tovush o'yinlari, she'r va tez aytishlar keng qo'llaniladi. "
        "Tarbiyachi bolaning har bir tovushni to'g'ri aytishi ustida "
        "tizimli ishlaydi. Agar bolada nutq tovushlari buzilgan bo'lsa, "
        "tarbiyachi logoped bilan hamkorlikda maxsus mashqlar o'tkazadi.",

        "Ikkinchi yo'nalish — leksik yo'nalish. Bu yo'nalish bolaning "
        "lug'at boyligini oshirish va so'z ma'nolarini chuqur tushunishga "
        "o'rgatishga qaratilgan. Bola predmet nomlari, harakatlar, sifatlar, "
        "miqdorlar, vaqt tushunchalari va mavhum so'zlarni o'zlashtiradi. "
        "Leksik ish ko'rgazmali materiallar, ekskursiyalar, kuzatishlar, "
        "o'yinlar va suhbatlar orqali olib boriladi.",

        "Uchinchi yo'nalish — grammatik yo'nalish. Bu yo'nalish bolaga "
        "tilning grammatik qoidalarini amaliy tarzda o'rgatishga qaratilgan. "
        "Bola so'zlarni birlashtirish, gap tuzish, to'g'ri grammatik "
        "shakllardan foydalanish ko'nikmalarini o'zlashtiradi. Grammatik "
        "ishda maxsus didaktik o'yinlar, mashqlar, topshmoqlar va "
        "rasm bo'yicha gaplar tuzish usullari qo'llaniladi.",

        "To'rtinchi yo'nalish — bog'lanishli nutq yo'nalishi. Bu "
        "yo'nalish bolani izchil, mantiqiy va mazmunan to'liq fikr "
        "bildira olishga o'rgatishga qaratilgan. Dialogik nutq — suhbat "
        "qilish, savol berish va javob berish orqali, monologik nutq esa — "
        "hikoya tuzish, qayta hikoyalash, tasvirlash va mulohaza yuritish "
        "orqali rivojlantiriladi. Beshinchi yo'nalish — nutq ijodkorligi, "
        "ya'ni bolani o'z nutqida badiiy vositalardan foydalanishga, ertak "
        "to'qishga va she'riy nutqni his qilishga o'rgatish.",
    ],

    # 6. Maktabgacha ta'limdagi o'rni
    [
        "Bolalarni nutqini o'stirish fani maktabgacha ta'lim tizimida "
        "markaziy o'rinlardan birini egallaydi. Bu fanning maktabgacha "
        "ta'limdagi o'rni shundan iboratki, nutqiy rivojlanish bolaning "
        "barcha boshqa rivojlanish sohalariga bevosita ta'sir ko'rsatadi. "
        "Aqliy rivojlanish nutq orqali amalga oshadi — bola tafakkur "
        "qilish uchun so'zlardan foydalanadi. Ijtimoiy rivojlanish ham "
        "nutqqa bog'liq — bola muloqot qilish uchun nutqga muhtoj.",

        "Maktabgacha ta'lim muassasalarida nutqiy tarbiya alohida "
        "mashg'ulotlar shaklida ham, kundalik faoliyat jarayonida ham "
        "amalga oshiriladi. Maxsus nutqiy mashg'ulotlarda tarbiyachi "
        "maqsadli ravishda bolalarning nutqiy ko'nikmalarini "
        "shakllantiradi. Kundalik hayotda esa — kiyinish, ovqatlanish, "
        "sayr, o'yin vaqtida — bola nutqi tabiiy tarzda rivojlanadi.",

        "Davlat ta'lim standartlari va \"Ilk qadam\" dasturida nutqiy "
        "rivojlanish alohida ta'limiy soha sifatida ajratilgan. Dasturda "
        "har bir yosh guruhi uchun nutqiy rivojlanishning kutilgan "
        "natijalari, ya'ni bolaning qanday nutqiy ko'nikmalarga ega "
        "bo'lishi kerakligi aniq ko'rsatilgan. Tarbiyachi o'z ishini "
        "aynan shu standartlar asosida tashkil etadi.",

        "Nutqini o'stirish fani boshqa ta'limiy sohalar — matematik "
        "rivojlanish, tabiat bilan tanishtirish, badiiy-estetik "
        "rivojlanish, jismoniy tarbiya bilan integratsiya qilingan "
        "holda amalga oshiriladi. Masalan, matematik mashg'ulotda bola "
        "sanash jarayonida so'zlar qo'llaydi, tabiat mashg'ulotida "
        "o'simliklar va hayvonlar haqida hikoya tuzadi, badiiy ijod "
        "mashg'ulotida o'z ishini tasvirlab beradi.",

        "Nutqini o'stirish fanining maktabgacha ta'limdagi yana bir "
        "muhim o'rni — bolani maktab ta'limiga nutqiy jihatdan "
        "tayyorlashdir. Maktabga kirgan bola o'qituvchining nutqini "
        "tushunishi, o'z fikrini aniq ifodalay olishi, savollarga "
        "javob berishi, matn mazmunini qayta aytib berishi va yozishga "
        "tayyorgarlik ko'rishi kerak. Bu ko'nikmalarning barchasi "
        "maktabgacha davrda shakllantirila boshlanadi.",
    ],

    # 7. Yosh guruhlariga bog'liqligi
    [
        "Bolalarni nutqini o'stirish fanining maqsad va vazifalari har "
        "bir yosh guruhida o'ziga xos mazmun kasb etadi. Erta yosh "
        "guruhi (1-3 yosh) uchun asosiy maqsad — bolaning birinchi "
        "so'zlarini shakllantirish, passiv lug'atini boyitish va oddiy "
        "gaplar tuzishga o'rgatishdir. Bu yoshda asosiy vazifalar: "
        "artikulyatsion apparatni mustahkamlash, lug'atni kengaytirish, "
        "oddiy ko'rsatmalarni tushunishga o'rgatish va kattalar nutqiga "
        "taqlid qilishni rag'batlantirish.",

        "Kichik guruh (3-4 yosh) uchun maqsad — bolaning nutqiy "
        "faolligini oshirish va oddiy bog'lanishli nutq ko'nikmalarini "
        "shakllantirishdir. Asosiy vazifalar: barcha asosiy tovushlarni "
        "mustahkamlash, lug'atni 1500-2000 so'zgacha kengaytirish, "
        "3-4 so'zli gaplar tuzish, suhbatda qatnashish, qisqa "
        "she'rlarni yodlash va rasmlar bo'yicha oddiy hikoya tuzish.",

        "O'rta guruh (4-5 yosh) uchun maqsad — nutqni sifat jihatidan "
        "takomillashtirish va dialogik nutqni rivojlantirishdir. Asosiy "
        "vazifalar: murakkab tovushlarni (s, z, sh, ch, ts) "
        "mustahkamlash, lug'atni 2500-3000 so'zgacha oshirish, qo'shma "
        "gaplar tuzishga o'rgatish, savollarga batafsil javob berish, "
        "o'yinchoq yoki rasm bo'yicha hikoya tuzish va ertaklarni qayta "
        "hikoya qilish.",

        "Katta guruh (5-6 yosh) uchun maqsad — monologik nutqni "
        "shakllantirish va nutqning barcha tomonlarini takomillashtirish. "
        "Asosiy vazifalar: barcha nutq tovushlarini (r, l) to'liq "
        "egallash, lug'atni 3000-3500 so'zgacha boyitish, murakkab "
        "gaplar tuzish, mustaqil hikoya va ertak tuzish, she'rlarni "
        "ifodali aytish va nutq madaniyati qoidalariga rioya qilish.",

        "Maktabga tayyorlov guruhi (6-7 yosh) uchun maqsad — bolani "
        "maktab ta'limiga nutqiy jihatdan to'liq tayyorlash. Asosiy "
        "vazifalar: nutqning barcha tomonlarini mukammallashtirish, "
        "lug'atni 4000 va undan ortiq so'zgacha yetkazish, murakkab "
        "monologik nutq — hikoya, tasvirlash, mulohaza tuzish, tovush "
        "tahlili va sintezini o'rgatish, harflar bilan tanishtirish va "
        "dastlabki o'qish ko'nikmalarini shakllantirish. Shunday qilib, "
        "har bir yosh guruhida vazifalar oldingi bosqichni mustahkamlab, "
        "yangi ko'nikmalar qo'shib boradi.",
    ],

    # 8. Zamonaviy tendensiyalar
    [
        "Bolalarni nutqini o'stirish fanida zamonaviy davrda bir qator "
        "yangi tendensiyalar kuzatilmoqda. Birinchi tendensiya — "
        "kommunikativ yondashuv. Zamonaviy metodikada nutqni "
        "o'rgatishning asosiy maqsadi bolani grammatik jihatdan to'g'ri "
        "gapirishdan ko'ra, samarali muloqot qila olishga o'rgatish "
        "deb qaralmoqda. Ya'ni nutq faqat tizim sifatida emas, balki "
        "muloqot vositasi sifatida o'rgatiladi.",

        "Ikkinchi tendensiya — integrativ yondashuv. Nutqiy rivojlanish "
        "alohida fan sifatida emas, balki barcha ta'limiy sohalar bilan "
        "integratsiya qilingan holda amalga oshirilmoqda. Bola nutqini "
        "faqat maxsus mashg'ulotlarda emas, balki matematik, tabiat, "
        "musiqa, jismoniy tarbiya va boshqa faoliyatlarda ham "
        "rivojlantirish ko'zda tutiladi.",

        "Uchinchi tendensiya — texnologik vositalardan foydalanish. "
        "Zamonaviy maktabgacha ta'limda audio-kitoblar, interaktiv "
        "doska, planshet ilovalari va kompyuter o'yinlari nutqiy "
        "rivojlanish vositasi sifatida qo'llanilmoqda. Bu vositalar "
        "bolaning qiziqishini oshiradi va nutqiy mashqlarni qiziqarli "
        "tarzda taqdim etadi. Biroq ular an'anaviy usullarni "
        "to'ldiradi, almashtirmaydi.",

        "To'rtinchi tendensiya — individual yondashuv. Har bir bolaning "
        "nutqiy rivojlanishi individual tezlikda kechishi tan olinmoqda. "
        "Shu sababli zamonaviy metodikada individual ta'limiy "
        "marshrutlar tuzish, differentsial topshiriqlar berish va "
        "bolaning kuchli tomonlariga tayangan holda ishlash tavsiya "
        "etilmoqda. Bu yondashuv har bir bolaning o'z imkoniyatlarini "
        "to'liq namoyon qilishiga sharoit yaratadi.",

        "Beshinchi tendensiya — ota-onalar ishtirokining kuchayishi. "
        "Zamonaviy yondashuvda oilaviy nutqiy muhitning ahamiyati "
        "alohida ta'kidlanmoqda. Tarbiyachi ota-onalarga bolaning "
        "nutqini uyda rivojlantirish bo'yicha aniq tavsiyalar beradi, "
        "ular bilan seminarlar o'tkazadi va birgalikda ish olib boradi. "
        "Oltinchi tendensiya — inklyuziv ta'limda nutqiy yordam. "
        "Maxsus ehtiyojli bolalarning nutqiy rivojlanishiga ham "
        "alohida e'tibor qaratilmoqda va ular uchun moslashtirilgan "
        "dasturlar ishlab chiqilmoqda.",
    ],

    # 9. Xulosa
    [
        "Mazkur mustaqil ishda bolalarni nutqini o'stirish fanining "
        "maqsadi va vazifalari har tomonlama tahlil qilindi. Tahlillar "
        "shuni ko'rsatdiki, bu fanning bosh maqsadi — bolada ona tilida "
        "erkin, to'g'ri va ifodali muloqot qilish qobiliyatini "
        "shakllantirish hamda uni maktab ta'limiga nutqiy jihatdan "
        "tayyorlashdir.",

        "Fanning vazifalari tizimli bo'lib, ular tovush madaniyatini "
        "tarbiyalash, lug'at boyligini oshirish, grammatik to'g'ri "
        "nutqni shakllantirish, bog'lanishli nutqni rivojlantirish, "
        "badiiy adabiyot bilan tanishtirish va savodxonlikka "
        "tayyorlashdan iborat. Bu vazifalar bir-birini to'ldiradi va "
        "birgalikda bolaning nutqiy kamolotini ta'minlaydi.",

        "Har bir yosh guruhida fanning maqsad va vazifalari bolaning "
        "psixologik va fiziologik imkoniyatlariga mos tarzda "
        "aniqlashtirilib, murakkablik darajasi asta-sekin oshirib "
        "boriladi. Bu izchillik bolaning nutqiy rivojlanishining "
        "uzluksizligini ta'minlaydi va har bir bosqichda yangi "
        "cho'qqilarni zabt etishga imkon beradi.",

        "Xulosa qilib aytganda, bolalarni nutqini o'stirish fanining "
        "maqsad va vazifalari aniq belgilangan, ilmiy asoslangan va "
        "amaliy jihatdan sinab ko'rilgan. Bu maqsad va vazifalar "
        "tarbiyachining kundalik ishida yo'l-yo'riq bo'lib xizmat "
        "qiladi va bolalarning nutqiy kamolotini maqsadli ravishda "
        "ta'minlashga yordam beradi. Zamonaviy tendensiyalar esa bu "
        "sohani yanada rivojlantirib, boyitib bormoqda.",

        "Foydalanilgan adabiyotlar:",
        "1. O'zbekiston Respublikasining \"Maktabgacha ta'lim va tarbiya "
        "to'g'risida\"gi Qonuni. — Toshkent, 2020.",
        "2. Maktabgacha ta'limning Davlat o'quv dasturi \"Ilk qadam\". — "
        "Toshkent, 2018.",
        "3. Sodiqova M. \"Bolalar nutqini o'stirish metodikasi\". — "
        "Toshkent, 2019.",
        "4. Raximova D. \"Maktabgacha yoshdagi bolalar nutqini "
        "rivojlantirish\". — Toshkent, 2021.",
        "5. Ushakova O.S. \"Metodika razvitiya rechi detey doshkolnogo "
        "vozrasta\". — Moskva, 2017.",
        "6. Alekseyeva M.M. \"Metodika razvitiya rechi i obucheniya "
        "rodnomu yazyku doshkolnikov\". — Moskva, 2018.",
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
    "Bolalarni nutqini o'stirish fanining maqsadi va vazifasi",
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
<dc:title>Mustaqil ish - Bolalarni nutqini o stirish fanining maqsadi va vazifasi</dc:title>
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

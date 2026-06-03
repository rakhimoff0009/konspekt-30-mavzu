#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mustaqil ish: Bolalar jismoniy tarbiya fanining tadqiqot metodlari
Target: 9 A4 pages, Times New Roman 14pt, 1.5 line spacing
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT = "/projects/sandbox/konspekt-30-mavzu/Jismoniy_tarbiya_tadqiqot_metodlari.docx"

PLAN = [
    "Kirish",
    "Bolalar jismoniy tarbiya fanining predmeti va ilmiy asoslari",
    "Tadqiqot metodlari tushunchasi va tasnifi",
    "Pedagogik kuzatish metodi",
    "Pedagogik eksperiment metodi",
    "Jismoniy rivojlanishni o'lchash va testlash metodlari",
    "So'rovnoma va suhbat metodlari",
    "Matematik-statistik tahlil metodlari",
    "Xulosa va foydalanilgan adabiyotlar",
]

SECTIONS = [
    # 1. Kirish
    [
        "Jismoniy tarbiya — bu shaxsning har tomonlama rivojlanishining muhim "
        "tarkibiy qismi bo'lib, u bolaning sog'lig'ini mustahkamlash, harakat "
        "ko'nikmalarini shakllantirish va jismoniy sifatlarini rivojlantirishga "
        "yo'naltirilgan pedagogik jarayondir. Maktabgacha yoshdagi bolalarning "
        "jismoniy tarbiyasi ularning keyingi hayotidagi sog'lom turmush tarzi "
        "uchun mustahkam asos yaratadi. Shu sababli bu soha pedagogik fanning "
        "eng muhim yo'nalishlaridan biri sifatida e'tirof etiladi.",

        "Bolalar jismoniy tarbiya fani har qanday pedagogik fan singari o'z "
        "tadqiqot metodlariga ega. Tadqiqot metodlari fanning ilmiy asosini "
        "tashkil etadi va jismoniy tarbiya jarayonining samaradorligini "
        "aniqlash, yangi usullarni ishlab chiqish va mavjud amaliyotni "
        "takomillashtirish imkonini beradi. Ilmiy tadqiqotlarsiz fan "
        "rivojlana olmaydi va pedagogik amaliyot ilmiy asosga ega bo'lmaydi.",

        "Mazkur mustaqil ishda bolalar jismoniy tarbiya fanining tadqiqot "
        "metodlari batafsil yoritiladi. Ishda pedagogik kuzatish, eksperiment, "
        "o'lchash va testlash, so'rovnoma va suhbat, matematik-statistik "
        "tahlil metodlari tahlil qilinadi. Har bir metodning mohiyati, "
        "qo'llanilish tartibi, afzalliklari va kamchiliklari ko'rsatiladi. "
        "Mavzuning dolzarbligi shundaki, ilmiy tadqiqot metodlarini "
        "to'g'ri qo'llash jismoniy tarbiya sifatini oshirishning asosiy "
        "sharti hisoblanadi.",
    ],

    # 2. Fanning predmeti va ilmiy asoslari
    [
        "Bolalar jismoniy tarbiya fani — bu maktabgacha yoshdagi bolalarning "
        "jismoniy rivojlanishi va jismoniy tarbiyasi qonuniyatlarini, "
        "mazmunini, usullarini va tashkiliy shakllarini o'rganuvchi pedagogik "
        "fan sohasidir. Fanning predmeti — maktabgacha yoshdagi bolalarning "
        "jismoniy tarbiyasi pedagogik jarayoni bo'lib, u bolaning harakat "
        "ko'nikmalarini shakllantirish, jismoniy sifatlarini rivojlantirish "
        "va sog'lig'ini mustahkamlashni o'z ichiga oladi.",

        "Bolalar jismoniy tarbiya fanining ilmiy asoslarini bir nechta "
        "fundamental fan tashkil etadi. Birinchidan, anatomiya va fiziologiya "
        "fanlari bolaning tana tuzilishi, mushaklari, suyaklari, ichki "
        "a'zolari va nerv tizimining yosh xususiyatlarini tushuntirib beradi. "
        "Bu bilimlar jismoniy mashqlarni tanlash va dozalashda muhim ahamiyatga "
        "ega. Bolaning yosh xususiyatlarini hisobga olmasdan berilgan ortiqcha "
        "yuk sog'lig'iga zarar yetkazishi mumkin.",

        "Ikkinchidan, gigiyena va valeologiya fanlari bolaning sog'lig'ini "
        "saqlash va mustahkamlash qoidalarini, to'g'ri ovqatlanish, kun "
        "tartibiga rioya qilish, chiniqish va sog'lom turmush tarzi "
        "asoslarini belgilaydi. Uchinchidan, pedagogika fani ta'lim va "
        "tarbiyaning umumiy qonuniyatlari, tamoyillari va usullarini "
        "taqdim etadi. To'rtinchidan, psixologiya fani bolaning harakat "
        "faoliyatiga bo'lgan motivatsiyasini, qiziqishlarini va emotsional "
        "holatini tushunishga yordam beradi.",

        "Bolalar jismoniy tarbiya fani quyidagi asosiy masalalarni hal "
        "qiladi: bolaning yosh davrlariga mos jismoniy mashqlarni aniqlash; "
        "harakat ko'nikmalarini shakllantirish usullarini ishlab chiqish; "
        "jismoniy sifatlar — tezkorlik, kuch, chidamlilik, egiluvchanlik, "
        "chaqqonlikni rivojlantirish yo'llarini belgilash; jismoniy "
        "tarbiya mashg'ulotlarining tashkiliy shakllarini takomillashtirish; "
        "bolaning jismoniy rivojlanishini baholash mezonlarini ishlab chiqish.",
    ],

    # 3. Tadqiqot metodlari tushunchasi va tasnifi
    [
        "Tadqiqot metodi — bu ilmiy bilimga erishish, muammoni o'rganish "
        "va hal etish uchun qo'llaniladigan usullar va vositalar majmuidir. "
        "Bolalar jismoniy tarbiya fanida tadqiqot metodlari jismoniy "
        "tarbiya jarayonining turli tomonlarini o'rganish, yangi usullarning "
        "samaradorligini aniqlash va ilmiy asoslangan xulosalar chiqarish "
        "uchun qo'llaniladi.",

        "Bolalar jismoniy tarbiya fanida qo'llaniladigan tadqiqot metodlarini "
        "quyidagi guruhlarga bo'lish mumkin. Birinchi guruh — nazariy "
        "metodlar. Bularga ilmiy adabiyotlar tahlili, hujjatlar tahlili, "
        "taqqoslash va umumlashtirish kiradi. Bu metodlar mavzu bo'yicha "
        "ilmiy bilimlarni o'rganish, oldingi tadqiqotlar natijalarini "
        "tahlil qilish va nazariy xulosalar chiqarish uchun ishlatiladi.",

        "Ikkinchi guruh — empirik (amaliy) metodlar. Bularga pedagogik "
        "kuzatish, pedagogik eksperiment, testlash, o'lchash, so'rovnoma, "
        "suhbat va intervyu kiradi. Bu metodlar bevosita amaliyotda, "
        "jismoniy tarbiya mashg'ulotlari jarayonida ma'lumotlar to'plash "
        "uchun qo'llaniladi. Ular tadqiqotning asosiy qismini tashkil "
        "etadi va aniq raqamli natijalar beradi.",

        "Uchinchi guruh — matematik-statistik metodlar. Bularga o'rtacha "
        "arifmetik, standart og'ish, korrelyatsiya koeffitsiyenti, "
        "Student t-mezoni va boshqa statistik ko'rsatkichlarni hisoblash "
        "kiradi. Bu metodlar to'plangan ma'lumotlarni qayta ishlash, "
        "natijalarning ishonchliligini aniqlash va ilmiy xulosalarning "
        "asoslanganligini ta'minlash uchun ishlatiladi.",

        "Har bir tadqiqotda odatda bir nechta metodlar birgalikda "
        "qo'llaniladi. Bu kompleks yondashuv tadqiqot natijalarining "
        "haqqoniyligini va ishonchliligini oshiradi. Tadqiqotchi "
        "muammoning xarakteriga, tadqiqot maqsadiga va sharoitlarga "
        "qarab eng mos metodlar majmuasini tanlaydi. Metodlarning "
        "to'g'ri tanlanishi va qo'llanilishi tadqiqot sifatini belgilaydi.",
    ],

    # 4. Pedagogik kuzatish metodi
    [
        "Pedagogik kuzatish — bolalar jismoniy tarbiya fanida eng keng "
        "tarqalgan va asosiy tadqiqot metodlaridan biridir. Bu metod "
        "tadqiqotchining bolalar jismoniy tarbiyasi jarayonini bevosita "
        "kuzatishi va muhim ma'lumotlarni qayd etishiga asoslanadi. "
        "Pedagogik kuzatish tabiiy sharoitda, ya'ni mashg'ulot, o'yin "
        "yoki erkin faoliyat jarayonida amalga oshiriladi.",

        "Pedagogik kuzatishning bir nechta turi mavjud. Ochiq kuzatishda "
        "tadqiqotchi ochiqdan-ochiq mashg'ulotda ishtirok etadi va "
        "bolalar uning kuzatayotganini biladi. Yashirin kuzatishda esa "
        "tadqiqotchi bolalarga ko'rinmasdan yoki aralashmasdan kuzatadi. "
        "Ishtirokchi kuzatishda tadqiqotchi mashg'ulotda bevosita "
        "qatnashib, bir vaqtda ma'lumot to'playdi. Har bir turning "
        "o'ziga xos afzalliklari va kamchiliklari bor.",

        "Kuzatish metodi qo'llanilganda tadqiqotchi oldindan kuzatish "
        "rejasini tuzadi. Rejada quyidagilar belgilanadi: nimani kuzatish "
        "kerak (kuzatish ob'ekti), qancha vaqt kuzatiladi, qanday "
        "ko'rsatkichlar qayd etiladi va natijalar qanday shaklda "
        "yozib boriladi. Masalan, bolalarning jismoniy mashq bajarish "
        "texnikasini kuzatishda har bir harakat elementi alohida baholanadi.",

        "Kuzatish natijalarini qayd etish uchun maxsus protokollar, "
        "kuzatuv kartochkalari va texnik vositalar (videokamera, "
        "diktofon) ishlatiladi. Zamonaviy tadqiqotlarda video yozuv "
        "keng qo'llaniladi, chunki u harakatni qayta-qayta ko'rib "
        "tahlil qilish imkonini beradi. Kuzatish metodi bolalarning "
        "harakat ko'nikmalarini, jismoniy sifatlarini va mashg'ulotga "
        "bo'lgan munosabatini o'rganishda samarali hisoblanadi.",

        "Pedagogik kuzatishning afzalliklari: tabiiy sharoitda amalga "
        "oshiriladi, bolaga qo'shimcha yuk tushmaydi, uzluksiz ma'lumot "
        "to'plash imkonini beradi. Kamchiliklari: tadqiqotchining "
        "sub'ektiv bahosi ta'sir qilishi mumkin, bir vaqtda ko'p "
        "bolani kuzatish qiyin, barcha muhim jihatlarni sezib olish "
        "har doim ham imkon bo'lmaydi. Shu sababli kuzatish boshqa "
        "metodlar bilan birgalikda qo'llanilishi tavsiya etiladi.",
    ],

    # 5. Pedagogik eksperiment metodi
    [
        "Pedagogik eksperiment — bolalar jismoniy tarbiya fanida eng "
        "muhim va ishonchli tadqiqot metodlaridan biri hisoblanadi. "
        "Eksperiment — bu maxsus yaratilgan sharoitlarda muayyan "
        "omilning ta'sirini sinab ko'rish va natijalarni aniqlash "
        "usulidir. Jismoniy tarbiyada eksperiment yangi mashq usullari, "
        "mashg'ulot shakllari yoki tarbiya vositalarining "
        "samaradorligini tekshirish uchun o'tkaziladi.",

        "Pedagogik eksperimentning bir nechta turi ajratiladi. Tabiiy "
        "eksperiment odatiy mashg'ulotlar sharoitida o'tkaziladi — "
        "bolalar o'zlarining sinab ko'rilayotganini bilmaydi. "
        "Laboratorion eksperiment maxsus yaratilgan sharoitlarda "
        "aniq o'lchov asboblari yordamida o'tkaziladi. Aniqlash "
        "eksperimenti mavjud holatni o'rganish uchun, shakllantirish "
        "eksperimenti esa yangi usulni joriy etib, uning natijasini "
        "kuzatish uchun ishlatiladi.",

        "Eksperiment o'tkazishning asosiy bosqichlari quyidagilardan "
        "iborat. Birinchi bosqich — eksperiment rejasini tuzish: "
        "maqsadni aniqlash, gipotezani shakllantirish, ishtirokchilarni "
        "tanlash. Ikkinchi bosqich — dastlabki o'lchov (pretest): "
        "eksperiment boshlanishidan oldin bolalarning jismoniy "
        "tayyorgarlik darajasini aniqlash. Uchinchi bosqich — "
        "eksperimental ta'sir: yangi usulni qo'llash.",

        "To'rtinchi bosqich — yakuniy o'lchov (posttest): eksperiment "
        "oxirida bolalarning natijalarini qayta o'lchash. Beshinchi "
        "bosqich — natijalarni taqqoslash va tahlil qilish: dastlabki "
        "va yakuniy ko'rsatkichlar orasidagi farqni aniqlash. Odatda "
        "eksperimentda ikki guruh ishtirok etadi — eksperimental guruh "
        "(yangi usul qo'llaniladi) va nazorat guruhi (an'anaviy usul "
        "davom ettiriladi).",

        "Jismoniy tarbiyada eksperiment misollar: yangi harakatli "
        "o'yinlar dasturining bolalar chaqqonligiga ta'sirini aniqlash; "
        "musiqa jo'natmasida gimnastika mashqlarining samaradorligini "
        "tekshirish; suv bilan chiniqish usulining bolalar "
        "kasallanish darajasiga ta'sirini o'rganish. Eksperiment "
        "metodi eng ishonchli natijalar beradi, lekin uni o'tkazish "
        "ko'p vaqt va imkoniyat talab qiladi. Shuningdek, bolalar "
        "bilan ishlashda axloqiy qoidalarga rioya qilish shart.",
    ],

    # 6. O'lchash va testlash metodlari
    [
        "Jismoniy rivojlanishni o'lchash va testlash metodlari bolalar "
        "jismoniy tarbiya fanida aniq miqdoriy ma'lumotlar olish uchun "
        "qo'llaniladi. Bu metodlar bolaning jismoniy rivojlanish "
        "darajasini, jismoniy tayyorgarligini va jismoniy sifatlarini "
        "raqamli ko'rsatkichlar orqali baholash imkonini beradi. "
        "O'lchash va testlash natijalari ob'ektiv bo'lib, ularni "
        "taqqoslash va statistik tahlil qilish mumkin.",

        "Jismoniy rivojlanishni o'lchash uchun antropometrik "
        "ko'rsatkichlardan foydalaniladi. Bularga bo'y uzunligi, tana "
        "og'irligi, ko'krak qafasining aylanasi, bosh aylanasi, "
        "qo'l va oyoq uzunligi kiradi. Bu ko'rsatkichlar maxsus "
        "asboblar — rost-o'mer (stadiometr), tarozi, santimetrli "
        "lenta yordamida o'lchanadi. Natijalar yosh-jins "
        "standartlari bilan taqqoslanib, bolaning jismoniy rivojlanish "
        "darajasi aniqlanadi.",

        "Jismoniy tayyorgarlikni testlash uchun maxsus motor testlar "
        "qo'llaniladi. Maktabgacha yoshdagi bolalar uchun quyidagi "
        "testlar keng tarqalgan: 30 metr masofaga yugurish (tezkorlik), "
        "uzunlikka sakrash (oyoq kuchi), to'pni uzoqqa otish (qo'l "
        "kuchi), 10 metr x 3 marta chopish (chaqqonlik), turib oldinga "
        "egilish (egiluvchanlik), bir oyoqda turish (muvozanat). "
        "Har bir test aniq qoidalar bo'yicha o'tkaziladi.",

        "Testlash jarayonida quyidagi talablarga rioya qilinadi: "
        "barcha bolalar bir xil sharoitda test topshiradi; testdan "
        "oldin yetarli isitish mashqlari o'tkaziladi; har bir bola "
        "ikki yoki uch marta urinadi va eng yaxshi natija hisobga "
        "olinadi; natijalar darhol qayd etiladi; bolaning sog'lig'i "
        "va kayfiyati hisobga olinadi. Test natijalari yosh-jins "
        "me'yoriy jadvallari asosida baholanadi.",

        "Funktsional holatni baholash testlari ham muhim ahamiyatga "
        "ega. Bularga yurak urish tezligini o'lchash, Shtige testi "
        "(bolaning jismoniy yuklamadan keyingi tiklanish tezligini "
        "aniqlash), nafas tutib turish testi va boshqalar kiradi. "
        "Bu testlar bolaning yurak-qon tomir va nafas olish "
        "tizimlarining jismoniy yuklamaga moslashish darajasini "
        "ko'rsatadi. O'lchash va testlash metodlarining asosiy "
        "afzalligi — ular ob'ektiv va raqamli natijalar berib, "
        "taqqoslash va dinamikani kuzatish imkonini yaratadi.",
    ],

    # 7. So'rovnoma va suhbat metodlari
    [
        "So'rovnoma va suhbat metodlari bolalar jismoniy tarbiya "
        "fanida qo'shimcha muhim ma'lumotlar olish uchun qo'llaniladi. "
        "Bu metodlar bevosita bolalardan, ularning ota-onalaridan va "
        "tarbiyachilardan axborot to'plash imkonini beradi. Ular "
        "jismoniy tarbiya jarayonining sub'ektiv tomonlarini — "
        "munosabatlarni, qiziqishlarni, fikrlarni va tajribalarni "
        "o'rganishda ayniqsa samarali hisoblanadi.",

        "So'rovnoma metodi — bu maxsus ishlab chiqilgan savollar "
        "ro'yxati (anketa) yordamida ma'lumot to'plash usulidir. "
        "Bolalar jismoniy tarbiya sohasida so'rovnomalar asosan "
        "ota-onalar va tarbiyachilar uchun tuziladi. Ota-onalarga "
        "berilgan so'rovnomalar bolaning uyda jismoniy faolligi, "
        "sport bilan shug'ullanishi, ovqatlanish tartibi va sog'lig'i "
        "haqida ma'lumot to'plashga qaratilgan.",

        "Tarbiyachilar uchun so'rovnomalar jismoniy tarbiya "
        "mashg'ulotlarining tashkil etilishi, qo'llaniladigan usullar, "
        "duch kelingan qiyinchiliklar va kasbiy ehtiyojlar haqida "
        "ma'lumot beradi. So'rovnomalar ochiq savollar (erkin javob), "
        "yopiq savollar (tayyor variantlardan tanlash) va aralash "
        "shaklda tuzilishi mumkin. Savollar aniq, tushunarli va "
        "ob'ektiv javob olishga yo'naltirilgan bo'lishi kerak.",

        "Suhbat (intervyu) metodi — bu tadqiqotchining respondent "
        "bilan yuzma-yuz muloqot qilishi orqali ma'lumot olish "
        "usulidir. Suhbat individuallashtirilgan bo'lib, tadqiqotchi "
        "javoblarga qarab qo'shimcha savollar berishi, mavzuni "
        "chuqurroq o'rganishi mumkin. Bolalar bilan suhbat "
        "o'tkazishda maxsus yondashuv talab etiladi — savollar "
        "sodda, qisqa va bolaning yoshiga mos bo'lishi kerak.",

        "Katta yoshdagi maktabgacha bolalar bilan suhbatda ularning "
        "jismoniy mashg'ulotlarga munosabati, sevimli o'yinlari, "
        "sport turlari haqidagi bilimlarini aniqlash mumkin. "
        "So'rovnoma va suhbat metodlarining afzalliklari: ko'p "
        "sonli respondentlardan bir vaqtda ma'lumot olish mumkin "
        "(so'rovnoma); chuqur va batafsil ma'lumot olish mumkin "
        "(suhbat). Kamchiliklari: sub'ektiv javoblar bo'lishi, "
        "respondentning halol javob bermasligi, maktabgacha yoshdagi "
        "bolalar bilan ishlashda qiyinchiliklar yuzaga kelishi mumkin.",
    ],

    # 8. Matematik-statistik tahlil
    [
        "Matematik-statistik tahlil metodlari bolalar jismoniy tarbiya "
        "fanida to'plangan ma'lumotlarni qayta ishlash, natijalarning "
        "ishonchliligini aniqlash va ilmiy xulosalar chiqarish uchun "
        "qo'llaniladi. Bu metodlar tadqiqot natijalarining ob'ektiv "
        "va ilmiy asoslangan bo'lishini ta'minlaydi. Statistik tahlilsiz "
        "birorta ham jiddiy ilmiy tadqiqot to'liq hisoblanmaydi.",

        "Eng ko'p qo'llaniladigan statistik ko'rsatkichlar quyidagilar. "
        "O'rtacha arifmetik (X) — guruhdagi barcha natijalarning "
        "o'rtacha qiymatini ko'rsatadi. Masalan, guruhdagi bolalarning "
        "o'rtacha bo'yi yoki o'rtacha yugurish natijasi. Standart "
        "og'ish (S) — natijalarning o'rtacha qiymatdan qanchalik "
        "tarqalganligini ko'rsatadi. Kichik standart og'ish "
        "natijalarning bir xilligini, katta standart og'ish esa "
        "katta farqlar mavjudligini bildiradi.",

        "Farqlarning ishonchliligini aniqlash uchun Student t-mezoni "
        "keng qo'llaniladi. Bu mezon ikki guruh o'rtasidagi yoki "
        "bitta guruhning eksperiment oldi va keyingi natijalari "
        "o'rtasidagi farqning statistik jihatdan ahamiyatli ekanligini "
        "aniqlaydi. Agar t-qiymat jadval qiymatidan katta bo'lsa "
        "(p<0.05), farq ishonchli deb hisoblanadi va u tasodifiy "
        "emas, balki eksperimental ta'sir natijasi deb qabul qilinadi.",

        "Korrelyatsiya tahlili ikki ko'rsatkich orasidagi bog'liqlikni "
        "aniqlash uchun ishlatiladi. Masalan, bolaning bo'yi va "
        "sakrash uzoqligi orasidagi bog'liqlik, yoki mashg'ulotlar "
        "soni va jismoniy tayyorgarlik darajasi orasidagi aloqa "
        "korrelyatsiya koeffitsiyenti orqali aniqlanadi. Koeffitsiyent "
        "+1 ga yaqin bo'lsa kuchli ijobiy, -1 ga yaqin bo'lsa "
        "kuchli salbiy, 0 ga yaqin bo'lsa aloqa yo'q deb baholanadi.",

        "Zamonaviy tadqiqotlarda matematik-statistik tahlil uchun "
        "kompyuter dasturlaridan — SPSS, Excel, Statistica — "
        "foydalaniladi. Bu dasturlar murakkab hisob-kitoblarni tez "
        "va aniq bajarish imkonini beradi. Statistik tahlil "
        "natijalari jadvallar va grafiklar shaklida taqdim etiladi. "
        "Matematik-statistik metodlarning ahamiyati shundaki, ular "
        "tadqiqot xulosalarining ilmiy ishonchliligini ta'minlaydi "
        "va sub'ektiv fikrlardan farqli ravishda, raqamlar asosida "
        "ob'ektiv xulosalar chiqarish imkonini beradi.",
    ],

    # 9. Xulosa
    [
        "Mazkur mustaqil ishda bolalar jismoniy tarbiya fanining "
        "tadqiqot metodlari har tomonlama tahlil qilindi. Tahlillar "
        "shuni ko'rsatdiki, bu fanda qo'llaniladigan tadqiqot metodlari "
        "xilma-xil bo'lib, ular nazariy, empirik va matematik-statistik "
        "guruhlarga bo'linadi. Har bir metod o'ziga xos vazifalarga "
        "xizmat qiladi va birgalikda tadqiqotning to'liq manzarasini "
        "beradi.",

        "Pedagogik kuzatish metodi jismoniy tarbiya jarayonini tabiiy "
        "sharoitda o'rganish imkonini beradi. Pedagogik eksperiment esa "
        "yangi usullarning samaradorligini ishonchli tarzda aniqlaydi. "
        "O'lchash va testlash metodlari ob'ektiv raqamli "
        "ko'rsatkichlar beradi. So'rovnoma va suhbat sub'ektiv "
        "ma'lumotlarni to'plashga xizmat qiladi. Matematik-statistik "
        "metodlar esa barcha natijalarning ishonchliligini ta'minlaydi.",

        "Tadqiqot metodlarining to'g'ri tanlanishi va qo'llanilishi "
        "ilmiy ishning sifatini belgilaydi. Har bir tadqiqotchi o'z "
        "muammosiga mos metodlar majmuasini tanlashi, ularni to'g'ri "
        "qo'llashi va natijalarni ilmiy asosda tahlil qilishi kerak. "
        "Buning uchun tadqiqot metodologiyasini chuqur bilish va "
        "amaliy tajribaga ega bo'lish zarur.",

        "Xulosa qilib aytganda, bolalar jismoniy tarbiya fanining "
        "tadqiqot metodlari — bu fanning ilmiy asosini tashkil etuvchi "
        "muhim vositalardir. Ular jismoniy tarbiya jarayonini ilmiy "
        "jihatdan o'rganish, samaradorligini baholash va "
        "takomillashtirish imkonini beradi. Bu metodlarni egallash har "
        "bir tarbiyachi va tadqiqotchi uchun kasbiy zaruratdir.",

        "Foydalanilgan adabiyotlar:",
        "1. O'zbekiston Respublikasining \"Jismoniy tarbiya va sport "
        "to'g'risida\"gi Qonuni. — Toshkent, 2015.",
        "2. Maktabgacha ta'limning Davlat o'quv dasturi \"Ilk qadam\". — "
        "Toshkent, 2018.",
        "3. Xolmatov A. \"Bolalar jismoniy tarbiyasi nazariyasi va "
        "metodikasi\". — Toshkent, 2020.",
        "4. Abdullayev A. \"Jismoniy tarbiya va sportda ilmiy tadqiqot "
        "metodlari\". — Toshkent, 2019.",
        "5. Stepanenkova E.Ya. \"Teoriya i metodika fizicheskogo "
        "vospitaniya i razvitiya rebyonka\". — Moskva, 2018.",
        "6. Xolodov J.K. \"Teoriya i metodika fizicheskoy kultury i "
        "sporta\". — Moskva, 2017.",
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
    "Bolalar jismoniy tarbiya fanining tadqiqot metodlari",
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
<dc:title>Mustaqil ish - Bolalar jismoniy tarbiya tadqiqot metodlari</dc:title>
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

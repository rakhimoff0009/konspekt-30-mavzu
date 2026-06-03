#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mustaqil ish: Bolalar jismoniy tarbiya fanining rivojlanish tizimi
Target: 9 A4 pages, Times New Roman 14pt, 1.5 line spacing
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT = "/projects/sandbox/konspekt-30-mavzu/Jismoniy_tarbiya_rivojlanish_tizimi.docx"

PLAN = [
    "Kirish",
    "Bolalar jismoniy tarbiya fanining shakllanish tarixi",
    "Jismoniy tarbiya fanining tizimi va tarkibiy qismlari",
    "Jismoniy tarbiya tizimining maqsad va vazifalari",
    "Jismoniy tarbiya tizimining tamoyillari",
    "Maktabgacha ta'limda jismoniy tarbiya tizimining tashkiliy shakllari",
    "Jismoniy tarbiya tizimida qo'llaniladigan vositalar",
    "Jismoniy tarbiya fanining zamonaviy rivojlanish yo'nalishlari",
    "Xulosa va foydalanilgan adabiyotlar",
]

SECTIONS = [
    # 1. Kirish
    [
        "Jismoniy tarbiya — shaxsning har tomonlama kamol topishida muhim "
        "o'rin tutadigan pedagogik jarayon bo'lib, u bolaning sog'lig'ini "
        "mustahkamlash, harakat ko'nikmalarini shakllantirish, jismoniy "
        "sifatlarini rivojlantirish va sog'lom turmush tarziga munosabatini "
        "tarbiyalashga yo'naltirilgandir. Maktabgacha yoshdagi bolalarning "
        "jismoniy tarbiyasi ularning butun hayoti davomidagi sog'lig'i va "
        "farovonligi uchun mustahkam poydevor yaratadi.",

        "Bolalar jismoniy tarbiya fani mustaqil pedagogik fan sifatida "
        "o'zining rivojlanish tizimiga ega. Bu tizim tarixiy shakllanish "
        "jarayoni, ilmiy-nazariy asoslar, tashkiliy shakllar, vositalar "
        "va usullarning yaxlit majmuini o'z ichiga oladi. Fanning "
        "rivojlanish tizimini o'rganish bugungi pedagogik amaliyotni "
        "to'g'ri tushunish va kelajakda yanada takomillashtirish uchun "
        "zarur bo'lgan bilimlarni beradi.",

        "Mazkur mustaqil ishda bolalar jismoniy tarbiya fanining rivojlanish "
        "tizimi batafsil yoritiladi. Ishda fanning shakllanish tarixi, "
        "tizimning tarkibiy qismlari, maqsad va vazifalari, tamoyillari, "
        "tashkiliy shakllari, vositalari va zamonaviy rivojlanish "
        "yo'nalishlari tahlil qilinadi. Mavzuning dolzarbligi shundaki, "
        "jismoniy tarbiya tizimini mukammal bilish tarbiyachining kasbiy "
        "mahoratini oshirish va bolalar sog'lig'ini samarali ta'minlash "
        "uchun zarurdir.",
    ],

    # 2. Shakllanish tarixi
    [
        "Bolalar jismoniy tarbiya fanining shakllanishi uzoq tarixiy "
        "jarayonning natijasi hisoblanadi. Qadimgi davrlardayoq inson "
        "bolalarini harakatli o'yinlar, yugurish, sakrash va suzish orqali "
        "jismoniy jihatdan tarbiyalagan. Biroq bu jarayon ilmiy asosga ega "
        "bo'lmagan va faqat hayotiy tajribaga tayanilgan. O'rta asrlarda "
        "Sharq mutafakkirlari — Abu Ali ibn Sino, Abu Rayhon Beruniy va "
        "boshqalar bolalar jismoniy tarbiyasining ahamiyatini ta'kidlab, "
        "ilmiy fikrlar bildirganlar.",

        "Abu Ali ibn Sino o'zining \"Tib qonunlari\" asarida bolalarning "
        "yoshiga qarab jismoniy mashqlar tizimini tavsiya etgan. U "
        "bolaning yoshi, jismi va sog'lig'iga qarab turli xil harakat "
        "mashqlarini tanlash zarurligini ko'rsatgan. Bu g'oyalar zamonaviy "
        "jismoniy tarbiya fanining dastlabki ilmiy asoslari sifatida "
        "baholanadi. Shuningdek, xalq pedagogikasida harakatli o'yinlar "
        "orqali bolalarni tarbiyalash an'anasi asrlar davomida saqlanib "
        "kelgan.",

        "Yevropa pedagogikasida Jan Amos Komenskiy, Jon Lokk, Jan-Jak "
        "Russo kabi mutafakkirlar bolalarning jismoniy tarbiyasiga "
        "alohida e'tibor qaratganlar. Komenskiy ta'limning tabiiylik "
        "tamoyilini ilgari surib, bolaning harakat ehtiyojini qondirish "
        "zarurligini ta'kidlagan. P.F. Lesgaft esa ilmiy jismoniy tarbiya "
        "tizimining asoschisi bo'lib, u bolalar jismoniy tarbiyasining "
        "pedagogik asoslarini ishlab chiqqan.",

        "XX asrda bolalar jismoniy tarbiya fani alohida mustaqil fan "
        "sifatida shakllandi. A.V. Keneman, D.V. Xuxlayeva, T.I. Osokina "
        "va boshqa olimlar maktabgacha yoshdagi bolalarning jismoniy "
        "tarbiyasi nazariyasi va metodikasini ishlab chiqdilar. O'zbekistonda "
        "ham A. Xolmatov, B. Abdullayev va boshqa olimlar milliy sharoitga "
        "mos jismoniy tarbiya tizimini yaratish ustida ishladilar. Bugungi "
        "kunda fan yangi ilmiy yutuqlar asosida doimiy takomillashib "
        "bormoqda.",
    ],

    # 3. Tizimi va tarkibiy qismlari
    [
        "Bolalar jismoniy tarbiya fanining tizimi — bu fanning barcha "
        "tarkibiy qismlari, ularning o'zaro bog'liqligi va birgalikdagi "
        "faoliyatini aks ettiruvchi yaxlit tuzilmadir. Bu tizim bir "
        "nechta muhim komponentlardan iborat bo'lib, ularning har biri "
        "o'ziga xos vazifani bajaradi va boshqa komponentlar bilan "
        "uzviy bog'langan.",

        "Tizimning birinchi komponenti — g'oyaviy-nazariy asos. Bu "
        "komponent fanning falsafiy, pedagogik, psixologik va biologik "
        "asoslarini o'z ichiga oladi. U jismoniy tarbiyaning maqsadi, "
        "mohiyati va qonuniyatlari haqidagi ilmiy bilimlar tizimini "
        "tashkil etadi. G'oyaviy asos — bolaning har tomonlama "
        "rivojlanishida jismoniy tarbiyaning ajralmas o'rni haqidagi "
        "asosiy tamoyildir.",

        "Ikkinchi komponent — dasturiy-me'yoriy asos. Bu komponent "
        "davlat ta'lim standartlari, o'quv dasturlari, me'yoriy "
        "hujjatlar va jismoniy tayyorgarlik normativlarini o'z ichiga "
        "oladi. O'zbekistonda \"Ilk qadam\" davlat o'quv dasturi har "
        "bir yosh guruhi uchun jismoniy rivojlanish ko'rsatkichlari "
        "va harakat ko'nikmalarining kutilgan natijalarini belgilab "
        "bergan.",

        "Uchinchi komponent — tashkiliy asos. Bu komponent jismoniy "
        "tarbiya mashg'ulotlarining shakllari — ertalabki gimnastika, "
        "jismoniy tarbiya mashg'uloti, harakatli o'yinlar, sport "
        "bayramlari, sayrlarda jismoniy mashqlar va mustaqil harakat "
        "faoliyati kabi tashkiliy shakllarni belgilaydi.",

        "To'rtinchi komponent — kadrlar tayyorlash tizimi. Bu "
        "komponent jismoniy tarbiya sohasida malakali tarbiyachilar, "
        "sport instruktorlari va metodistlarni tayyorlash, ularning "
        "malakasini oshirish va kasbiy rivojlanishini ta'minlashni o'z "
        "ichiga oladi. Beshinchi komponent — moddiy-texnik ta'minot. "
        "Bu sport zallari, sport maydonchalar, jihozlar, inventar va "
        "o'quv-uslubiy adabiyotlar bilan ta'minlanganlikni bildiradi. "
        "Barcha komponentlar birgalikda yaxlit tizimni tashkil etadi.",
    ],

    # 4. Maqsad va vazifalari
    [
        "Bolalar jismoniy tarbiya tizimining bosh maqsadi — maktabgacha "
        "yoshdagi bolalarning sog'lig'ini mustahkamlash, jismoniy "
        "rivojlanishini ta'minlash, hayot uchun zarur harakat "
        "ko'nikmalarini shakllantirish va sog'lom turmush tarzi "
        "asoslarini tarbiyalashdir. Bu maqsad bolaning butun maktabgacha "
        "davridagi jismoniy tarbiyasining umumiy yo'nalishini belgilaydi "
        "va konkret vazifalar orqali amalga oshiriladi.",

        "Jismoniy tarbiya tizimining vazifalari uch guruhga bo'linadi. "
        "Birinchi guruh — sog'lomlashtirish vazifalari. Bularga bolaning "
        "sog'lig'ini saqlash va mustahkamlash, organizmni chiniqtirish, "
        "to'g'ri gavda holatini shakllantirish, organlar va tizimlar "
        "funksiyalarini yaxshilash, kasalliklarning oldini olish kiradi. "
        "Bu vazifalar bolaning biologik rivojlanishini ta'minlaydi.",

        "Ikkinchi guruh — ta'limiy vazifalar. Bularga asosiy harakat "
        "turlarini o'rgatish (yurish, yugurish, sakrash, otish, "
        "irmalab chiqish, suzish), sport mashqlarining asosiy "
        "elementlarini egallash, harakat ko'nikmalarini "
        "takomillashtirish, jismoniy mashqlar texnikasini o'zlashtirish "
        "kiradi. Bola hayotda zarur bo'lgan barcha asosiy harakatlarni "
        "to'g'ri va ishonchli bajara olishi kerak.",

        "Uchinchi guruh — tarbiyaviy vazifalar. Bularga bolada jismoniy "
        "mashqlarga qiziqish va ehtiyoj tuyg'usini shakllantirish, "
        "irodaviy sifatlarni — jasorat, qat'iyat, chidamlilik, sabr-toqat — "
        "tarbiyalash, jamoaviylik, o'rtoqlik, intizomlilik kabi axloqiy "
        "sifatlarni shakllantirish kiradi. Jismoniy tarbiya jarayonida "
        "bola nafaqat jismonan, balki ruhan ham kamol topadi.",

        "Shuningdek, jismoniy tarbiya tizimining muhim vazifasi — bolaning "
        "jismoniy sifatlarini rivojlantirishdir. Asosiy jismoniy sifatlar "
        "quyidagilar: tezkorlik — harakatlarni tez bajarish qobiliyati; "
        "kuch — mushaklarning zo'riqish qobiliyati; chidamlilik — uzoq "
        "muddat harakat qilish qobiliyati; egiluvchanlik — keng amplituda "
        "bilan harakat qilish; chaqqonlik — vaziyatga qarab harakatni "
        "o'zgartirish qobiliyati. Har bir sifat maxsus mashqlar orqali "
        "rivojlantiriladi.",
    ],

    # 5. Tamoyillari
    [
        "Bolalar jismoniy tarbiya tizimi muayyan tamoyillarga asoslanadi. "
        "Bu tamoyillar jismoniy tarbiya jarayonining ilmiy asosini tashkil "
        "etadi va amaliy faoliyatda yo'l-yo'riq vazifasini bajaradi. "
        "Birinchi tamoyil — har tomonlama rivojlantirish tamoyili. Bu "
        "tamoyilga ko'ra, jismoniy tarbiya faqat tanani emas, balki "
        "shaxsning barcha tomonlarini — aqliy, axloqiy, estetik va "
        "mehnat tarbiyasini ham qamrab olishi kerak.",

        "Ikkinchi tamoyil — sog'lig'ni mustahkamlash bilan bog'liqlik "
        "tamoyili. Jismoniy tarbiya jarayonidagi barcha mashqlar, o'yinlar "
        "va faoliyatlar bolaning sog'lig'ini mustahkamlashga, organizmini "
        "chiniqtirishga va kasalliklarning oldini olishga xizmat qilishi "
        "kerak. Ortiqcha yuk yoki noto'g'ri tanlangan mashqlar bolaning "
        "sog'lig'iga zarar yetkazishi mumkin.",

        "Uchinchi tamoyil — yoshga muvofiqlik tamoyili. Bu tamoyilga "
        "ko'ra, jismoniy mashqlarning mazmuni, hajmi va intensivligi "
        "bolaning yosh xususiyatlariga, tayanch-harakat apparatining "
        "rivojlanish darajasiga va funksional imkoniyatlariga mos "
        "bo'lishi shart. Har bir yosh davrida bolaning organizmiga "
        "tushadigan yuk darajasi qat'iy me'yorlangan.",

        "To'rtinchi tamoyil — izchillik va tizimlilik tamoyili. "
        "Jismoniy mashqlar oddiydan murakkabga, yengildan og'irga, "
        "ma'lumdan noma'lumga tamoyili asosida izchil o'rgatiladi. "
        "Har bir yangi mashq avvalgilarni mustahkamlash asosida "
        "o'rgatiladi. Mashg'ulotlar muntazam va tizimli olib boriladi.",

        "Beshinchi tamoyil — ongli va faollik tamoyili. Bola jismoniy "
        "mashqlarni bajarishda nima uchun bu harakatni qilayotganini "
        "tushunishi va faol ishtirok etishi kerak. Majburlash emas, "
        "qiziqtirish orqali bola harakatga undashni taqozo etadi. "
        "Oltinchi tamoyil — ko'rgazmalilik tamoyili. Maktabgacha yoshdagi "
        "bolalarga mashqlar namuna ko'rsatish, rasm, video va jonli "
        "namoyish orqali o'rgatiladi. Bu tamoyillar hammasi birgalikda "
        "samarali jismoniy tarbiya tizimini tashkil etadi.",
    ],

    # 6. Tashkiliy shakllar
    [
        "Maktabgacha ta'lim muassasalarida jismoniy tarbiya tizimi "
        "turli tashkiliy shakllar orqali amalga oshiriladi. Birinchi "
        "va eng asosiy shakl — jismoniy tarbiya mashg'uloti. Bu "
        "tarbiyachi yoki sport instruktori tomonidan haftada 2-3 marta "
        "o'tkaziladigan maxsus mashg'ulotdir. Mashg'ulot kirish, asosiy "
        "va yakuniy qismlardan iborat bo'lib, davomiyligi yosh guruhiga "
        "qarab 15-35 daqiqani tashkil etadi.",

        "Jismoniy tarbiya mashg'ulotining kirish qismida bola organizmini "
        "kelgusi yuklamaga tayyorlash uchun isitish mashqlari — yurish, "
        "engil yugurish, umumrivojlantiruvchi mashqlar bajariladi. Asosiy "
        "qismda yangi harakat ko'nikmalari o'rgatiladi, jismoniy sifatlar "
        "rivojlantiriladi va harakatli o'yinlar o'tkaziladi. Yakuniy "
        "qismda bolaning organizmini tinch holatga keltirish uchun "
        "sekin yurish, nafas mashqlari va tinchlanish mashqlari bajariladi.",

        "Ikkinchi muhim shakl — ertalabki gimnastika. Bu har kuni "
        "ertalab 5-10 daqiqa davom etadigan mashqlar majmui bo'lib, "
        "bolaning organizmini uyg'otish, qon aylanishini yaxshilash "
        "va kun davomida faollikni ta'minlash uchun o'tkaziladi. "
        "Gimnastika musiqiy jo'natma asosida, ochiq havoda yoki sport "
        "zalida o'tkaziladi.",

        "Uchinchi shakl — jismoniy tarbiya daqiqalari (fizminutkalar). "
        "Bu mashg'ulotlar orasida 2-3 daqiqali qisqa harakat "
        "pauzalari bo'lib, bolaning charchashini bartaraf etadi va "
        "e'tiborini yangilaydi. To'rtinchi shakl — harakatli o'yinlar. "
        "Bu bolalarning eng sevimli jismoniy faoliyat turi bo'lib, "
        "ularda yugurish, sakrash, otish, irmalab chiqish kabi harakat "
        "ko'nikmalari qo'llaniladi.",

        "Beshinchi shakl — sport bayramlari va o'yin-kulgilar. Bu "
        "bayramlar oyda yoki faslda bir marta o'tkazilib, bolalarning "
        "erishgan natijalarini namoyish etish, sport quvonchini his "
        "qilish va jamoaviy ruhni mustahkamlashga xizmat qiladi. "
        "Oltinchi shakl — sayr vaqtidagi jismoniy mashqlar. Ochiq "
        "havoda turli xil harakatli o'yinlar, estafetalar, tabiat "
        "bilan bog'liq jismoniy mashqlar o'tkaziladi. Yettinchi "
        "shakl — mustaqil harakat faoliyati. Bolalar o'z xohishlari "
        "bilan sport burchagidagi jihozlardan foydalanib, erkin "
        "harakatlanadi.",
    ],

    # 7. Vositalar
    [
        "Bolalar jismoniy tarbiya tizimida turli xil vositalar "
        "qo'llaniladi. Vositalar — bu jismoniy tarbiya maqsad va "
        "vazifalarini amalga oshirish uchun ishlatiladigan pedagogik "
        "ta'sir yo'llaridir. Birinchi va asosiy vosita — jismoniy "
        "mashqlar. Jismoniy mashqlar — bu maxsus maqsad bilan "
        "bajariladigan harakat faoliyati turlari bo'lib, ular "
        "bolaning jismoniy rivojlanishini ta'minlaydi.",

        "Jismoniy mashqlar bir nechta turlarga bo'linadi. Gimnastika "
        "mashqlari — umumrivojlantiruvchi mashqlar, asosiy harakat "
        "turlari, qurilish mashqlari va ritm gimnastikasi. Sport "
        "mashqlari — velosipedda yurish, changida yurish, suzish "
        "elementlari. Harakatli o'yinlar — syujetli va syujetsiz "
        "o'yinlar, sport elementli o'yinlar. Turistik elementlar — "
        "piyoda sayohat, tabiatda orientatsiya.",

        "Ikkinchi vosita — tabiiy omillar. Bularga quyosh nuri, "
        "havo va suv kiradi. Bu omillar bolaning organizmini "
        "chiniqtirish, immunitetni mustahkamlash va sog'lig'ini "
        "yaxshilash uchun ishlatiladi. Quyosh vannalari, ochiq havoda "
        "mashg'ulotlar, suv bilan yuvinish va cho'milish kabi "
        "chiniqtirish turlari qo'llaniladi. Tabiiy omillar jismoniy "
        "mashqlar bilan birgalikda qo'llanganda ularning samaradorligi "
        "oshadi.",

        "Uchinchi vosita — gigiyenik omillar. Bularga to'g'ri kun "
        "tartibi, muvozanatli ovqatlanish, yetarli uyqu, shaxsiy "
        "gigiyena, xonaning havosi va harorati, kiyim-boshning "
        "qulayligi kiradi. Bu omillar bolaning sog'lig'ini saqlash "
        "va jismoniy tarbiya jarayonining samaradorligini oshirish "
        "uchun zarur shart-sharoit yaratadi.",

        "To'rtinchi vosita — sport inventari va jihozlar. Maktabgacha "
        "ta'lim muassasalarida jismoniy tarbiya uchun turli xil "
        "jihozlar ishlatiladi: gimnastika devorlari, skameykalar, "
        "koptoklar, arqonlar, halqalar, prujinali taxtalar, "
        "balansirovka tashlar va boshqalar. Bu jihozlar bolaning "
        "yoshiga mos, xavfsiz va sifatli bo'lishi kerak. Zamonaviy "
        "maktabgacha ta'limda fitball, batut, velosiped kabi yangi "
        "jihozlar ham keng qo'llanilmoqda. Barcha vositalar "
        "birgalikda jismoniy tarbiya tizimining samaradorligini "
        "ta'minlaydi.",
    ],

    # 8. Zamonaviy rivojlanish yo'nalishlari
    [
        "Bolalar jismoniy tarbiya fani zamonaviy davrda bir qator "
        "yangi yo'nalishlar bo'yicha rivojlanmoqda. Birinchi yo'nalish — "
        "sog'liqni saqlash texnologiyalari (zdorovyesberegayushchiye "
        "texnologii). Bu yondashuv jismoniy tarbiya jarayonida bolaning "
        "sog'lig'ini nafaqat mustahkamlash, balki avvalo zarar "
        "yetkazmaslik tamoyilini asosiy o'ringa qo'yadi. Har bir "
        "mashq va yuklamaning bolaga ta'siri sinchkovlik bilan "
        "kuzatiladi.",

        "Ikkinchi yo'nalish — individual va differentsial yondashuv. "
        "Zamonaviy fan har bir bolaning jismoniy rivojlanishi "
        "individual tezlikda kechishini tan oladi. Shu asosda "
        "jismoniy tarbiya mashg'ulotlarida turli sog'lik guruhlari, "
        "jismoniy tayyorgarlik darajalari va individual xususiyatlar "
        "hisobga olinadi. Har bir bola o'z imkoniyatlariga mos "
        "yuklamalar oladi.",

        "Uchinchi yo'nalish — innovatsion jismoniy tarbiya texnologiyalari. "
        "Zamonaviy amaliyotda fitball-gimnastika, step-aerobika "
        "elementlari, yoga elementlari, ritmik gimnastika, suv "
        "gimnastikasi kabi yangi shakllar joriy etilmoqda. Bu "
        "texnologiyalar bolalarning qiziqishini oshiradi va "
        "an'anaviy mashg'ulotlarni boyitadi. Shuningdek, interaktiv "
        "o'yinlar va raqamli texnologiyalar ham jismoniy faollikni "
        "rag'batlantirish uchun qo'llanilmoqda.",

        "To'rtinchi yo'nalish — inklyuziv jismoniy tarbiya. Bu "
        "yo'nalish maxsus ehtiyojli bolalarning ham umumiy jismoniy "
        "tarbiya jarayoniga qo'shilishini ta'minlaydi. Ular uchun "
        "moslashtirilgan mashqlar, maxsus jihozlar va individual "
        "yondashuvlar ishlab chiqilmoqda. Beshinchi yo'nalish — "
        "oilaviy jismoniy tarbiya. Ota-onalarni bolaning jismoniy "
        "rivojlanishiga faol jalb etish, oilaviy sport tadbirlari "
        "va uy sharoitida jismoniy mashqlar tizimini yaratish "
        "zamonaviy fanning muhim yo'nalishidir.",

        "Oltinchi yo'nalish — ilmiy-metodik ta'minotni "
        "takomillashtirish. Yangi o'quv dasturlari, metodik "
        "qo'llanmalar, raqamli resurslar va onlayn platformalar "
        "yaratilmoqda. Tarbiyachilarning malakasini oshirish "
        "dasturlari yangilanmoqda. Yettinchi yo'nalish — monitoring "
        "va baholash tizimini zamonaviylashtirish. Bolalarning "
        "jismoniy rivojlanishini raqamli vositalar orqali kuzatish, "
        "elektron portfoliolar yuritish va natijalarni tahlil qilish "
        "tizimlari joriy etilmoqda.",
    ],

    # 9. Xulosa
    [
        "Mazkur mustaqil ishda bolalar jismoniy tarbiya fanining "
        "rivojlanish tizimi har tomonlama tahlil qilindi. Tahlillar "
        "shuni ko'rsatdiki, bu fan uzoq tarixiy rivojlanish "
        "jarayonidan o'tib, bugungi kunda mustaqil, ilmiy asoslangan "
        "va amaliy jihatdan sinovdan o'tgan pedagogik fan sifatida "
        "shakllangan.",

        "Jismoniy tarbiya fanining tizimi ko'p qamrovli bo'lib, u "
        "g'oyaviy-nazariy, dasturiy-me'yoriy, tashkiliy, kadrlar "
        "tayyorlash va moddiy-texnik komponentlardan iborat. Bu "
        "komponentlar bir-birini to'ldiradi va birgalikda samarali "
        "jismoniy tarbiya jarayonini ta'minlaydi. Tizimning maqsad "
        "va vazifalari aniq belgilangan, tamoyillari ilmiy asoslangan.",

        "Maktabgacha ta'limda jismoniy tarbiya turli tashkiliy "
        "shakllar — mashg'ulotlar, gimnastika, o'yinlar, bayramlar "
        "orqali amalga oshiriladi. Vositalar — jismoniy mashqlar, "
        "tabiiy omillar, gigiyenik sharoitlar va sport jihozlari — "
        "birgalikda bolaning jismoniy kamolotini ta'minlaydi.",

        "Xulosa qilib aytganda, bolalar jismoniy tarbiya fanining "
        "rivojlanish tizimi — bu ilmiy asoslangan, tarixan "
        "shakllangan va doimiy takomillashib borayotgan yaxlit "
        "pedagogik tizimdir. Zamonaviy yo'nalishlar — sog'liqni "
        "saqlash texnologiyalari, individual yondashuv, innovatsion "
        "shakllar va inklyuziv ta'lim — bu tizimni yanada "
        "boyitmoqda va bolalar jismoniy tarbiyasining sifatini "
        "oshirishga xizmat qilmoqda.",

        "Foydalanilgan adabiyotlar:",
        "1. O'zbekiston Respublikasining \"Jismoniy tarbiya va sport "
        "to'g'risida\"gi Qonuni. — Toshkent, 2015.",
        "2. Maktabgacha ta'limning Davlat o'quv dasturi \"Ilk qadam\". — "
        "Toshkent, 2018.",
        "3. Xolmatov A. \"Bolalar jismoniy tarbiyasi nazariyasi va "
        "metodikasi\". — Toshkent, 2020.",
        "4. Abdullayev B. \"Maktabgacha yoshdagi bolalarning jismoniy "
        "tarbiyasi\". — Toshkent, 2021.",
        "5. Stepanenkova E.Ya. \"Teoriya i metodika fizicheskogo "
        "vospitaniya i razvitiya rebyonka\". — Moskva, 2018.",
        "6. Keneman A.V., Xuxlayeva D.V. \"Teoriya i metodika "
        "fizicheskogo vospitaniya detey doshkolnogo vozrasta\". — "
        "Moskva, 2016.",
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
    "Bolalar jismoniy tarbiya fanining rivojlanish tizimi",
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
<dc:title>Mustaqil ish - Bolalar jismoniy tarbiya rivojlanish tizimi</dc:title>
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

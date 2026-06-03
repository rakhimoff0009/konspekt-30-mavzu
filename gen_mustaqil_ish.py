#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build a .docx (Office Open XML) file manually using only stdlib (zipfile).
Topic: Guruh yoshi nisbatida rivojlantiruvchi muhitning tashkil etilishi.
       Rivojlantiruvchi markazlari tasnifi
Target: ~9 A4 pages of Uzbek text content.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT = "/projects/sandbox/Mustaqil_ish.docx"

# ---------- CONTENT ----------
TITLE = "MUSTAQIL ISH"
SUBTITLE = ("Mavzu: Guruh yoshi nisbatida rivojlantiruvchi muhitning "
            "tashkil etilishi. Rivojlantiruvchi markazlari tasnifi")

PLAN = [
    "Kirish",
    "Rivojlantiruvchi muhit tushunchasi va uning maqsadi",
    "Guruh yoshiga mos rivojlantiruvchi muhitni tashkil etish tamoyillari",
    "Erta yosh va kichik guruhlar uchun rivojlantiruvchi muhit",
    "O'rta va katta guruhlar uchun rivojlantiruvchi muhit",
    "Maktabga tayyorlov guruhi uchun rivojlantiruvchi muhit",
    "Rivojlantiruvchi markazlar tasnifi",
    "Rivojlantiruvchi markazlarning tarbiyaviy ahamiyati",
    "Xulosa va foydalanilgan adabiyotlar",
]

# Each section: list of paragraphs (plain text, Uzbek). No page numbers in plan.
SECTIONS = [
    # 1. Kirish
    [
        "Bugungi kunda maktabgacha ta'lim sohasida bolaning shaxs sifatida har "
        "tomonlama rivojlanishi uchun zarur sharoitlarni yaratish dolzarb "
        "masalalardan biri bo'lib qolmoqda. Bola hayotining dastlabki yillari "
        "uning aqliy, jismoniy, hissiy va ijtimoiy rivojlanishi uchun eng muhim "
        "davr hisoblanadi. Aynan shu davrda bola atrof-muhitni faol o'rganadi, "
        "tajriba to'playdi va dunyoqarashini shakllantiradi. Shu sababli "
        "maktabgacha ta'lim muassasalarida bolalarning yosh xususiyatlariga mos "
        "rivojlantiruvchi muhitni tashkil etish katta ahamiyatga ega.",

        "Rivojlantiruvchi muhit deganda bolaning erkin faoliyat yuritishi, "
        "o'ynash, izlanish, kashf etish va ijodiy fikrlash uchun maxsus "
        "tashkillashtirilgan jihozlangan makon tushuniladi. Bunday muhit "
        "bolaning qiziqishlarini hisobga olgan holda, uning shaxsiy "
        "imkoniyatlarini namoyon qilishiga zamin yaratadi. To'g'ri tashkil "
        "etilgan rivojlantiruvchi muhit tarbiyachining ishini ham yengillashtiradi, "
        "chunki bola mustaqil faoliyat orqali ko'p narsalarni o'rganib oladi.",

        "Mazkur mustaqil ishda guruh yoshi nisbatida rivojlantiruvchi muhitning "
        "qanday tashkil etilishi, har bir yosh bosqichida qo'yiladigan talablar, "
        "shuningdek, rivojlantiruvchi markazlarning turlari va ularning tasnifi "
        "batafsil yoritib beriladi. Ish davomida nazariy ma'lumotlar, amaliy "
        "tavsiyalar va pedagogik tajribalardan foydalanildi. Mavzuning "
        "dolzarbligi shundaki, hozirgi davlat ta'lim standartlari ham "
        "rivojlantiruvchi muhit talablarini yangicha yondashuv asosida belgilab "
        "bermoqda.",
    ],

    # 2. Rivojlantiruvchi muhit tushunchasi
    [
        "Rivojlantiruvchi muhit — bu bolaning shaxsiy rivojlanishini "
        "ta'minlovchi pedagogik, psixologik va moddiy-texnik shart-sharoitlar "
        "majmuidir. U faqat o'yinchoqlar yoki jihozlardan iborat bo'lmasdan, "
        "balki bolaning hissiyotlari, fikrlari va munosabatlariga ta'sir "
        "ko'rsatadigan butun atrof olamni o'z ichiga oladi. Rivojlantiruvchi "
        "muhitning asosiy maqsadi — bolada mustaqillik, tashabbuskorlik, "
        "ijodkorlik va o'z-o'zini namoyon qilish ko'nikmalarini shakllantirishdir.",

        "Rivojlantiruvchi muhit quyidagi vazifalarni bajaradi: bolaning bilim "
        "doirasini kengaytiradi; nutqi, tafakkuri va xotirasini rivojlantiradi; "
        "harakat ko'nikmalarini takomillashtiradi; ijodiy qobiliyatlarini "
        "yuzaga chiqaradi; tengdoshlari va kattalar bilan muloqot qilishga "
        "o'rgatadi; estetik did va madaniy qadriyatlarni shakllantiradi. Bundan "
        "tashqari, muhit bolaga xavfsiz va qulay bo'lishi shart.",

        "Pedagogika fanida rivojlantiruvchi muhitning bir nechta tamoyillari "
        "ajratiladi. Bular: ochiqlik tamoyili — bola istalgan paytda istalgan "
        "markazga o'tib, faoliyat yurita olishi; o'zgaruvchanlik tamoyili — "
        "muhit doimiy yangilanib turishi; faollik tamoyili — bolani harakatga "
        "undaydigan vositalar mavjudligi; mustaqillik tamoyili — bola yordamsiz "
        "ham o'ynashi mumkin bo'lishi; xavfsizlik tamoyili — barcha jihozlar "
        "bola sog'lig'iga zarar yetkazmasligi.",

        "Shuningdek, rivojlantiruvchi muhit bolaning yoshi, jinsi, qiziqishlari "
        "va individual xususiyatlariga moslashtirilgan bo'lishi kerak. Masalan, "
        "kichik yoshdagi bolalar uchun yorqin ranglar, yumshoq materiallar va "
        "katta o'lchamdagi o'yinchoqlar mos kelsa, katta yoshdagi bolalar uchun "
        "mantiqiy o'yinlar, konstruktorlar va kitoblar muhim hisoblanadi. "
        "Tarbiyachi bu xususiyatlarni hisobga olgan holda guruhdagi muhitni "
        "loyihalashi lozim.",
    ],

    # 3. Tamoyillar
    [
        "Guruh yoshiga mos rivojlantiruvchi muhitni tashkil etishda bir qator "
        "muhim tamoyillarga rioya qilinadi. Birinchi tamoyil — yoshga mosligi. "
        "Har bir yosh davrining o'ziga xos psixologik va fiziologik "
        "xususiyatlari mavjud. Tarbiyachi guruhdagi bolalarning yoshini hisobga "
        "olib, ularga mos jihozlar, o'yinchoqlar va didaktik vositalarni tanlashi "
        "kerak. Aks holda muhit bola uchun foydali bo'lmaydi.",

        "Ikkinchi tamoyil — funksionallik. Guruhdagi har bir burchak yoki "
        "markaz aniq vazifani bajarishi lozim. Masalan, kitob burchagi bolaning "
        "nutqini va savodxonligini, qurilish burchagi esa fazoviy tasavvur va "
        "mayda qo'l harakatlarini rivojlantirishga xizmat qiladi. Funksional "
        "muhit bolaga tanlash erkinligini beradi va uning qiziqishlarini "
        "qondiradi.",

        "Uchinchi tamoyil — estetiklik. Bola go'zal, ozoda va did bilan "
        "jihozlangan muhitda o'sganida, unda go'zallikni his qilish tuyg'usi "
        "shakllanadi. Devorlardagi rasmlar, jihozlarning rangi, mebel "
        "joylashuvi — barchasi estetik talablarga javob berishi lozim. "
        "To'rtinchi tamoyil — moslashuvchanlik. Muhit zarurat tug'ilganda "
        "qaytadan tashkil etilishi, jihozlar joyini almashtirish mumkin "
        "bo'lishi kerak.",

        "Beshinchi tamoyil — integrativlik. Bu tamoyilga ko'ra, turli "
        "rivojlantiruvchi markazlar bir-biri bilan o'zaro bog'liq holda "
        "ishlaydi. Masalan, bola tabiat burchagida o'simliklarni kuzatib, "
        "keyin badiiy ijod burchagiga o'tib, ularni rasm chizishi mumkin. "
        "Oltinchi tamoyil — gender yondashuvi. O'g'il va qiz bolalar uchun "
        "qiziqarli bo'lgan o'yinchoqlar va materiallar mavjud bo'lishi kerak. "
        "Bu tamoyillar hammasi birgalikda samarali muhit yaratish imkonini "
        "beradi.",

        "Tashkil etishda yana bir muhim jihat — bu xavfsizlik. Barcha jihozlar "
        "bolalar uchun zararsiz materiallardan tayyorlangan, o'tkir burchaklari "
        "bo'lmagan, mustahkam o'rnatilgan bo'lishi shart. Elektr rozetkalari "
        "himoyalangan, kichik detallar bola og'ziga ketib qolmaydigan tarzda "
        "saqlanishi kerak. Bularning barchasi tarbiyachining mas'uliyati "
        "doirasiga kiradi.",
    ],

    # 4. Erta yosh va kichik guruhlar
    [
        "Erta yosh guruhi (1–3 yosh) bolalari uchun rivojlantiruvchi muhit "
        "alohida e'tibor talab qiladi. Bu yoshda bola harakatchan, hamma "
        "narsani sinab ko'rishni, ushlab, tatib ko'rishni xohlaydi. Shu sababli "
        "guruh xonasida keng, ochiq joy bo'lishi, bolaning erkin yura olishi "
        "uchun sharoit yaratilishi kerak. Mebellar past, yumshoq qirralarga ega "
        "bo'lishi, o'yinchoqlar yorqin ranglarda bo'lishi tavsiya etiladi.",

        "Erta yosh guruhida sensor rivojlanishga alohida ahamiyat beriladi. "
        "Bola turli ranglar, shakllar, hidlar, tovushlar va sirt teksturalari "
        "bilan tanishishi kerak. Buning uchun maxsus sensor o'yinchoqlar, "
        "piramidalar, kublar, yumshoq kitoblar, musiqali o'yinchoqlar tayyor "
        "holda turishi lozim. Shuningdek, bu yoshda bolaning nutqi endigina "
        "shakllanayotgani uchun rasmli kitoblar va kartochkalar muhim rol "
        "o'ynaydi.",

        "Kichik guruh (3–4 yosh) bolalari uchun muhit yanada boyitiladi. "
        "Ushbu yoshda bola syujetli o'yinlarga qiziqa boshlaydi: \"Oila\", "
        "\"Shifokor\", \"Do'kon\" kabi rolli o'yinlar uchun maxsus burchaklar "
        "tashkil etiladi. Qo'g'irchoqlar, idishlar, kichik mebellar, "
        "shifokorlik to'plami, savatchalar — bularning hammasi bola "
        "tomonidan qo'lay foydalanish uchun joylashtiriladi.",

        "Kichik guruhda mayda qo'l harakatlarini rivojlantiruvchi vositalar "
        "ko'p bo'lishi kerak. Masalan, yog'och taxtachalar, shnurli "
        "o'yinchoqlar, mozaika, yirik bolg'a-mix to'plamlari va boshqalar. "
        "Shuningdek, badiiy ijod burchagi tashkil etilib, unda bo'yoqlar, "
        "qog'ozlar, plastilin va qalamlar bola qo'li yetadigan joyda "
        "saqlanadi. Bola istalgan vaqtda kelib, ijod qilishi mumkin.",

        "Erta va kichik yosh guruhlarida tarbiyachining roli juda katta. U "
        "bolaga doimiy yordam beradi, namuna ko'rsatadi, qo'llab-quvvatlaydi. "
        "Bola tarbiyachi yonida o'zini xavfsiz his qilishi va muhit bilan "
        "tanishish jarayonida hech qanday qo'rquv tuyishmasligi kerak. Aynan "
        "shunday muhit bolaning dastlabki ijtimoiylashuvi uchun zamin "
        "yaratadi.",
    ],

    # 5. O'rta va katta guruhlar
    [
        "O'rta guruh (4–5 yosh) bolalari uchun rivojlantiruvchi muhit yanada "
        "murakkablashadi. Bu yoshdagi bolalar tengdoshlari bilan birgalikda "
        "o'ynashga, qoidali o'yinlarga va izlanishga qiziqa boshlaydi. Ularning "
        "tasavvur olami kengayadi, savollari ko'payadi. Shu sababli guruh "
        "xonasida turli mavzudagi rivojlantiruvchi markazlar tashkil etiladi: "
        "tabiat burchagi, kitob burchagi, eksperiment burchagi, qurilish "
        "burchagi va boshqalar.",

        "O'rta guruhda matematik tushunchalarni shakllantiruvchi vositalar "
        "muhim o'rin tutadi. Geometrik shakllar, raqamli kartochkalar, "
        "sanoq tayoqchalari, o'lchov asboblari guruhda joylashtiriladi. "
        "Bola o'yin orqali sonlar, miqdorlar, kattalik va shakl haqida "
        "tushunchaga ega bo'ladi. Shuningdek, mantiqiy fikrlashni "
        "rivojlantiruvchi puzzllar, labirintlar va didaktik o'yinlar muhitning "
        "ajralmas qismi hisoblanadi.",

        "Katta guruh (5–6 yosh) bolalari uchun muhit allaqachon o'quv "
        "faoliyatiga yo'naltirila boshlaydi. Bu yoshda bola harflar, raqamlar "
        "bilan faol tanishadi, mustaqil o'qish va yozish ko'nikmalarining "
        "asoslari shakllanadi. Guruhda savodxonlik markazi tashkil etiladi: "
        "alfavit kartochkalari, magnit harflari, kitoblar, rasm-suratli "
        "lug'atlar joylashtiriladi.",

        "Katta guruhda bolalarning tadqiqotchilik faoliyatini qo'llab-quvvatlash "
        "uchun maxsus eksperiment markazi yaratiladi. Bu yerda lupa, magnit, "
        "tarozi, suv, qum, turli xil tabiiy materiallar bo'ladi. Bola sodda "
        "tajribalar o'tkazib, sabab-natija aloqalarini tushunishni o'rganadi. "
        "Shu tariqa unda ilmiy fikrlash asoslari shakllanadi.",

        "O'rta va katta yosh guruhlarida bolalarning ijtimoiy ko'nikmalarini "
        "rivojlantirishga ham katta e'tibor qaratiladi. Birgalikda o'ynaladigan "
        "stol o'yinlari, jamoaviy loyihalar uchun maxsus stollar, do'stlik "
        "burchagi, kayfiyat taxtasi kabi yangiliklar joriy etiladi. Bola "
        "o'zining va tengdoshlarining hissiyotlarini tushunishni o'rganadi, "
        "muloqot madaniyati shakllanadi.",
    ],

    # 6. Maktabga tayyorlov guruhi
    [
        "Maktabga tayyorlov guruhi (6–7 yosh) bolaning maktabgacha ta'lim "
        "muassasasidagi yakuniy bosqichi hisoblanadi. Bu davrda bolada o'quv "
        "faoliyatiga tayyorgarlik, mas'uliyat hissi, e'tiborni jamlash va "
        "mustaqil ishlash ko'nikmalari shakllanishi kerak. Shu sababli "
        "rivojlantiruvchi muhit ham maktabga moslashuvga yordam beradigan "
        "tarzda tashkil etiladi.",

        "Tayyorlov guruhida o'quv markazi alohida ahamiyatga ega. Bu yerda "
        "har bir bolaga mo'ljallangan stol va stul, daftar, qalam, ruchka, "
        "chizg'ich, o'chirg'ich va boshqa o'quv qurollari bo'ladi. Bola maktab "
        "muhitiga o'xshash sharoitda mashg'ulotlarga o'rganadi. Devorda alfavit, "
        "raqamlar jadvali, kun va oy nomlari, ob-havo taqvimi osilgan bo'ladi.",

        "Bu yoshda bolaning savodxonligini chuqurlashtirish maqsadida boy "
        "kutubxona burchagi yaratiladi. Unda turli janrdagi bolalar adabiyoti, "
        "ertaklar, she'rlar, qomuslar va o'rgatuvchi kitoblar joylashtiriladi. "
        "Bola istalgan vaqtda kitob olib, mustaqil o'qishi yoki rasmlarini "
        "ko'rishi mumkin. Shuningdek, bolaning nutqini boyitish uchun "
        "topishmoqlar, maqollar va matal kartochkalari ishlatiladi.",

        "Maktabga tayyorlov guruhida ijodiy markazlar ham yanada "
        "takomillashtiriladi. Bola o'zi mustaqil loyiha tayyorlashi, "
        "applikatsiya yasashi, plastilindan haykalcha yaratishi, hatto kichik "
        "spektakllar uyushtirishi mumkin. Bunday faoliyat unda tashabbuskorlik, "
        "ijodkorlik va o'z-o'zini ifodalash imkoniyatlarini rivojlantiradi.",

        "Bundan tashqari, tayyorlov guruhida bolaning vaqtni rejalashtirish, "
        "kun tartibiga rioya qilish ko'nikmalari ham mustahkamlanadi. Maxsus "
        "vaqt taxtasi, kun tartibi sxemasi, navbatchilik jadvali tayyorlanadi. "
        "Bola o'z faoliyatini rejalashtirib, navbatchilik vazifalarini bajarib, "
        "mustaqillik va mas'uliyatni o'rganadi. Bu maktabdagi keyingi hayotga "
        "muhim asos bo'lib xizmat qiladi.",
    ],

    # 7. Markazlar tasnifi
    [
        "Maktabgacha ta'lim muassasalarida tashkil etiladigan rivojlantiruvchi "
        "markazlarni quyidagi turlarga bo'lish mumkin. Bu tasnif zamonaviy "
        "pedagogika va davlat ta'lim standartlari asosida shakllantirilgan "
        "bo'lib, har bir markaz bolaning muayyan rivojlanish sohasini "
        "qo'llab-quvvatlashga xizmat qiladi.",

        "1. Nutq va savodxonlik markazi. Bu markaz bolaning og'zaki va yozma "
        "nutqini, lug'at boyligini, fonematik eshitishini va savodxonlikka "
        "qiziqishini rivojlantiradi. Markaz tarkibida kitoblar, harf "
        "kartochkalari, syujetli rasmlar, audio-ertaklar va nutq o'yinlari "
        "bo'ladi.",

        "2. Matematika va mantiq markazi. Markaz bolada matematik tushunchalar, "
        "raqam, sanoq, miqdor, shakl, kattalik haqidagi bilimlarni "
        "shakllantiradi. Sanoq tayoqchalari, geometrik figuralar, magnit "
        "raqamlar, mantiqiy o'yinlar, puzzllar va jadvallar markazning asosiy "
        "vositalari hisoblanadi.",

        "3. Tabiat va ekologiya markazi. Bu markaz bolada atrof-muhitga "
        "qiziqish, tirik tabiatga muhabbat va ekologik madaniyat hissini "
        "uyg'otadi. Xona o'simliklari, akvarium, tabiat taqvimi, hayvonlar "
        "rasmlari, tabiiy materiallar (toshlar, qobiqlar, barglar) markazda "
        "joylashtiriladi.",

        "4. Eksperiment va tadqiqot markazi. Bola sodda tajribalar o'tkazish "
        "orqali olamni o'rganadi. Markazda lupa, magnit, tarozi, mensurka, "
        "suv, qum, tuproq, turli moddalar va tajriba uchun zarur asboblar "
        "bo'ladi. Bola \"nima uchun?\", \"qanday qilib?\" degan savollarga "
        "amaliy javob topishni o'rganadi.",

        "5. Badiiy ijod va estetika markazi. Bu markaz bolaning ijodiy "
        "qobiliyatlarini, badiiy didini va estetik tafakkurini rivojlantiradi. "
        "Bo'yoqlar, qog'ozlar, qaychi, yelim, plastilin, gips, mato bo'laklari, "
        "tabiiy materiallardan tayyorlangan kollajlar uchun jihozlar markazning "
        "ajralmas qismidir.",

        "6. Musiqa va teatr markazi. Markazda turli musiqa asboblari "
        "(metallofon, baraban, marakas, bolalar pianinosi), audio-yozuvlar, "
        "qo'g'irchoq teatri, niqoblar va sahna jihozlari bo'ladi. Bola musiqa "
        "tinglaydi, sahna ko'rinishlarini namoyish etadi va o'zining badiiy "
        "qobiliyatlarini namoyon qiladi.",

        "7. Qurilish va konstruktorlik markazi. Bu markaz bolaning fazoviy "
        "tasavvurini, mantiqiy fikrlashini va mayda qo'l harakatlarini "
        "rivojlantiradi. Yog'och kublar, plastik konstruktorlar (Lego, Magformers "
        "kabi), katta yumshoq bloklar, sxemalar va loyihalar uchun rasmlar "
        "markazda saqlanadi.",

        "8. Rolli va syujetli o'yinlar markazi. Markazda \"Oila\", \"Shifokor\", "
        "\"Do'kon\", \"Sartaroshxona\", \"Oshxona\" kabi syujetli o'yinlar uchun "
        "kerakli atributlar — qo'g'irchoqlar, idishlar, kiyimlar, kichik "
        "mebellar joylashtiriladi. Bola katta dunyoni o'yin orqali tushunadi, "
        "ijtimoiy rollarni o'zlashtiradi.",

        "9. Sport va sog'lomlashtirish markazi. Bu markaz bolaning jismoniy "
        "rivojlanishini ta'minlaydi. Koptoklar, arqonlar, halqalar, kichik "
        "gimnastika gilamchalari, sport o'yinlari uchun jihozlar bu yerda "
        "saqlanadi. Bola harakatchanlik, chaqqonlik va sportga muhabbat "
        "ko'nikmalarini hosil qiladi.",
    ],

    # 8. Tarbiyaviy ahamiyati
    [
        "Rivojlantiruvchi markazlar bolaning hayotida juda katta tarbiyaviy "
        "ahamiyatga ega. Avvalo, ular bolada mustaqil faoliyat yuritish "
        "ko'nikmasini shakllantiradi. Bola tarbiyachi yo'l-yo'riq berishini "
        "kutmasdan, o'zi qiziqqan markazga borib, faoliyat tanlaydi va uni "
        "amalga oshiradi. Bu jarayonda u o'z-o'zini boshqarish, vaqtni taqsimlash "
        "va natijaga erishish qobiliyatlarini egallaydi.",

        "Markazlar bolada tanlash erkinligini tarbiyalaydi. Har kuni o'sha xil "
        "mashg'ulot bilan shug'ullanish bola uchun zerikarli bo'lishi mumkin. "
        "Lekin guruhda turli xil markazlar bo'lganda, bola o'z kayfiyati va "
        "qiziqishlariga qarab faoliyat tanlay oladi. Bu uning ichki "
        "motivatsiyasini kuchaytiradi va o'rganish jarayonini quvonchli "
        "tajribaga aylantiradi.",

        "Shuningdek, rivojlantiruvchi markazlar bolada ijtimoiy ko'nikmalarni "
        "shakllantiradi. Markazlarda ko'pincha bir necha bola birga ishlaydi. "
        "Ular bir-biri bilan kelishish, navbat kutish, fikr almashish, "
        "yordam berish kabi ijtimoiy malakalarni o'rganadi. Bu malakalar "
        "kelajakda jamiyatda muvaffaqiyatli faoliyat ko'rsatish uchun zarur "
        "asos bo'lib xizmat qiladi.",

        "Markazlar bolaning emotsional rivojlanishiga ham ijobiy ta'sir "
        "ko'rsatadi. Bola o'z mehnati natijasini ko'rganda — chizilgan rasm, "
        "qurilgan minora, tayyorlangan ovqat (o'yinda) — quvonadi, faxrlanadi. "
        "Bu unda ijobiy hissiyotlarni mustahkamlaydi va o'ziga ishonch hissini "
        "shakllantiradi. Muvaffaqiyatsizlikka uchraganda esa qayta urinish, "
        "sabr va matonatni o'rganadi.",

        "Pedagogik nuqtai nazardan rivojlantiruvchi markazlar tarbiyachiga "
        "individual yondashuvni amalga oshirish imkonini beradi. Tarbiyachi "
        "har bir bolani kuzatib, uning qaysi sohada ko'proq qiziqishi yoki "
        "qiyinchilik chekayotganini aniqlay oladi. Shu asosda u bola bilan "
        "individual ishlaydi, kerakli vositalarni taklif qiladi va uning "
        "rivojlanishini qo'llab-quvvatlaydi.",

        "Bundan tashqari, markazlar ota-onalar bilan hamkorlik qilish uchun "
        "ham yaxshi vosita hisoblanadi. Ota-onalar guruhga kelganida, bolaning "
        "qaysi markazda qanday faoliyat ko'rsatayotgani bilan tanishadi, uning "
        "yutuqlarini ko'radi. Bu oilaning ham bola tarbiyasiga faolroq "
        "qatnashishiga turtki beradi va bog'cha bilan oila o'rtasida ishonchli "
        "munosabat o'rnatishga yordam beradi.",
    ],

    # 9. Xulosa
    [
        "Yuqorida keltirilgan ma'lumotlar shundan dalolat beradiki, "
        "maktabgacha ta'lim muassasasida guruh yoshi nisbatida tashkil etilgan "
        "rivojlantiruvchi muhit bola shaxsining har tomonlama "
        "rivojlanishi uchun eng muhim shartlardan biridir. Har bir yosh "
        "bosqichida muhitga qo'yiladigan talablar o'ziga xos bo'lib, ular "
        "bolaning psixologik va fiziologik xususiyatlariga muvofiq "
        "shakllantirilishi lozim.",

        "Erta yoshdan boshlab to maktabga tayyorlov guruhigacha bo'lgan "
        "bosqichlarda muhit asta-sekin murakkablashib boradi: yorqin sensor "
        "vositalardan tortib, eksperiment asboblari va o'quv qurollarigacha "
        "bo'lgan keng vositalar guruhga kiritiladi. Tarbiyachi bu jarayonda "
        "bolaning ehtiyojlarini tushunib, muhitni doimiy yangilab, boyitib "
        "borishi kerak.",

        "Rivojlantiruvchi markazlarning tasnifi — nutq, matematika, tabiat, "
        "eksperiment, badiiy ijod, musiqa, qurilish, rolli o'yinlar va sport "
        "markazlari — bolaning barcha rivojlanish sohalarini qamrab oladi. "
        "Har bir markaz o'z vazifasiga ega bo'lib, ular birgalikda bolaning "
        "uyg'un shaxs sifatida shakllanishiga xizmat qiladi. Bunday yondashuv "
        "zamonaviy maktabgacha ta'limning asosiy yo'nalishlaridan biri "
        "hisoblanadi.",

        "Xulosa qilib aytganda, rivojlantiruvchi muhit va markazlar — bu "
        "shunchaki o'yinchoqlar to'plami emas, balki bolaning kelajagini "
        "shakllantiruvchi muhim pedagogik tizimdir. Tarbiyachining vazifasi — "
        "bu tizimni mahorat bilan boshqarib, har bir bolaga o'z imkoniyatlarini "
        "namoyon qilish uchun sharoit yaratib berishdir. Aynan shunday muhit "
        "kelajakda jamiyatga foydali, har tomonlama rivojlangan, ijodkor va "
        "mustaqil shaxslarni tarbiyalashga zamin yaratadi.",

        "Foydalanilgan adabiyotlar:",
        "1. O'zbekiston Respublikasining \"Ta'lim to'g'risida\"gi Qonuni. — "
        "Toshkent, 2020.",
        "2. Maktabgacha ta'lim Davlat o'quv dasturi \"Ilk qadam\". — Toshkent, "
        "2018.",
        "3. Sodiqova Sh., Shoumarova M. \"Maktabgacha pedagogika\" o'quv "
        "qo'llanmasi. — Toshkent, 2019.",
        "4. Qodirova F. \"Bolalar bog'chasida rivojlantiruvchi muhitni tashkil "
        "etish\". — Toshkent, 2021.",
        "5. Internet manbalari: uzedu.uz, eduportal.uz.",
    ],
]

# ---------- DOCX BUILDER ----------

def make_para(text, bold=False, size=28, align=None, before=120, after=120):
    """Create a w:p XML string. size is half-points (28 = 14pt)."""
    align_xml = f'<w:jc w:val="{align}"/>' if align else ""
    bold_xml = '<w:b/>' if bold else ""
    safe = escape(text)
    return (
        '<w:p>'
        '<w:pPr>'
        f'<w:spacing w:before="{before}" w:after="{after}" w:line="360" w:lineRule="auto"/>'
        f'{align_xml}'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
        'w:cs="Times New Roman"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        f'{bold_xml}'
        '</w:rPr>'
        '</w:pPr>'
        '<w:r>'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
        'w:cs="Times New Roman"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        f'{bold_xml}'
        '</w:rPr>'
        f'<w:t xml:space="preserve">{safe}</w:t>'
        '</w:r>'
        '</w:p>'
    )

def make_body_para(text):
    """Body paragraph: 14pt, justified, first-line indent."""
    safe = escape(text)
    return (
        '<w:p>'
        '<w:pPr>'
        '<w:spacing w:before="0" w:after="120" w:line="360" w:lineRule="auto"/>'
        '<w:ind w:firstLine="709"/>'
        '<w:jc w:val="both"/>'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
        'w:cs="Times New Roman"/>'
        '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        '</w:pPr>'
        '<w:r>'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
        'w:cs="Times New Roman"/>'
        '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        f'<w:t xml:space="preserve">{safe}</w:t>'
        '</w:r>'
        '</w:p>'
    )

def page_break_para():
    return (
        '<w:p>'
        '<w:r><w:br w:type="page"/></w:r>'
        '</w:p>'
    )

# Build body XML
body_parts = []

# ===== PAGE 1: Title page =====
# Add some empty space then title
for _ in range(4):
    body_parts.append(make_para("", size=28))

body_parts.append(make_para("O'ZBEKISTON RESPUBLIKASI",
                            bold=True, size=28, align="center"))
body_parts.append(make_para("MAKTABGACHA VA MAKTAB TA'LIMI VAZIRLIGI",
                            bold=True, size=28, align="center"))
body_parts.append(make_para("", size=24))
body_parts.append(make_para("MUSTAQIL ISH", bold=True, size=40, align="center"))
body_parts.append(make_para("", size=24))
body_parts.append(make_para("Mavzu:", bold=True, size=28, align="center"))
body_parts.append(make_para(
    "Guruh yoshi nisbatida rivojlantiruvchi muhitning tashkil etilishi. "
    "Rivojlantiruvchi markazlari tasnifi",
    bold=True, size=28, align="center"))
for _ in range(6):
    body_parts.append(make_para("", size=24))
body_parts.append(make_para("Bajardi: ______________________",
                            size=28, align="right"))
body_parts.append(make_para("Tekshirdi: ______________________",
                            size=28, align="right"))
for _ in range(2):
    body_parts.append(make_para("", size=24))
body_parts.append(make_para("Toshkent — 2026", bold=True, size=28, align="center"))
body_parts.append(page_break_para())

# ===== PAGE 2: Reja (Plan) =====
body_parts.append(make_para("REJA", bold=True, size=32, align="center"))
body_parts.append(make_para("", size=20))
for i, item in enumerate(PLAN, 1):
    body_parts.append(make_para(f"{i}. {item}", size=28))
body_parts.append(page_break_para())

# ===== PAGES 3-9: Sections =====
for idx, (heading, paragraphs) in enumerate(zip(PLAN, SECTIONS)):
    body_parts.append(make_para(f"{idx + 1}. {heading.upper()}",
                                bold=True, size=30, align="center"))
    body_parts.append(make_para("", size=18))
    for p in paragraphs:
        body_parts.append(make_body_para(p))
    # Page break after each section except the last
    if idx < len(SECTIONS) - 1:
        body_parts.append(page_break_para())

# Section properties (page size A4, margins)
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
    '<w:body>'
    + ''.join(body_parts)
    + sectPr
    + '</w:body></w:document>'
)

# ---------- minimal docx package files ----------
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
<w:docDefaults>
<w:rPrDefault>
<w:rPr>
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
<w:sz w:val="28"/>
<w:szCs w:val="28"/>
<w:lang w:val="uz-UZ"/>
</w:rPr>
</w:rPrDefault>
<w:pPrDefault>
<w:pPr>
<w:spacing w:line="360" w:lineRule="auto"/>
</w:pPr>
</w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
<w:name w:val="Normal"/>
<w:qFormat/>
</w:style>
</w:styles>'''

core_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>Mustaqil ish</dc:title>
<dc:creator>Talaba</dc:creator>
<cp:lastModifiedBy>Talaba</cp:lastModifiedBy>
</cp:coreProperties>'''

app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>Microsoft Office Word</Application>
</Properties>'''

# ---------- write zip ----------
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

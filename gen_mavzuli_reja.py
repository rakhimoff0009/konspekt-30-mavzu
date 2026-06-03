#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mustaqil ish: Mavzuli rejalashtirish metodik qo'llanmasidan foydalangan holda
rivojlanish markazlarida ta'limiy faoliyatni tashkil etish
Target: 9 A4 pages, Times New Roman 14pt, 1.5 line spacing
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT = "/projects/sandbox/Mustaqil_ish_2.docx"

PLAN = [
    "Kirish",
    "Mavzuli rejalashtirish tushunchasi va mohiyati",
    "Mavzuli rejalashtirish metodik qo'llanmasining tuzilishi",
    "Rivojlanish markazlari va ularning ta'limiy jarayondagi o'rni",
    "Mavzuli rejalashtirish asosida rivojlanish markazlarida faoliyatni tashkil etish bosqichlari",
    "Turli yosh guruhlarida mavzuli rejalashtirishdan foydalanish xususiyatlari",
    "Mavzuli rejalashtirish asosida tashkil etiladigan ta'limiy faoliyat shakllari",
    "Tarbiyachining mavzuli rejalashtirishdagi roli va vazifalari",
    "Xulosa va foydalanilgan adabiyotlar",
]

SECTIONS = [
    # 1. Kirish
    [
        "Zamonaviy maktabgacha ta'lim tizimida ta'limiy jarayonni samarali "
        "tashkil etish eng muhim pedagogik vazifalardan biri hisoblanadi. "
        "Bugungi kunda maktabgacha ta'lim muassasalarida bolalarning yosh "
        "xususiyatlariga mos ravishda rivojlantiruvchi ta'limiy muhitni "
        "shakllantirish va unda ta'limiy faoliyatni rejalashtirish dolzarb "
        "masala bo'lib turibdi. Xususan, mavzuli rejalashtirish metodik "
        "qo'llanmasidan foydalangan holda rivojlanish markazlarida ta'limiy "
        "faoliyatni tashkil etish pedagogik amaliyotda keng qo'llanilmoqda.",

        "Mavzuli rejalashtirish — bu ta'limiy jarayonni muayyan mavzu atrofida "
        "tashkil etish usuli bo'lib, unda barcha faoliyat turlari, o'yinlar, "
        "mashg'ulotlar va mustaqil ishlar bitta mavzu doirasida integratsiya "
        "qilinadi. Bu yondashuv bolaga bilimlarni parchalanmagan holda, yaxlit "
        "va mantiqiy bog'liq tarzda o'zlashtirishga yordam beradi. Rivojlanish "
        "markazlari esa bolaning mustaqil faoliyat yuritishi, izlanishi va "
        "kashf etishi uchun maxsus jihozlangan hududlardir.",

        "Mazkur mustaqil ishda mavzuli rejalashtirish metodik qo'llanmasining "
        "mohiyati, tuzilishi, rivojlanish markazlarida ta'limiy faoliyatni "
        "tashkil etish bosqichlari, turli yosh guruhlari uchun xususiyatlari "
        "va tarbiyachining roli batafsil yoritiladi. Ishning maqsadi — "
        "mavzuli rejalashtirish orqali rivojlanish markazlarida ta'limiy "
        "faoliyatni samarali tashkil etish yo'llarini o'rganish va tahlil "
        "qilishdir.",
    ],

    # 2. Mavzuli rejalashtirish tushunchasi
    [
        "Mavzuli rejalashtirish — bu maktabgacha ta'lim jarayonida ta'limiy "
        "faoliyatni muayyan mavzu yoki tematik yo'nalish atrofida tashkil etish "
        "metodidir. Bunda bir hafta yoki ikki hafta davomida barcha ta'limiy "
        "faoliyat — ertalabki davra, uyushtirilgan mashg'ulotlar, mustaqil "
        "faoliyat, sayr, o'yinlar — bitta mavzuga bo'ysundiriladi. Masalan, "
        "\"Kuz fasli\", \"Mening oilam\", \"Hayvonlar olami\", \"Transportlar\" "
        "kabi mavzular tanlanadi.",

        "Mavzuli rejalashtirishning asosiy g'oyasi shundaki, bola bir mavzu "
        "doirasida turli faoliyat turlari orqali ko'p qirrali bilim va "
        "ko'nikmalarni egallaydi. Masalan, \"Kuz fasli\" mavzusida bola tabiat "
        "burchagida barglarni kuzatadi, badiiy ijod markazida kuzgi manzara "
        "chizadi, kitob burchagida kuz haqidagi she'rlarni o'qiydi, matematik "
        "markazda barglarni sanaydi. Shunday qilib, bir mavzu turli "
        "rivojlanish sohalarini qamrab oladi.",

        "Mavzuli rejalashtirishning afzalliklari quyidagilardan iborat: "
        "birinchidan, ta'limiy jarayon izchil va mantiqiy bo'ladi; ikkinchidan, "
        "bola bilimlarni yaxlit tizim ko'rinishida o'zlashtiradi; uchinchidan, "
        "tarbiyachiga ta'limiy jarayonni oldindan rejalashtirib, zarur "
        "materiallarni tayyorlab qo'yish imkonini beradi; to'rtinchidan, "
        "ota-onalar ham mavzudan xabardor bo'lib, uyda bolaga yordam bera "
        "oladi; beshinchidan, bolaning qiziqishi uzoqroq vaqt saqlanadi.",

        "Mavzuli rejalashtirish an'anaviy kunlik rejalashtirishdan farq qiladi. "
        "An'anaviy usulda har kuni alohida mavzu olinsa, mavzuli "
        "rejalashtirishda bir mavzu bir necha kun davom etadi va har kuni "
        "yangi qirralari ochiladi. Bu bolaning e'tiborini bir mavzuga "
        "to'plashiga, chuqurroq tushunishiga va mustahkamlashiga yordam beradi. "
        "Shuningdek, bu usul integrativ ta'lim tamoyiliga to'liq javob beradi.",
    ],

    # 3. Metodik qo'llanma tuzilishi
    [
        "Mavzuli rejalashtirish metodik qo'llanmasi — bu tarbiyachi uchun "
        "ta'limiy jarayonni tashkil etishda yo'l-yo'riq bo'ladigan hujjatdir. "
        "U odatda bir o'quv yiliga mo'ljallangan bo'lib, har bir hafta yoki "
        "ikki hafta uchun alohida mavzu belgilangan. Qo'llanma tarkibida "
        "mavzuning nomi, maqsad va vazifalari, kerakli materiallar ro'yxati, "
        "rivojlanish markazlari uchun topshiriqlar va bola faoliyatini baholash "
        "mezonlari keltiriladi.",

        "Metodik qo'llanmaning birinchi qismi — yillik mavzular rejasi. Bu "
        "rejada o'quv yili davomida o'rganiladigan barcha mavzular izchil "
        "tartibda keltiriladi. Mavzular bolalarning yosh xususiyatlari, "
        "mavsumiy o'zgarishlar, bayramlar va davlat dasturi talablariga "
        "muvofiq tanlanadi. Masalan, sentyabrda \"Bog'chamiz\", oktyabrda "
        "\"Kuz va uning in'omlari\", noyabrda \"Mening Vatanim\" kabi "
        "mavzular rejalashtiriladi.",

        "Ikkinchi qismi — haftalik ta'limiy faoliyat rejasi. Bu yerda har bir "
        "kun uchun rivojlanish markazlarida qanday faoliyat tashkil etilishi "
        "batafsil ko'rsatiladi. Masalan, dushanba kuni nutq markazi uchun "
        "suhbat mavzusi, seshanba kuni matematik markaz uchun didaktik o'yin, "
        "chorshanba kuni ijod markazi uchun appliqatsiya ishi "
        "rejalashtiriladi. Har bir faoliyat uchun maqsad, kerakli material va "
        "o'tkazish usuli belgilanadi.",

        "Uchinchi qismi — rivojlanish markazlari bo'yicha topshiriqlar banki. "
        "Bu qismda har bir markaz uchun mavzuga oid turli topshiriqlar, "
        "o'yinlar va mashqlar jamlangan. Tarbiyachi bu bankdan bolalarning "
        "qiziqishi va darajasiga qarab mos topshiriqlarni tanlaydi. Masalan, "
        "\"Hayvonlar\" mavzusida tabiat markazi uchun \"Hayvonlarni "
        "guruhlarga ajrat\", ijod markazi uchun \"Hayvon niqobini yasa\", "
        "kitob markazi uchun \"Hayvonlar haqida ertak tanlash\" kabi "
        "topshiriqlar beriladi.",

        "To'rtinchi qismi — baholash va monitoring. Bu yerda bolaning har bir "
        "markazdagi faoliyati natijalarini kuzatish va qayd etish uchun "
        "jadvallar, kuzatuv kartochkalari va indikatorlar berilgan. Tarbiyachi "
        "bu vositalar yordamida har bir bolaning rivojlanish darajasini "
        "aniqlaydi va keyingi hafta uchun individual ishlar rejalashtiradi.",
    ],

    # 4. Rivojlanish markazlari va ta'limiy jarayondagi o'rni
    [
        "Rivojlanish markazlari — bu guruhdagi bolalarning turli rivojlanish "
        "sohalarini qo'llab-quvvatlaydigan, maxsus jihozlangan hududlardir. "
        "Har bir markaz muayyan ta'limiy maqsadga xizmat qiladi va bolaga "
        "mustaqil faoliyat yuritish imkonini beradi. Mavzuli rejalashtirish "
        "doirasida bu markazlar bitta mavzuga moslashtiriladi, ya'ni "
        "markazdagi barcha materiallar va topshiriqlar shu hafta mavzusiga "
        "oid bo'ladi.",

        "Nutq va savodxonlik markazi ta'limiy jarayonda muhim o'rin tutadi. "
        "Bu markazda bolaning og'zaki nutqi, lug'at boyligi, grammatik "
        "to'g'ri gapirish va savodxonlik ko'nikmalari rivojlantiriladi. "
        "Mavzuli rejalashtirish doirasida bu markazga mavzuga oid kitoblar, "
        "rasmli kartochkalar, so'z o'yinlari, audio-ertaklar joylashtiriladi. "
        "Masalan, \"Transport\" mavzusida transport vositalari rasmlari, "
        "ularning nomlari yozilgan kartochkalar va transport haqidagi "
        "she'rlar kiritiladi.",

        "Matematik va mantiq markazi bolaning matematik tafakkurini "
        "shakllantiradi. Mavzuli rejalashtirish asosida bu markazga har hafta "
        "mavzuga tegishli matematik topshiriqlar joylashtiriladi. \"Mevalar\" "
        "mavzusida mevalarni sanash, guruhlash, katta-kichigiga qarab "
        "tartiblash topshiriqlari beriladi. Bu bola uchun matematikani "
        "qiziqarli va hayotiy qiladi.",

        "Tabiat va tadqiqot markazi bolada atrof-muhitga qiziqish uyg'otadi. "
        "Mavzuli rejalashtirish asosida bu markaz har hafta yangi materiallar "
        "bilan boyitiladi. \"Suv\" mavzusida suvning xossalari bilan "
        "tajribalar, \"O'simliklar\" mavzusida urug'lar ekish kuzatuvi, "
        "\"Qishki tabiat\" mavzusida qor va muz bilan eksperimentlar "
        "o'tkaziladi. Bola kuzatish, taqqoslash va xulosa chiqarishni "
        "o'rganadi.",

        "Badiiy ijod markazi bolaning estetik tafakkuri va ijodiy "
        "qobiliyatlarini rivojlantiradi. Mavzuli rejalashtirish doirasida "
        "bu markazda mavzuga oid rasm chizish, applikatsiya, plastilindan "
        "shakllar yasash va boshqa ijodiy ishlar rejalashtiriladi. \"Kosmik "
        "olam\" mavzusida yulduzlar va sayyoralar rasmi chiziladi, "
        "\"Bayramlar\" mavzusida tabrknoma tayyorlanadi. Ijod markazi "
        "bolaga o'z fikrini badiiy vositalar orqali ifodalashga o'rgatadi.",

        "Rolli o'yinlar markazi bolaning ijtimoiy ko'nikmalarini "
        "shakllantiradi. Mavzuli rejalashtirish asosida bu markaz har hafta "
        "mavzuga mos o'yin stsenariysi bilan boyitiladi. \"Kasb-hunarlar\" "
        "mavzusida shifokor, oshpaz, sotuvchi, haydovchi o'yinlari tashkil "
        "etiladi. Bola kasblar haqida amaliy tushuncha hosil qiladi va "
        "ijtimoiy rollarni o'zlashtirishni o'rganadi.",
    ],

    # 5. Tashkil etish bosqichlari
    [
        "Mavzuli rejalashtirish asosida rivojlanish markazlarida ta'limiy "
        "faoliyatni tashkil etish bir necha bosqichda amalga oshiriladi. "
        "Birinchi bosqich — mavzuni tanlash va maqsadlarni belgilash. "
        "Tarbiyachi metodik qo'llanmaga asoslanib, hafta uchun mavzuni "
        "tanlaydi va har bir rivojlanish markazi uchun ta'limiy maqsadlarni "
        "aniqlab oladi. Maqsadlar bolaning yosh xususiyatlariga va "
        "rivojlanish darajasiga mos bo'lishi kerak.",

        "Ikkinchi bosqich — muhitni tayyorlash. Tarbiyachi tanlangan mavzuga "
        "muvofiq barcha rivojlanish markazlarini qayta jihozlaydi. Eski "
        "materiallar olib tashlanadi, mavzuga oid yangi o'yinchoqlar, "
        "kitoblar, didaktik materiallar, rasmlar va jihozlar joylashtiriladi. "
        "Bu bosqichda muhitning xavfsizligi, estetik ko'rinishi va "
        "bolalarning erkin foydalanishi uchun qulay joylashtirilishi "
        "ta'minlanadi.",

        "Uchinchi bosqich — mavzu bilan tanishtirish. Hafta boshida "
        "tarbiyachi ertalabki davra vaqtida bolalarni yangi mavzu bilan "
        "tanishtiradi. Bu suhbat, ko'rgazma, video yoki surpriz element "
        "orqali amalga oshirilishi mumkin. Maqsad — bolalarda mavzuga "
        "qiziqish uyg'otish va ularni markazlardagi faoliyatga undash. "
        "Tarbiyachi markazlarda qanday yangiliklar borligini tushuntiradi.",

        "To'rtinchi bosqich — mustaqil faoliyat jarayoni. Bolalar o'zlari "
        "tanlagan markazga o'tib, mavzuga oid topshiriqlarni bajaradilar. "
        "Tarbiyachi bu jarayonda kuzatuvchi va yo'naltiruvchi vazifasini "
        "bajaradi. U zarurat bo'lganda yordam beradi, savollarga javob "
        "beradi, lekin bolaning mustaqilligini cheklamaydi. Har bir bola "
        "o'z tezligida va qiziqishiga qarab ishlaydi.",

        "Beshinchi bosqich — yakunlash va refleksiya. Hafta oxirida "
        "yoki har kunning oxirida tarbiyachi bolalar bilan mavzu bo'yicha "
        "yakuniy suhbat o'tkazadi. Bolalar nima o'rganganlarini, nimalar "
        "qilganlarini gapirib beradilar. Bu bolalarda o'z faoliyatini "
        "tahlil qilish ko'nikmasini shakllantiradi. Shuningdek, tarbiyachi "
        "bolalarning ishlarini ko'rgazmaga qo'yadi, ota-onalarga namoyish "
        "etadi.",

        "Oltinchi bosqich — baholash va tahlil. Tarbiyachi hafta davomida "
        "to'plangan kuzatuv natijalarini tahlil qiladi, har bir bolaning "
        "yutuqlari va qiyinchiliklarini qayd etadi. Bu ma'lumotlar keyingi "
        "hafta rejasini tuzishda, individual ishlar rejalashtirishda va "
        "ota-onalar bilan suhbatlarda qo'llaniladi. Shunday qilib, ta'limiy "
        "jarayon uzluksiz takomillashib boradi.",
    ],

    # 6. Yosh guruhlari bo'yicha xususiyatlar
    [
        "Mavzuli rejalashtirishdan foydalanish turli yosh guruhlarida o'ziga "
        "xos xususiyatlarga ega. Kichik guruh (3-4 yosh) bolalari uchun "
        "mavzular sodda va bolaning kundalik hayotiga yaqin bo'lishi kerak. "
        "\"Mening o'yinchoqlarim\", \"Oilam\", \"Uy hayvonlari\" kabi "
        "mavzular tanlanadi. Markazlardagi topshiriqlar ham sodda, qisqa "
        "muddatli va ko'rgazmali bo'lishi shart. Bola bir markazda 10-15 "
        "daqiqa faoliyat yuritishi kutiladi.",

        "Kichik guruhda rivojlanish markazlarining soni 4-5 tadan oshmasligi "
        "maqsadga muvofiq. Chunki bu yoshdagi bola ko'p tanlov oldida "
        "o'zini yo'qotib qo'yishi mumkin. Markazlardagi materiallar yirik, "
        "yorqin va xavfsiz bo'lishi kerak. Tarbiyachi bolalarga doimiy "
        "yo'l-yo'riq beradi, namuna ko'rsatadi va birgalikda faoliyat "
        "olib boradi. Mustaqillik darajasi asta-sekin oshirib boriladi.",

        "O'rta guruh (4-5 yosh) bolalari uchun mavzular yanada kengroq va "
        "murakkab bo'ladi. \"Yil fasllari\", \"Kasb-hunarlar\", \"Sog'lom "
        "turmush tarzi\", \"Shahrimiz\" kabi mavzular tanlanadi. Markazlar "
        "soni 5-6 tagacha oshiriladi va topshiriqlar murakkablashadi. Bola "
        "bir markazda 15-20 daqiqa faoliyat yuritishi mumkin. Bu yoshda "
        "bolalar mavzuga oid savollar bera boshlaydi va izlanishga "
        "qiziqadi.",

        "Katta guruh (5-6 yosh) va maktabga tayyorlov guruhi (6-7 yosh) "
        "bolalari uchun mavzular chuqurroq va ko'p qirrali bo'ladi. "
        "\"Kosmik olam\", \"Dunyo mamlakatlari\", \"Tabiat muhofazasi\", "
        "\"Ixtirolar\" kabi mavzular tanlanadi. Markazlar soni 7-8 tagacha "
        "yetadi va har bir markazda murakkab, ko'p bosqichli topshiriqlar "
        "bo'ladi. Bola mustaqil ravishda mavzuni o'rganadi, loyihalar "
        "tayyorlaydi va natijalarini taqdimot qiladi.",

        "Katta yoshdagi bolalar uchun mavzuli rejalashtirishda loyiha "
        "metodi ham qo'llaniladi. Bunda butun hafta davomida bolalar "
        "birgalikda bitta katta loyiha ustida ishlaydi. Masalan, \"Bizning "
        "bog'chamiz\" mavzusida bolalar bog'cha maketini yasaydi, tabiat "
        "burchagidagi o'simliklarni kuzatadi, bog'cha haqida kitobcha "
        "tayyorlaydi. Bu jamoaviy ish ko'nikmalarini shakllantiradi va "
        "bolani maktabdagi loyiha faoliyatiga tayyorlaydi.",
    ],

    # 7. Ta'limiy faoliyat shakllari
    [
        "Mavzuli rejalashtirish asosida rivojlanish markazlarida tashkil "
        "etiladigan ta'limiy faoliyat turli shakllarda amalga oshiriladi. "
        "Birinchi shakl — mustaqil tadqiqot faoliyati. Bola markazdagi "
        "materiallar bilan mustaqil ishlaydi, kuzatadi, sinab ko'radi va "
        "xulosa chiqaradi. Masalan, eksperiment markazida bola mustaqil "
        "ravishda magnit tajribasini o'tkazib, qaysi buyumlar tortilishini "
        "aniqlaydi.",

        "Ikkinchi shakl — didaktik o'yinlar. Har bir markaz uchun mavzuga "
        "oid maxsus o'yinlar tayyorlanadi. \"Loto\", \"Domino\", \"Juftini "
        "top\", \"Ortiqchasini chiqar\" kabi klassik didaktik o'yinlar "
        "mavzuga moslashtiriladi. Masalan, \"Mevalar va sabzavotlar\" "
        "mavzusida meva-sabzavot lotosi, ularni rangi, shakli yoki "
        "o'sish joyiga qarab guruhlash o'yinlari tayyorlanadi.",

        "Uchinchi shakl — ijodiy faoliyat. Badiiy ijod markazida bola "
        "mavzuga oid rasm chizadi, applikatsiya qiladi, plastilindan "
        "shakl yasaydi yoki qurilish markazida maket yaratadi. Ijodiy "
        "faoliyat bolaning fantaziyasini, mayda motorikasini va estetik "
        "didini rivojlantiradi. Tarbiyachi namuna ko'rsatadi, lekin bola "
        "o'z g'oyasini amalga oshirishi uchun erkinlik beradi.",

        "To'rtinchi shakl — rolli va syujetli o'yinlar. Mavzuga mos "
        "syujetli o'yinlar tashkil etiladi. \"Kasb-hunarlar\" mavzusida "
        "\"Shifoxona\", \"Maktab\", \"Oshxona\" o'yinlari; \"Transport\" "
        "mavzusida \"Avtobus\", \"Poyezd\", \"Aeroport\" o'yinlari "
        "o'ynaladi. Bola kattalar hayotini qayta yaratib, ijtimoiy "
        "tajriba to'playdi va muloqot ko'nikmalarini rivojlantiradi.",

        "Beshinchi shakl — kitob bilan ishlash. Nutq markazida bola "
        "mavzuga oid kitoblar varaqlab, rasmlarini ko'rib, tarbiyachi "
        "o'qigan ertakni qayta hikoya qilib beradi. Katta yoshdagi "
        "bolalar mustaqil o'qish mashqlarini bajaradi. Oltinchi shakl — "
        "musiqa va harakatli faoliyat. Musiqa markazida mavzuga oid "
        "qo'shiqlar tinglanadi, ritmik harakatlar bajariladi, teatr "
        "ko'rinishlari namoyish etiladi.",

        "Yettinchi shakl — kuzatish va tajriba. Tabiat va eksperiment "
        "markazlarida bola mavzuga oid hodisalarni kuzatadi, sodda "
        "tajribalar o'tkazadi. \"Suv\" mavzusida suvning oqishi, "
        "ranglanishi, muzlashi kuzatiladi. \"Havo\" mavzusida pufaklarga "
        "havo to'ldiriladi, shamol tajribasi o'tkaziladi. Bu faoliyat "
        "bolaning ilmiy tafakkur asoslarini shakllantiradi va sabab-natija "
        "aloqalarini tushunishga yordam beradi.",
    ],

    # 8. Tarbiyachining roli
    [
        "Mavzuli rejalashtirish asosida rivojlanish markazlarida ta'limiy "
        "faoliyatni tashkil etishda tarbiyachining roli nihoyatda muhim. "
        "Avvalo, tarbiyachi rejalashtiruvchi sifatida faoliyat yuritadi. "
        "U metodik qo'llanma asosida haftalik va kunlik rejalarni tuzadi, "
        "har bir markaz uchun maqsad va topshiriqlar belgilaydi, zarur "
        "materiallarni tayyorlaydi. Rejalashtirish jarayonida u bolalarning "
        "yosh xususiyatlari, qiziqishlari va rivojlanish darajasini "
        "hisobga oladi.",

        "Ikkinchi rol — muhit tashkilotchisi. Tarbiyachi har hafta "
        "markazlarni mavzuga moslashtirib qayta jihozlaydi. U materiallarni "
        "bolalar uchun qulay va jozibali tarzda joylashtiradi, estetik "
        "ko'rinishni ta'minlaydi, xavfsizlik qoidalariga rioya qiladi. "
        "Muhitning har bir elementi bolani faoliyatga undashi kerak. "
        "Tarbiyachining mahorati shundaki, u markazlarni shunday "
        "tayyorlaydiki, bola o'z-o'zidan ularni kashf etishni xohlaydi.",

        "Uchinchi rol — kuzatuvchi va yo'naltiruvchi. Ta'limiy faoliyat "
        "jarayonida tarbiyachi bolalarni kuzatadi, ularning qanday "
        "ishlayotganini, qayerda qiyinchilik chekayotganini aniqlaydi. "
        "Zarurat bo'lganda u yengil ko'rsatma beradi, savol orqali bolani "
        "to'g'ri yo'lga yo'naltiradi, lekin javobni tayyor holda bermaydi. "
        "Bu yondashuv bolaning mustaqil fikrlashini saqlab qoladi.",

        "To'rtinchi rol — baholovchi va tahlilchi. Tarbiyachi har bir "
        "bolaning markazlardagi faoliyatini muntazam kuzatib, natijalarni "
        "qayd etadi. U maxsus kuzatuv kartochkalari, diagnostika "
        "jadvallari va portfolio yaratadi. Bu ma'lumotlar asosida "
        "individual ta'limiy marshrut tuziladi, ota-onalarga bolaning "
        "rivojlanishi haqida ma'lumot beriladi va keyingi hafta rejasi "
        "takomillashtiriladi.",

        "Beshinchi rol — hamkor va motivator. Tarbiyachi bola bilan "
        "teng huquqli muloqot olib boradi, uning g'oyalarini qadrlaydi, "
        "muvaffaqiyatlarini nishonlaydi va qo'llab-quvvatlaydi. U bolada "
        "markazlarga borish istagini uyg'otadi, yangi topshiriqlar bilan "
        "qiziqtiradi va ijobiy muhit yaratadi. Shuningdek, tarbiyachi "
        "ota-onalar bilan hamkorlik qiladi — mavzu haqida xabar beradi, "
        "uyda qanday mashqlar bajarish mumkinligini tushuntiradi.",

        "Oltinchi rol — o'z-o'zini rivojlantiruvchi mutaxassis. "
        "Tarbiyachi mavzuli rejalashtirish texnologiyasini mukammal "
        "egallashi, yangi pedagogik yondashuvlarni o'rganishi va "
        "amaliyotiga tatbiq etishi kerak. U metodik birlashmalarda "
        "tajriba almashadi, seminarlar va treninglarda qatnashadi, "
        "ilg'or tajribalarni o'z guruhiga moslaydi. Tarbiyachining "
        "kasbiy o'sishi bevosita bolalar rivojlanishiga ijobiy ta'sir "
        "ko'rsatadi.",
    ],

    # 9. Xulosa
    [
        "Yuqoridagi tahlillar shuni ko'rsatadiki, mavzuli rejalashtirish "
        "metodik qo'llanmasidan foydalangan holda rivojlanish markazlarida "
        "ta'limiy faoliyatni tashkil etish — bu zamonaviy maktabgacha "
        "ta'limning eng samarali usullaridan biri. Bu yondashuv bolaning "
        "bilim olishini izchil, qiziqarli va hayotiy tajribaga "
        "yaqinlashtiradi.",

        "Mavzuli rejalashtirish tarbiyachiga ta'limiy jarayonni puxta "
        "rejalashtirib, zarur materiallarni oldindan tayyorlab qo'yish "
        "imkonini beradi. Rivojlanish markazlari esa bolaga mustaqil "
        "faoliyat yuritish, o'z qiziqishlariga qarab ish tanlash va "
        "shaxsiy tezlikda ishlash imkonini yaratadi. Bu ikki yondashuv "
        "birgalikda bolaning har tomonlama rivojlanishini ta'minlaydi.",

        "Mavzuli rejalashtirish samaradorligi ko'p jihatdan tarbiyachining "
        "mahorati, metodik tayyorgarligi va ijodiy yondashuviga bog'liq. "
        "Tarbiyachi mavzuni qanchalik qiziqarli tarzda ochib bersa, "
        "markazlarni qanchalik jozibali jihozlasa, bolaning ta'limiy "
        "faoliyatga qiziqishi shunchalik yuqori bo'ladi. Shu sababli "
        "tarbiyachining uzluksiz kasbiy rivojlanishi muhim ahamiyatga ega.",

        "Xulosa qilib aytganda, mavzuli rejalashtirish metodik qo'llanmasi "
        "va rivojlanish markazlari — bu maktabgacha ta'limning zamonaviy "
        "pedagogik texnologiyalari bo'lib, ular bolaning mustaqillik, "
        "ijodkorlik, izlanuvchanlik va ijtimoiy ko'nikmalarini samarali "
        "rivojlantiradi. Bu texnologiyalarni amaliyotga keng joriy etish "
        "maktabgacha ta'lim sifatini yanada oshirishga xizmat qiladi.",

        "Foydalanilgan adabiyotlar:",
        "1. O'zbekiston Respublikasining \"Maktabgacha ta'lim va tarbiya "
        "to'g'risida\"gi Qonuni. — Toshkent, 2020.",
        "2. Maktabgacha ta'limning Davlat o'quv dasturi \"Ilk qadam\". — "
        "Toshkent, 2018.",
        "3. Xasanova M., Raximova D. \"Maktabgacha ta'limda mavzuli "
        "rejalashtirish\". — Toshkent, 2021.",
        "4. Jumayeva G. \"Rivojlanish markazlarini tashkil etish "
        "metodikasi\". — Toshkent, 2022.",
        "5. Karimova N. \"Bolalar bog'chasida ta'limiy muhitni "
        "loyihalash\". — Toshkent, 2020.",
        "6. Internet manbalar: uzedu.uz, mtt.uz.",
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
    "Mavzuli rejalashtirish metodik qo'llanmasidan foydalangan holda "
    "rivojlanish markazlarida ta'limiy faoliyatni tashkil etish",
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

# CONTENT PAGES (9 sections)
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

# Supporting XML files
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
<dc:title>Mustaqil ish - Mavzuli rejalashtirish</dc:title>
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

import os
import asyncio
from playwright.async_api import async_playwright
from pptx import Presentation
from pptx.util import Inches

SLIDES_DATA = [
    {
        "bg": "linear-gradient(135deg, #0b192c 0%, #1e3e62 100%)",
        "badge": "O'ZBEKISTONNING ENG YANGI TARIXI",
        "title": "Yangi O'zbekiston Taraqqiyot Strategiyasi va Inson Qadri Tamoyili",
        "subtitle": "Inson qadrini ulug'lash, erkin fuqarolik jamiyatini barpo etish va milliy iqtisodiyotni yuksaltirishga qaratilgan strategik islohotlarning atroflicha ilmiy-vizual tahlili",
        "type": "hero"
    },
    {
        "bg": "linear-gradient(135deg, #0b192c 0%, #112136 100%)",
        "badge": "MUNDARIJA VA REJA",
        "title": "Taqdimot Rejasi va Asosiy Bo'limlar",
        "type": "agenda",
        "items": [
            ("01", "Yangi O'zbekiston Tushunchasi va Strategik Maqsadlar", "Milliy tiklanishdan milliy yuksalish sari o'tish bosqichlari", "https://api.iconify.design/lucide:target.svg?color=%2338bdf8"),
            ("02", "Inson Qadri va Uning Davlat Siyosatidagi Ustuvor O'rni", "Inson - jamiyat - davlat tamoyili va ijtimoiy kafolatlar", "https://api.iconify.design/lucide:user-check.svg?color=%2338bdf8"),
            ("03", "Iqtisodiy Islohotlar: Erkin Bozor va Investitsion Muhit", "Tadbirkorlikni qo'llab-quvvatlash va raqamli iqtisodiyot", "https://api.iconify.design/lucide:trending-up.svg?color=%2338bdf8"),
            ("04", "Ta'lim va Sog'liqni Saqlash Sohasidagi Tub O'zgarishlar", "Inson kapitaliga investitsiya va tibbiy xizmatlar sifati", "https://api.iconify.design/lucide:graduation-cap.svg?color=%2338bdf8"),
            ("05", "Tashqi Siyosat va Markaziy Osiyodagi Yangi Muhit", "Ochiq, pragmatik va yaxshi qo'shnichilik diplomatiyasi", "https://api.iconify.design/lucide:globe.svg?color=%2338bdf8")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%)",
        "badge": "KIRISH VA KONSEPSIYA",
        "title": "Yangi Davrning Tarixiy Zaruriyati va Bosh G'oyasi",
        "type": "cards",
        "cards": [
            ("Tarixiy Burilish Bosqichi", "2016-yildan e'tiboran O'zbekiston rivojlanishining mutlaqo yangi davri boshlandi.", "Davlat va jamiyat o'rtasidagi munosabatlar tubdan o'zgardi va ochiqlik ta'minlandi.", "https://api.iconify.design/lucide:history.svg?color=%2338bdf8", "https://images.unsplash.com/photo-1541872703-74c5e44368f9?w=600&auto=format&fit=crop&q=80"),
            ("Tarixiy Konsepsiya", "'Milliy tiklanishdan - Milliy yuksalish sari' g'oyasi bosh tamoyilga aylandi.", "Islohotlar har bir fuqaro hayotida sezilishi shart qilib belgilandi.", "https://api.iconify.design/lucide:compass.svg?color=%2338bdf8", "https://images.unsplash.com/photo-1517048676732-d65bc937f952?w=600&auto=format&fit=crop&q=80"),
            ("Bosh Maqsad va Vazifa", "Erkin, obod va farovon Yangi O'zbekistonni barpo etish.", "Kelajak avlod uchun barqaror va qudratli rivojlanish poydevorini yaratish.", "https://api.iconify.design/lucide:flag.svg?color=%2338bdf8", "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #064e3b 100%)",
        "badge": "REJA 1: STRATEGIK MAQSADLAR",
        "title": "1. Yangi O'zbekiston Taraqqiyot Strategiyasi (2022-2026)",
        "type": "chart_slide",
        "chart_title": "Taraqqiyot Strategiyasi Ustuvor Yo'nalishlari Salmog'i",
        "labels": ["Inson Qadri va Ijtimoiy Himoya", "Erkin Bozor va Iqtisodiyot", "Adolat va Qonun Ustuvorligi", "Tashqi Siyosat va Xavfsizlik", "Ma'naviyat va Ta'lim"],
        "values": [30, 25, 20, 15, 10],
        "colors": ["#38bdf8", "#34d399", "#fbbf24", "#a78bfa", "#f472b6"],
        "cards": [
            ("7 Ta Ustuvor Yo'nalish", "Erkin fuqarolik jamiyati hamda adolat va qonun ustuvorligini ta'minlash.", "Milliy iqtisodiyotni jadal rivojlantirish va adolatli ijtimoiy siyosat.", "https://api.iconify.design/lucide:layers.svg?color=%2334d399"),
            ("Kutilayotgan Natijalar", "Aholi jon boshiga YIM hajmini oshirish va o'rta daromadli mamlakatlar safga kirish.", "Kambag'allik darajasini kamida 2 baravarga qisqartirish.", "https://api.iconify.design/lucide:award.svg?color=%2334d399")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #312e81 100%)",
        "badge": "REJA 1: DAVLAT BOSHQARUVI",
        "title": "Davlat Boshqaruvi va Mahalla Tizimida Transformatsiya",
        "type": "cards",
        "cards": [
            ("Tizimli Transformatsiya", "Ixcham va samarali davlat boshqaruv apparatini shakllantirish.", "Vazirliklar mas'uliyatini oshirish va xizmatlarni to'liq raqamlashtirish.", "https://api.iconify.design/lucide:cpu.svg?color=%23818cf8", "https://images.unsplash.com/photo-1451187580459-43490279c0fa?w=600&auto=format&fit=crop&q=80"),
            ("Mahallabay Tizimi", "Mahalla - jamiyatning tayanch bo'g'iniga aylantirildi.", "'Hokim yordamchisi', 'Yoshlar yetakchisi' va 'Xotin-qizlar faoli' institutlari.", "https://api.iconify.design/lucide:home.svg?color=%23818cf8", "https://images.unsplash.com/photo-1582213782179-e0d53f98f2ca?w=600&auto=format&fit=crop&q=80"),
            ("Xalq Bilan Muloqot", "Prezident Xalq qabulxonalari va my.gov.uz elektron xizmatlari samaradorligi.", "Jamoatchilik nazorati va fuqarolar tashabbusining oshishi.", "https://api.iconify.design/lucide:message-square.svg?color=%23818cf8", "https://images.unsplash.com/photo-1577563908411-5077b6dc7624?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #451a03 100%)",
        "badge": "REJA 2: INSON QADRI",
        "title": "2. 'Inson Qadri Uchun' Tamoyilining Mazmun-Mohiyati",
        "type": "cards",
        "cards": [
            ("Inson Qadri - Bosh Mezon", "'Inson - jamiyat - davlat' tamoyili ustuvor etib belgilandi.", "Davlat idoralari xalqqa xizmat qilishi konstitutsiyaviy majburiyat qilindi.", "https://api.iconify.design/lucide:heart.svg?color=%23fbbf24", "https://images.unsplash.com/photo-1531206715517-5c0ba140b2b8?w=600&auto=format&fit=crop&q=80"),
            ("Ijtimoiy Adolat va Himoya", "Kam ta'minlangan oilalarni manzilli qo'llab-quvvatlash.", "'Ayollar daftari', 'Yoshlar daftari' va 'Saxovat daftari' tizimlari.", "https://api.iconify.design/lucide:shield-check.svg?color=%23fbbf24", "https://images.unsplash.com/photo-1469571486292-0ba58a3f068b?w=600&auto=format&fit=crop&q=80"),
            ("Konstitutsiyaviy Islohotlar", "Yangi tahrirdagi Konstitutsiyada 'Ijtimoiy Davlat' maqomining belgilanishi.", "Insonning mehnati va ta'lim olish huquqlarining kengaytirilishi.", "https://api.iconify.design/lucide:book-open.svg?color=%23fbbf24", "https://images.unsplash.com/photo-1455390582262-044cdead277a?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #1e293b 100%)",
        "badge": "REJA 2: HUQUQIY KAFOLATLAR",
        "title": "Inson Huquqlari, Sud Mustaqilligi va So'z Erkinligi",
        "type": "cards",
        "cards": [
            ("Sud Mustaqilligi", "Sudlarning haqiqiy mustaqilligini ta'minlash va inson huquqlarini himoya qilish.", "Advokatura institutini va himoya huquqini tubdan kuchaytirish.", "https://api.iconify.design/lucide:scale.svg?color=%2338bdf8", "https://images.unsplash.com/photo-1589829545856-d10d557cf95f?w=600&auto=format&fit=crop&q=80"),
            ("Majburiy Mehnatga Barham", "Bolalar mehnati va majburiy mehnat to'liq tugatildi va XMT tomonidan e'tirof etildi.", "Paxta boykoti bekor qilinishiga erishildi.", "https://api.iconify.design/lucide:check-circle.svg?color=%2338bdf8", "https://images.unsplash.com/photo-1509099836639-18ba1795216d?w=600&auto=format&fit=crop&q=80"),
            ("So'z Erkinligi va OAV", "Ommaviy axborot vositalari va blogerlar faoliyatiga keng imkoniyatlar.", "Davlat organlari faoliyatining ochiqligi va shaffofligi.", "https://api.iconify.design/lucide:mic.svg?color=%2338bdf8", "https://images.unsplash.com/photo-1504711434969-e33886168f5c?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #065f46 100%)",
        "badge": "REJA 3: IQTISODIYOT",
        "title": "3. Iqtisodiy Islohotlar: Valyuta, Soliq va Erkin Bozor",
        "type": "chart_slide",
        "chart_title": "Iqtisodiy O'sish va QQS Stavkasi Pasayishi (%)",
        "labels": ["Eski QQS Stavkasi", "Yangi QQS Stavkasi", "Valyuta Erkinligi", "Biznes Imtiyozlari"],
        "values": [15, 12, 95, 88],
        "colors": ["#ef4444", "#34d399", "#38bdf8", "#fbbf24"],
        "cards": [
            ("Valyuta Bozorini Erkinlashtirish", "2017-yildagi erkin valyuta konvertatsiyasining joriy etilishi.", "Xorijiy investorlar va mahalliy biznes uchun teng sharoitlar yaratish.", "https://api.iconify.design/lucide:dollar-sign.svg?color=%2334d399"),
            ("Soliq Yukini Kamaytirish", "Soliq stavkalarining maqbullashtirilishi va soddalashtirilishi.", "QQS stavkasining 15% dan 12% ga tushirilishi hamda yengilliklar.", "https://api.iconify.design/lucide:percent.svg?color=%2334d399")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%)",
        "badge": "REJA 3: INVESTITSIYA VA BIZNES",
        "title": "Investitsiyalar Jalb Etish va Tadbirkorlikni Qo'llab-quvvatlash",
        "type": "cards",
        "cards": [
            ("Investitsiya Muhiti", "Xorijiy sarmoyadorlar uchun huquqiy va iqtisodiy kafolatlar.", "Maxsus iqtisodiy zonalar (MIZ) va texnoparklar tarmog'ini kengaytirish.", "https://api.iconify.design/lucide:briefcase.svg?color=%23a78bfa", "https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=600&auto=format&fit=crop&q=80"),
            ("Kichik va O'rta Biznes", "Tadbirkorlik faoliyatini tekshirishlarga moratoriy va yengilliklar.", "Yangi ish o'rinlarini yaratishda biznesning drayverga aylanishi.", "https://api.iconify.design/lucide:pie-chart.svg?color=%23a78bfa", "https://images.unsplash.com/photo-1556761175-5973dc0f32e7?w=600&auto=format&fit=crop&q=80"),
            ("Sanoat va Eksport", "Xomashyo eksportidan tayyor mahsulot eksportiga o'tish strategiyasi.", "Jahon Savdo Tashkilotiga (JST) a'zo bo'lish bo'yicha jadal muzokaralar.", "https://api.iconify.design/lucide:globe-2.svg?color=%23a78bfa", "https://images.unsplash.com/photo-1586528116311-ad8dd3c8310d?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #3730a3 100%)",
        "badge": "REJA 4: TA'LIM ISLOHOTLARI",
        "title": "4. Ta'lim Sohasidagi Tub O'zgarishlar va Inson Kapitali",
        "type": "chart_slide",
        "chart_title": "Oliy va Maktabgacha Ta'lim Qamrovi O'sishi (%)",
        "labels": ["Bog'cha Qamrovi (Eski)", "Bog'cha Qamrovi (Hozir)", "OTM Qamrovi (Eski)", "OTM Qamrovi (Hozir)"],
        "values": [27, 72, 9, 38],
        "colors": ["#94a3b8", "#38bdf8", "#94a3b8", "#a78bfa"],
        "cards": [
            ("Maktablar va O'qituvchi Qadri", "O'qituvchilar ish haqini oshirish va ularni majburiy mehnatsiz qilish.", "Prezident maktablari va Ijod maktablari tarmog'ini yaratish.", "https://api.iconify.design/lucide:book-open.svg?color=%23a78bfa"),
            ("Oliy Ta'lim Qamrovi", "Oliy ta'lim bilan qamrovni 9% dan 38% ga oshirish va filiallar ochish.", "OTMlarga akademik va moliyaviy mustaqillik berilishi.", "https://api.iconify.design/lucide:graduation-cap.svg?color=%23a78bfa")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #15803d 100%)",
        "badge": "REJA 4: SOG'LIQNI SAQLASH",
        "title": "Sog'liqni Saqlash Tizimi va Aholi Salomatligi Muhofazasi",
        "type": "cards",
        "cards": [
            ("Tibbiy Xizmat Sifati", "Birinchi bo'g'in (oilaviy shifokorlik)ni rivojlantirish.", "Qishloq joylarda tibbiy maskanlar va tez yordam sifatini oshirish.", "https://api.iconify.design/lucide:activity.svg?color=%234ade80", "https://images.unsplash.com/photo-1538108149393-fbbd81895907?w=600&auto=format&fit=crop&q=80"),
            ("Sug'urta va Texnologiya", "Davlat tibbiy sug'urtasi tizimini bosqichma-bosqich joriy etish.", "Yuqori texnologik murakkab operatsiyalarni mahalliy darajada o'tkazish.", "https://api.iconify.design/lucide:stethoscope.svg?color=%234ade80", "https://images.unsplash.com/photo-1516549655169-df83a0774514?w=600&auto=format&fit=crop&q=80"),
            ("Aholi Salomatligi", "Ommaviy sport va sog'lom turmush tarzini targ'ib qilish.", "Aholining o'rtacha umr ko'rish davomiyligining sezilarli uzayishi.", "https://api.iconify.design/lucide:smile.svg?color=%234ade80", "https://images.unsplash.com/photo-1571019613454-1cb2f99b2d8b?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%)",
        "badge": "REJA 5: TASHQI SIYOSAT",
        "title": "5. Tashqi Siyosat: Markaziy Osiyodagi Yangi Muhit",
        "type": "cards",
        "cards": [
            ("Markaziy Osiyo Ustuvorligi", "Qo'shni davlatlar bilan do'stona aloqalarni tiklash.", "Chegara va suv-energetika muammolarini tinch muloqot orqali hal etish.", "https://api.iconify.design/lucide:map-pin.svg?color=%2360a5fa", "https://images.unsplash.com/photo-1526778548025-fa2f459cd5c1?w=600&auto=format&fit=crop&q=80"),
            ("Mintaqaviy Integratsiya", "Markaziy Osiyo davlatlari rahbarlarining Maslahat uchrashuvlari.", "Mintaqada o'zaro ishonch va barqarorlik muhitining qaror topishi.", "https://api.iconify.design/lucide:users.svg?color=%2360a5fa", "https://images.unsplash.com/photo-1511632765486-a01980e01a18?w=600&auto=format&fit=crop&q=80"),
            ("Xalqaro Tashkilotlar", "BMT, ShHT, TDT va MDB doirasida faol va tashabbuskor diplomatiya.", "O'zbekistonning xalqaro sammitlar va muloqotlar markaziga aylanishi.", "https://api.iconify.design/lucide:globe.svg?color=%2360a5fa", "https://images.unsplash.com/photo-1541872703-74c5e44368f9?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #0f766e 100%)",
        "badge": "REJA 5: GLOBAL TASHABBUSLAR",
        "title": "Global Tashabbuslar, Logistika va Ekologik Diplomatiya",
        "type": "cards",
        "cards": [
            ("Transport Koridorlari", "Xitoy - Qirg'iziston - O'zbekiston temir yo'li loyihasi.", "Trans-Afg'on koridori va Janubiy Osiyoga chiqish strategiyasi.", "https://api.iconify.design/lucide:truck.svg?color=%232dd4bf", "https://images.unsplash.com/photo-1519003722824-194d4455a60c?w=600&auto=format&fit=crop&q=80"),
            ("Ekologik Diplomatiya", "Orolbo'yi mintaqasi bo'yicha global BMT rezolutsiyalari.", "'Yashil makon' umummilliy loyihasi va ekologik barqarorlik.", "https://api.iconify.design/lucide:leaf.svg?color=%232dd4bf", "https://images.unsplash.com/photo-1464822759023-fed622ff2c3b?w=600&auto=format&fit=crop&q=80"),
            ("Muvozanatli Diplomatiya", "AQSH, Yevropa Ittifoqi, Xitoy va Rossiya bilan muvozanatli aloqalar.", "O'zbekistonning iqtisodiy diplomatiyasi va eksport geografiyasini oshirish.", "https://api.iconify.design/lucide:compass.svg?color=%232dd4bf", "https://images.unsplash.com/photo-1507679799987-c73779587ccf?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #701a75 100%)",
        "badge": "XALQARO E'TIROF",
        "title": "Yangi O'zbekiston Islohotlarining Xalqaro E'tirofi",
        "type": "cards",
        "cards": [
            ("BMT va Tashkilotlar", "O'zbekiston ilgari surgan rezolutsiyalarning BMT tomonidan qabul qilinishi.", "Inson huquqlari bo'yicha Kengashga a'zolik va faoliyat.", "https://api.iconify.design/lucide:award.svg?color=%23f0abfc", "https://images.unsplash.com/photo-1451187580459-43490279c0fa?w=600&auto=format&fit=crop&q=80"),
            ("Nufuzli Reytinglar", "'The Economist' tomonidan O'zbekistonning 'Yil mamlakati' deb e'tirof etilishi.", "Jahon bankining 'Doing Business' va BMT indekslaridagi ko'tarilish.", "https://api.iconify.design/lucide:star.svg?color=%23f0abfc", "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=600&auto=format&fit=crop&q=80"),
            ("Mintaqaviy Liderlik", "Markaziy Osiyoda tinchlik va hamkorlik kafolatchisi sifatida e'tirof etilishi.", "Mintaqada barqaror rivojlanish drayveriga aylanishi.", "https://api.iconify.design/lucide:zap.svg?color=%23f0abfc", "https://images.unsplash.com/photo-1517245386807-bb43f82c33c4?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%)",
        "badge": "XULOSA VA ISTIQBOL",
        "title": "Xulosa: Yangi O'zbekistonning Tarixiy Roli va Kelajagi",
        "type": "cards",
        "cards": [
            ("Taraqqiyotning Ortga Qaytmasligi", "Yangi O'zbekiston strategiyasi - mamlakatimiz taraqqiyotining yangi poydevorini yaratdi.", "Islohotlar ortga qaytmas tus oldi va xalq tomonidan qo'llab-quvvatlanmoqda.", "https://api.iconify.design/lucide:shield.svg?color=%23a78bfa", "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=600&auto=format&fit=crop&q=80"),
            ("Inson Qadri - Doimiy Maydon", "Inson qadrini ulug'lash - amaliy harakatlar mezoniga aylandi.", "Har bir fuqaro uchun teng imkoniyatlar va adolatli jamiyat barpo etilmoqda.", "https://api.iconify.design/lucide:user-check.svg?color=%23a78bfa", "https://images.unsplash.com/photo-1531206715517-5c0ba140b2b8?w=600&auto=format&fit=crop&q=80"),
            ("Kelajakka Nigoh", "O'zbekiston-2030 strategiyasi doirasida yanada yuksak marralar belgilandi.", "Raqamli iqtisodiyot va yashil energetika sari sabit qadam tashlanmoqda.", "https://api.iconify.design/lucide:arrow-up-right.svg?color=%23a78bfa", "https://images.unsplash.com/photo-1451187580459-43490279c0fa?w=600&auto=format&fit=crop&q=80")
        ]
    },
    {
        "bg": "linear-gradient(135deg, #0f172a 0%, #030712 100%)",
        "badge": "E'TIBORINGIZ UCHUN RAHMAT",
        "title": "E'TIBORINGIZ UCHUN RAHMAT!",
        "subtitle": "Yangi O'zbekiston Taraqqiyot Strategiyasi va Inson Qadri Tamoyili  |  Savollar va muhokama uchun tashakkur!",
        "type": "hero"
    }
]

def generate_html(data, slide_num):
    bg = data.get("bg", "linear-gradient(135deg, #0f172a, #1e293b)")
    badge = data.get("badge", "O'ZBEKISTON TARIXI")
    title = data.get("title", "")
    subtitle = data.get("subtitle", "")
    stype = data.get("type", "cards")

    # Emblem Logo SVG for Top Right Header
    emblem_logo = """
    <div style="position: absolute; top: 60px; right: 90px; display: flex; align-items: center; gap: 14px; background: rgba(255,255,255,0.05); padding: 10px 20px; border-radius: 40px; border: 1px solid rgba(255,255,255,0.1);">
        <img src="https://upload.wikimedia.org/wikipedia/commons/7/77/Emblem_of_Uzbekistan.svg" style="width: 32px; height: 32px; filter: drop-shadow(0 2px 4px rgba(0,0,0,0.4));">
        <span style="font-size: 14px; font-weight: 700; color: #f8fafc; letter-spacing: 1px;">YANGI O'ZBEKISTON</span>
    </div>
    """

    content_html = ""
    if stype == "hero":
        content_html = f"""
        <div style="margin-top: 60px;">
            <div style="font-size: 58px; font-weight: 800; line-height: 1.25; background: linear-gradient(90deg, #ffffff, #93c5fd); -webkit-background-clip: text; -webkit-text-fill-color: transparent; max-width: 1400px; margin-bottom: 30px;">{title}</div>
            <div style="font-size: 24px; color: #94a3b8; max-width: 1200px; line-height: 1.6;">{subtitle}</div>
        </div>
        """
    elif stype == "agenda":
        items_html = ""
        for num, t, d, icon_url in data.get("items", []):
            items_html += f"""
            <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 20px; padding: 22px 30px; display: flex; align-items: center; gap: 24px;">
                <div style="background: linear-gradient(135deg, #38bdf8, #818cf8); color: #0f172a; font-size: 22px; font-weight: 800; width: 54px; height: 54px; border-radius: 16px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">{num}</div>
                <img src="{icon_url}" style="width: 32px; height: 32px; flex-shrink: 0;">
                <div>
                    <div style="font-size: 22px; font-weight: 700; color: #f8fafc; margin-bottom: 4px;">{t}</div>
                    <div style="font-size: 16px; color: #94a3b8;">{d}</div>
                </div>
            </div>
            """
        content_html = f"""
        <div>
            <div style="font-size: 40px; font-weight: 800; color: #ffffff; margin-top: 15px; margin-bottom: 25px;">{title}</div>
            <div style="display: flex; flex-direction: column; gap: 14px;">
                {items_html}
            </div>
        </div>
        """
    elif stype == "chart_slide":
        chart_title = data.get("chart_title", "")
        labels = data.get("labels", [])
        values = data.get("values", [])
        colors = data.get("colors", [])

        bars_html = ""
        for lbl, val, col in zip(labels, values, colors):
            bars_html += f"""
            <div style="margin-bottom: 20px;">
                <div style="display: flex; justify-content: space-between; font-size: 18px; font-weight: 600; color: #f8fafc; margin-bottom: 8px;">
                    <span>{lbl}</span>
                    <span style="color: {col}; font-weight: 800;">{val}%</span>
                </div>
                <div style="background: rgba(255,255,255,0.08); height: 16px; border-radius: 10px; overflow: hidden;">
                    <div style="background: {col}; width: {val}%; height: 100%; border-radius: 10px;"></div>
                </div>
            </div>
            """

        cards_html = ""
        for ctitle, p1, p2, icon_url in data.get("cards", []):
            cards_html += f"""
            <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 20px; padding: 28px; display: flex; flex-direction: column; gap: 14px;">
                <div style="display: flex; align-items: center; gap: 12px;">
                    <img src="{icon_url}" style="width: 28px; height: 28px;">
                    <div style="font-size: 22px; font-weight: 700; color: #f8fafc;">{ctitle}</div>
                </div>
                <div style="font-size: 15px; color: #cbd5e1; line-height: 1.5;">✔ {p1}</div>
                <div style="font-size: 15px; color: #cbd5e1; line-height: 1.5;">✔ {p2}</div>
            </div>
            """

        content_html = f"""
        <div>
            <div style="font-size: 38px; font-weight: 800; color: #ffffff; margin-top: 15px; margin-bottom: 30px;">{title}</div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 40px; align-items: start;">
                <!-- Left: Visual Chart Block -->
                <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 24px; padding: 36px;">
                    <div style="font-size: 22px; font-weight: 700; color: #38bdf8; margin-bottom: 24px;">📊 {chart_title}</div>
                    {bars_html}
                </div>
                <!-- Right: Content Cards -->
                <div style="display: flex; flex-direction: column; gap: 20px;">
                    {cards_html}
                </div>
            </div>
        </div>
        """
    else:
        cards_html = ""
        for item in data.get("cards", []):
            ctitle = item[0]
            p1 = item[1]
            p2 = item[2]
            icon_url = item[3] if len(item) > 3 else "https://api.iconify.design/lucide:star.svg?color=%2338bdf8"
            img_url = item[4] if len(item) > 4 else None

            img_block = f'<div style="height: 140px; border-radius: 14px; overflow: hidden; margin-bottom: 16px;"><img src="{img_url}" style="width: 100%; height: 100%; object-fit: cover;"></div>' if img_url else ''

            cards_html += f"""
            <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); backdrop-filter: blur(16px); border-radius: 24px; padding: 30px; display: flex; flex-direction: column; justify-content: space-between; box-shadow: 0 20px 40px rgba(0,0,0,0.3);">
                <div>
                    {img_block}
                    <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 12px;">
                        <img src="{icon_url}" style="width: 28px; height: 28px;">
                        <div style="font-size: 22px; font-weight: 700; color: #f8fafc; line-height: 1.3;">{ctitle}</div>
                    </div>
                </div>
                <div style="display: flex; flex-direction: column; gap: 12px;">
                    <div style="background: rgba(255,255,255,0.04); border-left: 4px solid #38bdf8; padding: 12px 16px; border-radius: 8px; font-size: 15px; color: #cbd5e1; line-height: 1.5;">{p1}</div>
                    <div style="background: rgba(255,255,255,0.04); border-left: 4px solid #818cf8; padding: 12px 16px; border-radius: 8px; font-size: 15px; color: #cbd5e1; line-height: 1.5;">{p2}</div>
                </div>
            </div>
            """
        content_html = f"""
        <div>
            <div style="font-size: 38px; font-weight: 800; color: #ffffff; margin-top: 15px; margin-bottom: 30px;">{title}</div>
            <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 24px; height: 500px;">
                {cards_html}
            </div>
        </div>
        """

    return f"""<!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', system-ui, -apple-system, sans-serif; }}
            body {{
                width: 1920px;
                height: 1080px;
                background: {bg};
                color: white;
                padding: 65px 85px;
                display: flex;
                flex-direction: column;
                justify-content: space-between;
                position: relative;
                overflow: hidden;
            }}
            .top-badge {{
                display: inline-flex;
                align-items: center;
                gap: 8px;
                background: rgba(56, 189, 248, 0.12);
                border: 1px solid rgba(56, 189, 248, 0.4);
                color: #38bdf8;
                padding: 10px 22px;
                border-radius: 30px;
                font-size: 15px;
                font-weight: 700;
                letter-spacing: 2px;
                width: fit-content;
            }}
            .footer-bar {{
                display: flex;
                align-items: center;
                justify-content: space-between;
                border-top: 1px solid rgba(255,255,255,0.1);
                padding-top: 18px;
                font-size: 16px;
                color: #64748b;
            }}
        </style>
    </head>
    <body>
        {emblem_logo}
        <div>
            <div class="top-badge">✦  {badge}</div>
            {content_html}
        </div>
        <div class="footer-bar">
            <div>Yangi O'zbekiston Taraqqiyot Strategiyasi va Inson Qadri Tamoyili</div>
            <div>Slayd {slide_num} / {len(SLIDES_DATA)}</div>
        </div>
    </body>
    </html>
    """

async def main():
    img_dir = r"C:\Users\user\Desktop\SlideTranslate_AI\rich_slides_img"
    os.makedirs(img_dir, exist_ok=True)
    
    image_paths = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        
        for idx, sdata in enumerate(SLIDES_DATA, 1):
            html = generate_html(sdata, idx)
            await page.set_content(html)
            # Wait for network idle to ensure icons and Unsplash images load completely
            try:
                await page.wait_for_load_state("networkidle", timeout=5000)
            except Exception:
                pass
            img_p = os.path.join(img_dir, f"slide_{idx}.png")
            await page.screenshot(path=img_p)
            image_paths.append(img_p)
            print(f"Rendered Visual Rich Slide {idx}/{len(SLIDES_DATA)}")
            
        await browser.close()

    # Build High Quality PPTX
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for img_p in image_paths:
        slide = prs.slides.add_slide(blank_layout)
        slide.shapes.add_picture(img_p, 0, 0, Inches(13.333), Inches(7.5))

    out_pptx = r"C:\Users\user\Desktop\SlideTranslate_AI\generated_presentations\Yangi_Ozbekiston_Visual_Rich_Masterpiece.pptx"
    os.makedirs(os.path.dirname(out_pptx), exist_ok=True)
    prs.save(out_pptx)
    print(f"VISUAL_RICH_MASTERPIECE_CREATED: {out_pptx}")

if __name__ == "__main__":
    asyncio.run(main())

import asyncio
from pyppeteer import launch
import pandas as pd
import time


class Amaliyotchi():
    def __init__(self, Fish, Gruh, Company, KFIO, KorRaxbar):
        self.Fish = Fish,
        self.Gruh = Gruh,
        self.Company = Company,
        self.KFIO = KFIO,
        self.KorRaxbar = KorRaxbar


async def generate_pdf_from_html(html_content, pdf_path):
    browser = await launch()
    page = await browser.newPage()

    await page.setContent(html_content)

    await page.pdf({'path': pdf_path, 'format': 'A4'})

    await browser.close()


excel_path = 'exel.xlsx'
df = pd.read_excel(excel_path)
list_amaliyot = []
for index, row in df.iterrows():
    # Har bir ustunni o'qib olib, PDF ga yozamiz
    amaliyotchi = Amaliyotchi(row['Fish'], row['Gruh'], row['Company'], row['KFIO'], row['KorRaxbar'])
    list_amaliyot.append(amaliyotchi)


def format(amaliyotchi):
    html_content = '''
<body>
    <h1>YO’LLANMA</h1>
    <div class="content">
        <p>Muxammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universiteti O’zbekiston Respublikasi oliy va o’quv yurtlari talabalarining amaliy malakasi oshirish to’g’risidagi qarori va O’zbekiston Respublikasi raqamli texnologiyalar vazirligi hamda Muhammad al-Xorazmiy nomidagi TATU o’rtasidagi shartnoma asosida Dasturiy injiniring fakulteti bakalavriat talabasi</p>
        <p style="text-align: center;">{Fish}</p>
        <p>bitiruv oldi amaliyotni o’tash uchun {Company} akademiyasi yubordi.</p>
        <p>Amaliyot muddati ______ yildan _______ yilgacha</p>
        <p style="text-align: right;">DIF dekani O.Ro’ziboyev</p>

        <div class="signature-section">
            <div>
                <div>
                    <p>TATU dan ketdi “__” _____ 2024-yil</p>
                    <p>Korxonaga keldi “___” ________2024-yil</p>
                    <p>M.O’ ________</p>
                    <p>(imzo)</p>
                </div>
                <div>
                    <p>Korxonadan ketdi “_____” ________ 2024-yil</p>
                    <p>TATU ga keldi “__”______2024-yil</p>
                    <p>M.O’ ________</p>
                    <p>(imzo)</p>
                </div>
            </div>
        </div>

        <h2 style="text-align: center;">TEXNIKA XAVFSIZLIGI BO’YICHA IMTIXON</h2>
        <p>Muhammad al-Xorazmiy nomidagi Toshkent axborot texnologiyalari universitetining amaliyoti talabasi</p>
        <p style="text-align: center;">{Fish}</p>
        <p>{Company} amaliyotga keldi texnika xavfsizligi boyicha imtihoni “___” bahoga topshirdi.</p>
        <p style="text-align: right;">{Fish} </p>
        <p>Talaba __________________________________________________________________________ ishla quyildi ____________________________________________________________</p>
        <p style="text-align: center;">(ish turi)</p>
        <p style="text-align: right;">Kommissiya raisi ________________________________________(imzo)</p>

        <h2 style="text-align: center;">AMALIYOT ISHLARINING QISQACHA TAVSIFI</h2>
        <p>{Company}</p>
        <p>Amaliyot boshlash sanasi “__” ________2024 yil</p>
        <p>Talaba ishining tavsifi (amaliyot dasturini bajarilishi, AKT yangiliklarini egallash amaliy ko’nikmalariga ega bo’lish va b.q)</p>
        <p>__________________________________________________________________________</p>
        <p>__________________________________________________________________________</p>
        <p>Amaliyot tugash muddati “___” _______ 2024-yil</p>
        <p>Korxonadagi amaliyot rahbarining imzosi _____________________________</p>
        <p style="text-align: right;">M.O’</p>

        <h2 style="text-align: center;">KAFEDRA XULOSASI :</h2>
        <p>__________________________________________________________________________</p>
        <p>_______________________bahoga topshirdi</p>
        <p style="text-align: right;">Kafedra mudiri ___________________</p>
    </div>
</body>
</html>
'''.format(Fish=amaliyotchi.Fish[0], Company=amaliyotchi.Company[0],)
    return html_content


def Pdf_yaratish():
    for amaliyotchi in list_amaliyot:
        pdf_path = f'pdf/{amaliyotchi.Fish[0]}.pdf'
        asyncio.get_event_loop().run_until_complete(generate_pdf_from_html(format(amaliyotchi), pdf_path))


if __name__ == '__main__':
    reexcel_path = 'exel.xlsx'
    Pdf_yaratish()

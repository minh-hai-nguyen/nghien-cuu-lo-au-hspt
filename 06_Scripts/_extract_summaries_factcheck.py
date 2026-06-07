from docx import Document
import sys, io

files = [
 ('QT001','Tom-tat-tung-bai/QT001_Jenkins_et_al_2023_USA.docx'),
 ('QT010','Tom-tat-tung-bai/QT010_Xu_et_al_2021_China_LargestEpi_FIXED_27052026.docx'),
 ('QT020','Tom-tat-tung-bai/QT020_Liu_CBT_Delivery_GAD_NMA_2025.docx'),
 ('QT030','Tom-tat-tung-bai/QT030_GBD_Trends_10-24y_2025.docx'),
 ('QT040','Tom-tat-tung-bai/QT040_Walder_JMIR_DMHI_SAD_2025.docx'),
 ('QT052','Tom-tat-tung-bai/QT052_Mindfulness_NatureMH_IPD_MA_2023.docx'),
 ('QT060','Tom-tat-tung-bai/QT060_Bie_GlobalAnxiety_GBD_10-24y_1990_2021_FrontPsych_2024.docx'),
 ('QT072','Tom-tat-tung-bai/QT072_Lee_2025_CyberbullyingMeta_TVA.docx'),
 ('VN001','Tom-tat-tung-bai/VN001_Hoa_2024_Frontiers_HaNoi_FIXED_27052026.docx'),
 ('VN002','Tom-tat-tung-bai/VN002_VNAMHS_2022_National.docx'),
]

with open('06_Scripts/_summaries_extract.txt','w',encoding='utf-8') as out:
    for id_, fp in files:
        out.write(f'\n\n========== {id_} : {fp} ==========\n')
        try:
            d = Document(fp)
            text = '\n'.join(p.text for p in d.paragraphs if p.text.strip())
            out.write(text[:5000])
        except Exception as e:
            out.write(f'ERR {e}')
print('Done')

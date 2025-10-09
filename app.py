import streamlit as st
import pandas as pd
import io
import chardet
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import Workbook

# 🌞 Беттің жалпы түрі (фон түсі және стиль)
st.set_page_config(page_title="Оқу жетістіктерін талдау", page_icon="📚", layout="centered")

# CSS арқылы әдемі дизайн беру
st.markdown("""
    <style>
    body {
        background: linear-gradient(135deg, #e0f7fa, #ffffff);
        font-family: 'Segoe UI', sans-serif;
    }
    .stApp {
        background: linear-gradient(135deg, #e0f7fa, #ffffff);
    }
    h1 {
        color: #0078D7 !important;
        text-align: center;
    }
    .css-1v3fvcr, .css-10trblm {
        color: #0078D7;
    }
    </style>
""", unsafe_allow_html=True)

# 🎯 Тақырып
st.title("📘 Оқушылардың оқу жетістіктерін талдау")
st.markdown("### Деректерді талдап, әр оқушыға жеке ұсыныс алыңыз 💡")

# 📂 Файл жүктеу
uploaded_file = st.file_uploader("📂 Файлды жүктеңіз:", type=["csv", "xlsx", "xls", "json", "txt"])

# 📥 Файлды оқу функциясы
def load_file(uploaded_file):
    if uploaded_file is not None:
        file_type = uploaded_file.name.split('.')[-1]

        if file_type == 'csv':
            raw_data = uploaded_file.read()
            detected_encoding = chardet.detect(raw_data)['encoding']
            df = pd.read_csv(io.BytesIO(raw_data), encoding=detected_encoding, low_memory=False)
        elif file_type in ['xls', 'xlsx']:
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        elif file_type == 'json':
            df = pd.read_json(uploaded_file)
        elif file_type == 'txt':
            df = pd.read_csv(uploaded_file, delimiter="\t", encoding="utf-8", low_memory=False)
        else:
            st.error("❌ Бұл файл түрі қолдау таппайды.")
            return None
        return df
    return None

# 📊 Орташа баллды есептеу және ұсыныс беру
def analyze_performance(data):
    numeric_data = data.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
    data['Орташа балл'] = numeric_data.mean(axis=1, skipna=True)
    recommendations = []
    
    for score in data['Орташа балл']:
        if pd.isna(score):
            rec = "❓ Мәлімет жеткіліксіз."
        elif score >= 9:
            rec = "🌟 Керемет! Осы қарқынмен жалғастырыңыз!"
        elif score >= 7:
            rec = "👍 Жақсы нәтиже! Тағы аздап еңбек етсеңіз, үздік боласыз."
        elif score >= 5:
            rec = "📘 Қосымша дайындалу қажет."
        else:
            rec = "🚀 Тьюторлық немесе жеке сабақтарды қарастырыңыз."
        recommendations.append(rec)
    
    data['Ұсыныс'] = recommendations
    return data

# 📥 Excel жүктеу
def download_excel(df):
    output = io.BytesIO()
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(list(df.columns))  # баған атаулары
    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            sheet.cell(row=r_idx, column=c_idx, value=value)
    workbook.save(output)
    return output.getvalue()

# 🏁 Негізгі логика
if uploaded_file:
    df = load_file(uploaded_file)
    if df is not None:
        result = analyze_performance(df)

        # 📋 Кесте
        st.success("✅ Файл сәтті талданды!")
        st.subheader("📊 Оқушылар нәтижесі:")
        st.dataframe(result, use_container_width=True)

        # 🎨 Дөңгелек диаграмма
        st.subheader("📈 Орташа балл үлесі:")
        fig, ax = plt.subplots(figsize=(5, 5))
        colors = sns.color_palette("cool", len(result))
        score_counts = result['Орташа балл'].round(1).value_counts().sort_index()
        ax.pie(score_counts, labels=score_counts.index, autopct='%1.1f%%', colors=colors, startangle=120)
        ax.set_title("Орташа балл үлестері")
        st.pyplot(fig)

        # 💾 Excel жүктеу
        excel_data = download_excel(result)
        st.download_button(
            label="📥 Excel форматында жүктеу",
            data=excel_data,
            file_name="Оку_жетистик_талдау.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("📄 Талдау үшін файлды жүктеңіз. CSV, Excel немесе JSON форматтары қолданылады.")

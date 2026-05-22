import streamlit as st
import pandas as pd
import io
import zipfile

# Настройка страницы
st.set_page_config(page_title="Разделитель Excel файлов", page_icon="✂️", layout="centered")

st.title("Утилита для разделения Excel-таблиц")
st.write(
    "Эта программа разделяет тяжелые файлы на части, строго сохраняя верхние технические строки и заголовки во всех новых файлах.")

# Загрузка файла
uploaded_file = st.file_uploader("Загрузите Excel файл (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file:
    # Настройки разделения выведены в две колонки для красоты
    col1, col2 = st.columns(2)
    with col1:
        tech_rows = st.number_input(
            "Сколько строк отвести под заголовок?",
            min_value=1,
            value=2,
            help="Например, 1 строка с техническим названием и 1 строка с названиями колонок = 2 строки."
        )
    with col2:
        chunk_size = st.number_input(
            "Количество строк данных в одном файле",
            min_value=1,
            value=1000
        )

    if st.button("Разделить файл", type="primary"):
        with st.spinner("Обработка файла..."):
            try:
                # Читаем верхние строки (шапку)
                df_headers = pd.read_excel(uploaded_file, nrows=tech_rows - 1, header=None)
                # Читаем основные данные, пропуская шапку
                df_data = pd.read_excel(uploaded_file, skiprows=tech_rows - 1)

                # Создаем ZIP архив прямо в оперативной памяти (без сохранения на диск)
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

                    # Цикл для дробления данных на куски
                    for i in range(0, len(df_data), chunk_size):
                        chunk = df_data.iloc[i:i + chunk_size]

                        # Создаем временный Excel-файл в памяти
                        output_buffer = io.BytesIO()
                        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                            # Записываем шапку
                            if not df_headers.empty:
                                df_headers.to_excel(writer, index=False, header=False, startrow=0)
                            # Записываем кусок данных
                            chunk.to_excel(writer, index=False, startrow=tech_rows - 1)

                        # Кладем готовый кусок в ZIP-архив
                        file_num = i // chunk_size + 1
                        zip_file.writestr(f"split_part_{file_num}.xlsx", output_buffer.getvalue())

                st.success("Файл успешно разделен!")
                st.download_button(
                    label="Скачать архив с результатами (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="split_excel_files.zip",
                    mime="application/zip"
                )
            except Exception as e:
                st.error(f"Произошла ошибка при обработке файла: {e}")
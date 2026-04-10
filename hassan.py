import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
from xlsxwriter.utility import xl_col_to_name

# إعدادات الصفحة
st.set_page_config(page_title="ICU Roster Generator", page_icon="🏥")

st.title("🏥 نظام إدارة جداول تمريض العناية المركزة")
st.subheader("إنشاء جدول شفتات MLLNNOO ذكي")

# واجهة المدخلات في الشريط الجانبي (Sidebar)
with st.sidebar:
    st.header("إعدادات الجدول")
    year = st.number_input("السنة", min_value=2024, max_value=2030, value=datetime.now().year)
    month = st.number_input("الشهر", min_value=1, max_value=12, value=datetime.now().month)
    num_staff = st.slider("عدد طاقم التمريض", min_value=1, max_value=50, value=14)
    
    generate_btn = st.button("توليد الجدول الآن")

if generate_btn:
    try:
        pattern = ['M', 'L', 'L', 'N', 'N', 'O', 'O']
        
        # حساب الأيام
        start_date = datetime(year, month, 1)
        if month == 12:
            end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = datetime(year, month + 1, 1) - timedelta(days=1)
        
        days_list = [(start_date + timedelta(days=x)) for x in range((end_date - start_date).days + 1)]
        day_names = [d.strftime('%a') for d in days_list]

        # بناء البيانات
        data = {}
        for i in range(num_staff):
            staff_name = f"Nurse {i+1}"
            data[staff_name] = [pattern[(d + i) % len(pattern)] for d in range(len(days_list))]

        df = pd.DataFrame(data)
        df.insert(0, 'Day', day_names)
        df.index = [d.strftime('%d-%m') for d in days_list]

        # معاينة الجدول في الموقع
        st.success(f"تم إنشاء الجدول لشهر {month} بنجاح!")
        st.dataframe(df, height=400)

        # تحويل الإكسيل إلى "ذاكرة مؤقتة" لتحميله عبر الويب
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Roster')
        
        workbook  = writer.book
        worksheet = writer.sheets['Roster']
        
        # التنسيقات (نفس منطق تطبيق سطح المكتب)
        header_f = workbook.add_format({'bg_color': '#1A237E', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center'})
        stat_f = workbook.add_format({'bg_color': '#E8EAF6', 'bold': True, 'border': 1, 'align': 'center'})
        
        fmt_m = workbook.add_format({'bg_color': '#FFF9C4', 'border': 1, 'align': 'center'})
        fmt_l = workbook.add_format({'bg_color': '#C8E6C9', 'border': 1, 'align': 'center'})
        fmt_n = workbook.add_format({'bg_color': '#BBDEFB', 'border': 1, 'align': 'center'})
        fmt_o = workbook.add_format({'bg_color': '#FFCDD2', 'border': 1, 'align': 'center'})

        last_row_num = len(days_list) + 1
        last_col_letter = xl_col_to_name(num_staff + 1)
        full_range = f"C2:{last_col_letter}{last_row_num}"

        worksheet.conditional_format(full_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"M"', 'format': fmt_m})
        worksheet.conditional_format(full_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"L"', 'format': fmt_l})
        worksheet.conditional_format(full_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"N"', 'format': fmt_n})
        worksheet.conditional_format(full_range, {'type': 'cell', 'criteria': 'equal to', 'value': '"O"', 'format': fmt_o})

        # تنسيق العناوين والحسابات
        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num + 1, value, header_f)
        worksheet.write(0, 0, 'Date', header_f)

        for idx, label in enumerate(['Total Hours', 'L Count', 'N Count', 'M Count']):
            row = last_row_num + idx
            worksheet.write(row, 1, label, header_f)
            for col_num in range(num_staff):
                col_let = xl_col_to_name(col_num + 2)
                d_range = f"{col_let}2:{col_let}{last_row_num}"
                if idx == 0: form = f'=(COUNTIF({d_range},"L")*12)+(COUNTIF({d_range},"N")*12)+(COUNTIF({d_range},"M")*6)'
                elif idx == 1: form = f'=COUNTIF({d_range},"L")'
                elif idx == 2: form = f'=COUNTIF({d_range},"N")'
                else: form = f'=COUNTIF({d_range},"M")'
                worksheet.write_formula(row, col_num + 2, form, stat_f)

        worksheet.set_column(2, num_staff + 1, 6)
        worksheet.set_column(0, 1, 10)
        writer.close()
        
        # زر تحميل الملف
        st.download_button(
            label="📥 تحميل جدول الإكسيل الملون",
            data=output.getvalue(),
            file_name=f"ICU_Roster_{year}_{month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"حدث خطأ: {e}")
else:
    st.info("قم بتعديل الخيارات من القائمة الجانبية ثم اضغط على زر التوليد.")

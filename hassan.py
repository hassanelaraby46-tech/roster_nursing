import customtkinter as ctk
import pandas as pd
from datetime import datetime, timedelta
from tkinter import messagebox

# إعداد مظهر الواجهة
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class RosterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("نظام إدارة تمريض العناية المركزة - ICU Roster")
        self.geometry("450x500")

        # العنوان الرئيسي
        self.label_title = ctk.CTkLabel(self, text="إنشاء جدول شفتات MLLNNOO", font=("Arial", 20, "bold"))
        self.label_title.pack(pady=20)

        # مدخل السنة
        self.label_year = ctk.CTkLabel(self, text="السنة (Year):")
        self.label_year.pack(pady=5)
        self.entry_year = ctk.CTkEntry(self, placeholder_text="2024")
        self.entry_year.insert(0, str(datetime.now().year))
        self.entry_year.pack(pady=5)

        # مدخل الشهر
        self.label_month = ctk.CTkLabel(self, text="الشهر (1-12):")
        self.label_month.pack(pady=5)
        self.entry_month = ctk.CTkEntry(self, placeholder_text="5")
        self.entry_month.insert(0, str(datetime.now().month))
        self.entry_month.pack(pady=5)

        # مدخل عدد التمريض
        self.label_staff = ctk.CTkLabel(self, text="عدد طاقم التمريض:")
        self.label_staff.pack(pady=5)
        self.entry_staff = ctk.CTkEntry(self, placeholder_text="14")
        self.entry_staff.insert(0, "14")
        self.entry_staff.pack(pady=5)

        # زر الإنشاء
        self.btn_generate = ctk.CTkButton(self, text="استخراج جدول إكسيل ذكي", command=self.generate_roster_logic)
        self.btn_generate.pack(pady=30)

        # تذييل
        self.label_footer = ctk.CTkLabel(self, text="بناءً على نمط التوزيع العادل MLLNNOO", font=("Arial", 10))
        self.label_footer.pack(side="bottom", pady=10)

    def generate_roster_logic(self):
        try:
            year = int(self.entry_year.get())
            month = int(self.entry_month.get())
            num_staff = int(self.entry_staff.get())
            
            pattern = ['M', 'L', 'L', 'N', 'N', 'O', 'O']
            
            # حساب الأيام
            start_date = datetime(year, month, 1)
            if month == 12:
                end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
            else:
                end_date = datetime(year, month + 1, 1) - timedelta(days=1)
            days_list = [(start_date + timedelta(days=x)) for x in range((end_date - start_date).days + 1)]

            # بناء البيانات
            data = {}
            day_names = [d.strftime('%a') for d in days_list]
            for i in range(num_staff):
                staff_name = f"Nurse {i+1}"
                data[staff_name] = [pattern[(d + i) % len(pattern)] for d in range(len(days_list))]

            df = pd.DataFrame(data)
            df.insert(0, 'Day', day_names)
            df.index = [d.strftime('%d-%m') for d in days_list]

            # تصدير الإكسيل (نفس منطق التنسيق السابق)
            file_name = f"ICU_Roster_{year}_{month}.xlsx"
            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Roster')
            
            workbook  = writer.book
            worksheet = writer.sheets['Roster']
            
            # تنسيقات سريعة
            header_f = workbook.add_format({'bg_color': '#1A237E', 'font_color': 'white', 'bold': True, 'border': 1, 'align': 'center'})
            stat_f = workbook.add_format({'bg_color': '#E8EAF6', 'bold': True, 'border': 1, 'align': 'center'})
            
            # كتابة العناوين وتنسيق الشفتات (اختصاراً للمساحة استخدمنا التنسيق الشرطي)
            for col_num, value in enumerate(df.columns):
                worksheet.write(0, col_num + 1, value, header_f)
            worksheet.write(0, 0, 'Date', header_f)

            # إضافة صفوف الحسابات أسفل الجدول
            last_row = len(days_list) + 1
            calc_labels = ['Total Hours', 'L Count', 'N Count', 'M Count']
            for idx, label in enumerate(calc_labels):
                worksheet.write(last_row + idx, 1, label, header_f)
                for col_num in range(num_staff):
                    col_letter = chr(67 + col_num)
                    data_range = f"{col_letter}2:{col_letter}{last_row}"
                    if idx == 0: formula = f'=(COUNTIF({data_range},"L")*12)+(COUNTIF({data_range},"N")*12)+(COUNTIF({data_range},"M")*6)'
                    elif idx == 1: formula = f'=COUNTIF({data_range},"L")'
                    elif idx == 2: formula = f'=COUNTIF({data_range},"N")'
                    else: formula = f'=COUNTIF({data_range},"M")'
                    worksheet.write_formula(last_row + idx, col_num + 2, formula, stat_f)

            worksheet.set_column(2, num_staff + 1, 6)
            worksheet.set_column(0, 1, 10)
            writer.close()

            messagebox.showinfo("نجاح", f"تم إنشاء الملف بنجاح!\nاسم الملف: {file_name}")

        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ: {str(e)}")

if __name__ == "__main__":
    app = RosterApp()
    app.mainloop()
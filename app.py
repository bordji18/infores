from flask import Flask, render_template, request
import pandas as pd
import os
import unicodedata

app = Flask(__name__)

# تحميل الملف مرة واحدة عند بداية التشغيل
excel_path = 'static/data.xlsx'
sheets = pd.read_excel(excel_path, sheet_name=None)

# تطبيع الحروف العربية
def normalize_arabic(text):
    text = str(text)
    text = text.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
    return text.strip()

@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    error = None

    if request.method == 'POST':
        section = request.form['section']
        last_name = normalize_arabic(request.form['last_name'])
        first_name = normalize_arabic(request.form['first_name'])
        birth_date = request.form['birth_date']

        sheet = sheets.get(section)
        if sheet is None:
            error = "القسم غير موجود."
        else:
            # تطبيع الأسماء داخل الجدول أيضاً
            sheet['اللقب_معدل'] = sheet.iloc[:, 1].apply(normalize_arabic)
            sheet['الاسم_معدل'] = sheet.iloc[:, 2].apply(normalize_arabic)

            match = sheet[
                (sheet['اللقب_معدل'] == last_name) &
                (sheet['الاسم_معدل'] == first_name) &
                (sheet.iloc[:, 3].astype(str).str.strip() == birth_date)
            ]

            if not match.empty:
                average = match.iloc[0, 6]
                exam = match.iloc[0, 7]

                result = {
                    'last_name': request.form['last_name'],
                    'first_name': request.form['first_name'],
                    'average': average if pd.notna(average) else "غائب",
                    'exam': exam if pd.notna(exam) else "غائب"
                }
            else:
                error = "لا توجد نتيجة بهذا الاسم وتاريخ الميلاد في هذا القسم."

    return render_template('index.html', result=result, error=error)
    
if __name__ == "__main__":
    app.run(debug=True)

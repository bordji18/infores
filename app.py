from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)
EXCEL_PATH = "static/data.xlsx"

# تحميل الملف مرة واحدة لكل ورقة
sheet_data = pd.read_excel(EXCEL_PATH, sheet_name=None)

@app.route("/", methods=["GET", "POST"])
def index():
    result = None
    error = None

    if request.method == "POST":
        section = request.form.get("section")
        last_name = request.form.get("last_name", "").strip().lower()
        first_name = request.form.get("first_name", "").strip().lower()

        df = sheet_data.get(section)
        if df is not None:
            for _, row in df.iterrows():
                ln = str(row.iloc[1]).strip().lower()
                fn = str(row.iloc[2]).strip().lower()
                if ln == last_name and fn == first_name:
                    result = {
                        "last_name": row.iloc[1],
                        "first_name": row.iloc[2],
                        "average": row.iloc[6] if len(row) > 6 else "غير متوفر",
                        "exam": row.iloc[7] if len(row) > 7 else "غير متوفر"
                    }
                    break
            if not result:
                error = "لم يتم العثور على التلميذ. تأكد من المعلومات."
        else:
            error = "القسم غير موجود."

    return render_template("index.html", result=result, error=error)

if __name__ == "__main__":
    app.run(debug=True)

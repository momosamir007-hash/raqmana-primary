# app.py
import sys
import io
import re
import logging
import unicodedata
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.chart import BarChart, Reference
from streamlit.web import cli as stcli

# ==========================================
# 1. التكوينات والإعدادات (Configuration)
# ==========================================

DEFAULT_CONFIG = {
    "SUBJECT_PATTERNS": {
        "arabic": {"keywords": ["عربي", "العربية", "لغة عربية", "ع.ع", "ل.ع"], "lang": "ar"},
        "math": {"keywords": ["رياض", "رياضيات", "math", "ر.ع", "ريا"], "lang": "ar"},
        "french": {"keywords": ["فرنس", "الفرنسية", "français", "francais", "ف.ع", "ل.ف"], "lang": "fr"},
        "english": {"keywords": ["انجليز", "english", "ل.ا", "eng"], "lang": "en"},
        "tamazight": {"keywords": ["أمازيغ", "امازيغ", "tamazight", "ل.أ", "أمازيغية"], "lang": "ar"},
        "plastic_arts": {"keywords": ["تشكيل", "تشكيلية", "رسم", "art plastique", "ت.تشكيلية"], "lang": "ar"},
        "music": {"keywords": ["موسيق", "موسيقية", "نشيد", "musique", "ت.موسيقية"], "lang": "ar"},
        "physical_ed": {"keywords": ["بدنية", "رياضية", "sport", "eps", "ت.بدنية"], "lang": "ar"},
    },
    "COL_PATTERNS": {
        "ar_expr": ["تعبير", "تواصل شفوي", "شفوي", "التعبير"],
        "ar_read": ["قراءة", "محفوظات", "القراءة"],
        "ar_write": ["كتابة", "إملاء", "الكتابة"],
        "ma_num": ["أعداد", "عمليات", "حساب", "الأعداد"],
        "ma_meas": ["مقادير", "قياس", "القياس"],
        "ma_data": ["تنظيم", "معطيات", "إحصاء"],
        "ma_geo": ["فضاء", "هندسة", "الهندسة"],
        "fl_expr": ["expression", "orale", "شفهي", "تواصل"],
        "fl_read": ["lecture", "compréhension", "قراءة"],
        "fl_prod": ["production", "écrite", "إنتاج", "كتابي"],
        "plastic_eval": ["تقويم", "إنجاز", "عمل", "مشاركة"],
        "music_eval": ["نشيد", "أداء", "استماع", "إيقاع"],
        "sport_eval": ["أداء", "نشاط", "مهارة"],
        "exam": ["اختبار", "فرض", "امتحان", "exam", "test"],
        "remark": ["ملاحظة", "تقدير", "remarque", "appreciation", "تعليق"]
    },
    "REMARKS": {
        "ar": [(0.90, "عمل ممتاز جداً ومتميز"), (0.80, "عمل ممتاز"), (0.70, "عمل جيد جداً"), (0.65, "عمل جيد"), (0.55, "عمل مقبول"), (0.50, "متوسط"), (0.00, "دون المتوسط")],
        "fr": [(0.90, "Excellent"), (0.80, "Très bien"), (0.70, "Bien"), (0.65, "Assez bien"), (0.55, "Passable"), (0.50, "Suffisant"), (0.00, "Insuffisant")],
        "en": [(0.90, "Outstanding"), (0.80, "Excellent"), (0.70, "Very Good"), (0.65, "Good"), (0.55, "Acceptable"), (0.50, "Sufficient"), (0.00, "Insufficient")],
    },
    "COLORS": {
        "ar": {"عمل ممتاز جداً ومتميز": "1A6B3C", "عمل ممتاز": "27AE60", "عمل جيد جداً": "52BE80", "عمل جيد": "F39C12", "عمل مقبول": "E67E22", "متوسط": "E74C3C", "دون المتوسط": "C0392B", "غائب": "95A5A6", "معفى": "7F8C8D"},
        "fr": {"Excellent": "1A6B3C", "Très bien": "27AE60", "Bien": "52BE80", "Assez bien": "F39C12", "Passable": "E67E22", "Suffisant": "E74C3C", "Insuffisant": "C0392B", "Absent": "95A5A6", "Dispensé": "7F8C8D"},
        "en": {"Outstanding": "1A6B3C", "Excellent": "27AE60", "Very Good": "52BE80", "Good": "F39C12", "Acceptable": "E67E22", "Sufficient": "E74C3C", "Insufficient": "C0392B", "Absent": "95A5A6", "Exempt": "7F8C8D"}
    }
}

# ==========================================
# 2. الفئات المساعدة (Core Logic)
# ==========================================

class TextHelper:
    """فئة مساعدة لمعالجة النصوص وتوحيدها للبحث"""
    @staticmethod
    def normalize(text: str) -> str:
        if not text: return ""
        text = str(text).strip().lower()
        text = unicodedata.normalize("NFD", text)
        text = "".join(c for c in text if unicodedata.category(c) != "Mn")
        return re.sub(r"\s+", " ", text)


class GradeCalculator:
    """فئة مسؤولة عن العمليات الحسابية ومنطق التقييم"""
    def __init__(self, config):
        self.remarks = config["REMARKS"]

    def parse_grade(self, value) -> dict:
        if value is None or str(value).strip() == "":
            return {"status": "empty", "value": None}
            
        v_str = str(value).strip().lower()
        if v_str in ["غ", "غائب", "abs", "absent"]: return {"status": "absent", "value": "غائب"}
        if v_str in ["م", "معفى", "dispensé", "dispense", "exempt", "مُعفى"]: return {"status": "exempt", "value": "معفى"}
            
        try:
            num = float(v_str.replace(",", "."))
            if num != num: return {"status": "invalid", "value": None}
            return {"status": "valid", "value": num}
        except ValueError:
            return {"status": "text", "value": v_str}

    def detect_max_grade(self, header_val, default: float) -> float:
        if not header_val: return default
        m = re.search(r"/\s*(\d+(?:\.\d+)?)", str(header_val))
        if m and float(m.group(1)) > 0: return float(m.group(1))
        if "20" in str(header_val): return 20.0
        if "10" in str(header_val): return 10.0
        return default

    def calc_average(self, values: list, ignore_zero: bool) -> float:
        valid = [v for v in values if v is not None]
        if ignore_zero: valid = [v for v in valid if v != 0.0]
        return sum(valid) / len(valid) if valid else None


class ExcelProcessor:
    """المحرك الأساسي لمعالجة ملف الإكسيل وتطبيق القواعد"""
    def __init__(self, config, user_settings, logger):
        self.config = config
        self.settings = user_settings
        self.logger = logger
        self.calc = GradeCalculator(config)
        self.stats = []

    def style_cell(self, cell, remark: str, lang: str):
        color_map = self.config["COLORS"].get(lang, self.config["COLORS"]["ar"])
        hex_col = color_map.get(remark, "7F8C8D")
        cell.fill = PatternFill(fill_type="solid", fgColor=hex_col)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def detect_subject(self, sheet_name: str):
        norm = TextHelper.normalize(sheet_name)
        for stype, info in self.config["SUBJECT_PATTERNS"].items():
            for kw in info["keywords"]:
                if TextHelper.normalize(kw) in norm: return stype, info["lang"]
        return "other", "ar"

    def find_column(self, ws, keywords, max_row=80):
        norm_kws = [TextHelper.normalize(kw) for kw in keywords]
        for row in ws.iter_rows(min_row=1, max_row=min(max_row, ws.max_row or 1)):
            for cell in row:
                if cell.value:
                    norm_val = TextHelper.normalize(str(cell.value))
                    if any(kw in norm_val for kw in norm_kws):
                        return {"col": cell.column, "row": cell.row, "header": str(cell.value)}
        return None

    def find_fallback_numeric_cols(self, ws, header_row, exclude_cols, max_grade):
        numeric_cols = []
        if not ws.max_row or ws.max_row <= header_row: return numeric_cols
        
        check_rows = range(header_row + 1, min(header_row + 6, ws.max_row + 1))
        for col in range(1, (ws.max_column or 1) + 1):
            if col in exclude_cols: continue
            for r in check_rows:
                parsed = self.calc.parse_grade(ws.cell(row=r, column=col).value)
                if parsed["status"] == "valid" and 0 <= parsed["value"] <= max_grade * 1.1:
                    numeric_cols.append(col)
                    break
        return numeric_cols

    def map_columns(self, ws, subject_type, default_max):
        res = {"found": False, "remark_col": None, "exam_col": None, "eval_cols": [], "header_row": 1, "exam_header": "", "method": "none", "notes": ""}
        
        rem_info = self.find_column(ws, self.config["COL_PATTERNS"]["remark"])
        if not rem_info:
            res["notes"] = "لم يُكتشف عمود الملاحظات"
            return res
            
        res["remark_col"] = rem_info["col"]
        res["header_row"] = rem_info["row"]

        exam_info = self.find_column(ws, self.config["COL_PATTERNS"]["exam"])
        if exam_info:
            res["exam_col"] = exam_info["col"]
            res["exam_header"] = exam_info["header"]
            res["header_row"] = max(res["header_row"], exam_info["row"])

        eval_groups = []
        if subject_type == "arabic": eval_groups = ["ar_expr", "ar_read", "ar_write"]
        elif subject_type == "math": eval_groups = ["ma_num", "ma_meas", "ma_data", "ma_geo"]
        elif subject_type in ("french", "english", "tamazight"): eval_groups = ["fl_expr", "fl_read", "fl_prod"]
        elif subject_type == "plastic_arts": eval_groups = ["plastic_eval"]
        elif subject_type == "music": eval_groups = ["music_eval"]
        elif subject_type == "physical_ed": eval_groups = ["sport_eval"]

        for grp in eval_groups:
            info = self.find_column(ws, self.config["COL_PATTERNS"].get(grp, []))
            if info:
                res["eval_cols"].append(info["col"])
                res["header_row"] = max(res["header_row"], info["row"])

        needs_fallback = (subject_type in ("plastic_arts", "music", "physical_ed", "other") 
                          and not res["eval_cols"] and res["exam_col"] is None)
        
        if needs_fallback:
            num_cols = self.find_fallback_numeric_cols(ws, res["header_row"], {res["remark_col"]}, default_max)
            if num_cols:
                res["eval_cols"] = num_cols
                res["method"] = "fallback"
            else:
                res["notes"] = "لم تُكتشف أي أعمدة رقمية"
        else:
            res["method"] = "keyword"

        res["found"] = bool(res["eval_cols"] or res["exam_col"] is not None)
        return res

    def compute_final_grade(self, eval_vals, exam_val, subject_type, method):
        eval_avg = self.calc.calc_average(eval_vals, self.settings["ignore_zero"]) if eval_vals else None

        if subject_type in ("arabic", "math", "french", "english", "tamazight"):
            if eval_avg is not None and exam_val is not None: return (eval_avg + exam_val) / 2
            if eval_avg is not None: return eval_avg
            if exam_val is not None: return exam_val
            
        elif subject_type in ("plastic_arts", "music", "physical_ed"):
            if method == "fallback": return eval_avg
            if eval_avg is not None and exam_val is not None: return (eval_avg + exam_val) / 2
            if eval_avg is not None: return eval_avg
            if exam_val is not None: return exam_val
            
        else:
            if exam_val is not None: return exam_val
            if eval_avg is not None: return eval_avg
            
        return None

    def process_workbook(self, file_bytes):
        wb = openpyxl.load_workbook(file_bytes)
        report = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            subj, lang = self.detect_subject(sheet_name)
            cols = self.map_columns(ws, subj, self.settings["max_grade"])
            
            if not cols["found"]:
                self.logger.warning(f"تخطي الورقة '{sheet_name}': {cols.get('notes', 'أعمدة مفقودة')}")
                continue

            sheet_max = self.calc.detect_max_grade(cols["exam_header"], self.settings["max_grade"])
            added, skipped = 0, 0
            sheet_grades = []

            for row_idx in range(cols["header_row"] + 1, ws.max_row + 1):
                remark_cell = ws.cell(row=row_idx, column=cols["remark_col"])
                existing = remark_cell.value
                has_remark = existing is not None and str(existing).strip() != ""

                if self.settings["overwrite_mode"] == "تخطي الكل":
                    skipped += 1; continue
                if self.settings["overwrite_mode"] == "ملء الفارغة فقط" and has_remark:
                    skipped += 1; continue

                eval_vals, exam_val = [], None
                is_absent, is_exempt = False, False

                for c in cols["eval_cols"]:
                    p = self.calc.parse_grade(ws.cell(row=row_idx, column=c).value)
                    if p["status"] == "absent": is_absent = True
                    elif p["status"] == "exempt": is_exempt = True
                    elif p["status"] == "valid": eval_vals.append(p["value"])

                if cols["exam_col"]:
                    p = self.calc.parse_grade(ws.cell(row=row_idx, column=cols["exam_col"]).value)
                    if p["status"] == "absent": is_absent = True
                    elif p["status"] == "exempt": is_exempt = True
                    elif p["status"] == "valid": exam_val = p["value"]

                if is_absent or is_exempt:
                    txt = "Absent" if lang != "ar" else "غائب"
                    if is_exempt: txt = "Dispensé" if lang == "fr" else ("Exempt" if lang == "en" else "معفى")
                    remark_cell.value = txt
                    if self.settings["apply_style"]: self.style_cell(remark_cell, txt, lang)
                    added += 1
                    continue

                final = self.compute_final_grade(eval_vals, exam_val, subj, cols["method"])
                
                if final is not None:
                    sheet_grades.append(final)
                    final_rounded = round(final, self.settings["round_dec"])
                    remark_txt = self.calc.get_remark(final_rounded, sheet_max, lang)
                    remark_cell.value = remark_txt
                    if self.settings["apply_style"]: self.style_cell(remark_cell, remark_txt, lang)
                    added += 1
                else:
                    skipped += 1

            if sheet_grades:
                self.stats.append({
                    "المادة": sheet_name,
                    "السلم": sheet_max,
                    "أعلى معدل": round(max(sheet_grades), 2),
                    "أدنى معدل": round(min(sheet_grades), 2),
                    "المعدل العام": round(sum(sheet_grades) / len(sheet_grades), 2),
                    "نسبة النجاح %": round((sum(1 for g in sheet_grades if g >= (sheet_max / 2)) / len(sheet_grades)) * 100, 1)
                })

            formula_desc = "غير محدد"
            n_eval = len(cols["eval_cols"])
            if subj == "other": formula_desc = "نقطة الاختبار مباشرة"
            elif cols["method"] == "fallback": formula_desc = f"معدل {n_eval} عمود رقمي (تلقائي)"
            elif n_eval > 0 and cols["exam_col"]: formula_desc = f"معدل {n_eval} تقويم + اختبار ÷ 2"
            elif n_eval > 0: formula_desc = f"معدل {n_eval} تقويم"
            elif cols["exam_col"]: formula_desc = "نقطة الاختبار"

            report.append({
                "الورقة": sheet_name, 
                "القاعدة الحسابية": formula_desc, 
                "ملاحظات أضيفت": added, 
                "تم تخطيها": skipped
            })

        self.generate_dashboard(wb)
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output, report

    def generate_dashboard(self, wb):
        if not self.stats: return
        ws = wb.create_sheet(title="إحصائيات القسم")
        headers = ["المادة", "السلم", "أعلى معدل", "أدنى معدل", "المعدل العام", "نسبة النجاح %"]
        ws.append(headers)
        
        for i, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=i)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor="2980B9")
            cell.alignment = Alignment(horizontal="center")
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 15

        for stat in self.stats:
            ws.append([stat["المادة"], stat["السلم"], stat["أعلى معدل"], stat["أدنى معدل"], stat["المعدل العام"], stat["نسبة النجاح %"]])

        chart = BarChart()
        chart.type, chart.style = "col", 10
        chart.title = "المعدل العام ونسبة النجاح"
        chart.add_data(Reference(ws, min_col=5, min_row=1, max_row=len(self.stats)+1), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=len(self.stats)+1))
        ws.add_chart(chart, "H2")

# ==========================================
# 5. واجهة المستخدم (Streamlit App)
# ==========================================

def run_ui():
    st.set_page_config(page_title="أداة الملاحظات الذكية", page_icon="🎓", layout="wide")

    st.markdown("""
    <style>
        .stApp { direction: rtl; text-align: right; font-family: 'Tajawal', sans-serif;}
        .css-1d391kg { direction: rtl; } 
        .stButton>button { width: 100%; border-radius: 5px; font-weight: bold; background-color: #27AE60; color: white;}
        .stButton>button:hover { background-color: #1A6B3C;}
    </style>
    """, unsafe_allow_html=True)

    st.title("🎓 أداة حجز الملاحظات الذكية — الطور الابتدائي (V4.1)")
    st.markdown("نظام أوتوماتيكي متكامل لحساب المعدلات، تدوين الملاحظات الذكية، واستخراج إحصائيات القسم من ملفات الإكسيل.")

    with st.sidebar:
        st.header("⚙️ إعدادات المعالجة")
        max_grade = st.selectbox("سلم التنقيط الافتراضي:", options=[10, 20], index=0)
        round_dec = st.selectbox("دقة تقريب النقطة:", options=[2, 1, 0], index=0)
        overwrite_mode = st.radio("وضع الملاحظات الموجودة:", ["استبدال الموجودة", "ملء الفارغة فقط", "تخطي الكل"])
        ignore_zero = st.checkbox("تجاهل الأصفار عند حساب التقويمات؟", value=False)
        apply_style = st.checkbox("تلوين الخلايا (أخضر/أحمر)؟", value=True)

    user_settings = {
        "max_grade": max_grade, "round_dec": round_dec,
        "overwrite_mode": overwrite_mode, "ignore_zero": ignore_zero, "apply_style": apply_style
    }

    uploaded_file = st.file_uploader("📂 اسحب وأفلت ملف الإكسيل هنا (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        if st.button("🚀 بدء المعالجة الذكية"):
            
            logger = logging.getLogger("GradeTool")
            logger.setLevel(logging.WARNING)
            log_stream = io.StringIO()
            ch = logging.StreamHandler(log_stream)
            ch.setFormatter(logging.Formatter('⚠️ %(message)s'))
            if not logger.handlers: logger.addHandler(ch)

            with st.spinner('جاري التحليل المعمق للملف وتطبيق القواعد...'):
                processor = ExcelProcessor(DEFAULT_CONFIG, user_settings, logger)
                try:
                    processed_bytes, report_data = processor.process_workbook(uploaded_file)
                    
                    st.success("✅ اكتملت المعالجة بنجاح! تم تطبيق المعادلات وإنشاء ورقة الإحصائيات.")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.subheader("📋 تقرير المعالجة المفصل")
                        if report_data:
                            st.dataframe(pd.DataFrame(report_data), use_container_width=True)
                        else:
                            st.warning("لم يتم معالجة أي أوراق.")

                    with col2:
                        st.subheader("📊 لوحة قيادة القسم (مختصر)")
                        if processor.stats:
                            st.dataframe(pd.DataFrame(processor.stats), use_container_width=True)
                        else:
                            st.info("لا توجد بيانات كافية لإنشاء الإحصائيات.")

                    logs = log_stream.getvalue()
                    if logs:
                        st.warning("ملاحظات النظام:")
                        st.code(logs)

                    st.download_button(
                        label="📥 تحميل الملف الجاهز للرقمنة",
                        data=processed_bytes,
                        file_name=uploaded_file.name.replace(".xlsx", "_جاهز_للرقمنة.xlsx"),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"❌ حدث خطأ غير متوقع: {str(e)}")

# نقطة الدخول (تدعم التشغيل المباشر من Pydroid 3)
if __name__ == '__main__':
    if st._is_running_with_streamlit:
        run_ui()
    else:
        sys.argv = ["streamlit", "run", __file__]
        sys.exit(stcli.main())

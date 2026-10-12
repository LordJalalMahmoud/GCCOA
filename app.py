import sys
import os
import json
import subprocess
import re
import datetime
from typing import Optional

# Setup base app dir before anything else
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

os.environ["QT_LOGGING_RULES"] = "qt.text.font.db=false;*.debug=false"

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QLineEdit, QTextEdit, QPushButton, QComboBox,
    QSpinBox, QTableWidget, QTableWidgetItem, QHeaderView, QGroupBox,
    QMessageBox, QFileDialog, QSplitter, QFrame, QScrollArea, QDialog,
    QDialogButtonBox, QStatusBar, QDateEdit, QMenu
)
from PyQt6.QtCore import Qt, QTimer, QDate
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette, QAction

from order_parser import parse_order_text, parse_order_with_gemini, validate_egyptian_phone, BOSTA_CITIES, BOSTA_CITY_AREAS
from excel_manager import ExcelManager, APP_DIR, DEFAULT_EXCEL_FILENAME

CONFIG_FILE = os.path.join(APP_DIR, 'config.json')

def load_config():
    default_config = {
        "excel_path": os.path.join(APP_DIR, DEFAULT_EXCEL_FILENAME),
        "gemini_api_key": os.environ.get("GEMINI_API_KEY", "")
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                default_config.update(saved)
        except Exception:
            pass
    return default_config

def save_config(cfg):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class ExportDialog(QDialog):
    """Dialog to export orders for a specific day."""
    def __init__(self, parent=None, excel_manager=None):
        super().__init__(parent)
        self.excel_manager = excel_manager
        self.setWindowTitle("📅 تصدير شيت بوسطة ليوم محدد")
        self.setMinimumWidth(520)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        info_lbl = QLabel(
            "<b>تصدير شيت بوسطة اليومي:</b><br>"
            "اختر اليوم المراد تصديره لإنشاء ملف إكسيل مستقل وجاهز للرفع على منصة بوسطة مباشرة."
        )
        info_lbl.setWordWrap(True)
        layout.addWidget(info_lbl)

        # Date Picker
        date_group = QGroupBox("📅 اختر تاريخ اليوم")
        date_layout = QVBoxLayout(date_group)

        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDate(QDate.currentDate())
        self.date_picker.setDisplayFormat("yyyy-MM-dd")
        self.date_picker.dateChanged.connect(self.update_count)
        date_layout.addWidget(self.date_picker)

        self.count_lbl = QLabel()
        self.count_lbl.setStyleSheet("font-weight: bold; color: #ff5722; font-size: 14px;")
        date_layout.addWidget(self.count_lbl)
        layout.addWidget(date_group)

        # Export Path
        path_group = QGroupBox("📁 مسار حفظ الملف المصدر")
        path_layout = QHBoxLayout(path_group)
        self.path_edit = QLineEdit()
        self.update_default_path()
        browse_btn = QPushButton("استعراض...")
        browse_btn.clicked.connect(self.browse_output)
        path_layout.addWidget(self.path_edit)
        path_layout.addWidget(browse_btn)
        layout.addWidget(path_group)

        # Action Buttons
        btn_layout = QHBoxLayout()
        self.export_btn = QPushButton("🚀 تصدير الشيت الآن")
        self.export_btn.setStyleSheet("background-color: #16a34a; color: white; font-weight: bold; padding: 10px 20px; font-size: 14px;")
        self.export_btn.clicked.connect(self.do_export)

        cancel_btn = QPushButton("إلغاء")
        cancel_btn.clicked.connect(self.reject)

        btn_layout.addWidget(self.export_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        self.update_count()

    def update_default_path(self):
        d_str = self.date_picker.date().toString("yyyy-MM-dd")
        default_dir = os.path.join(APP_DIR, "exports")
        os.makedirs(default_dir, exist_ok=True)
        self.path_edit.setText(os.path.join(default_dir, f"Bosta_Orders_{d_str}.xlsx"))

    def update_count(self):
        self.update_default_path()
        d_str = self.date_picker.date().toString("yyyy-MM-dd")
        matching = self.excel_manager.get_orders_by_date(d_str)
        total_cod = sum(float(o.get("cod") or 0) for o in matching)
        self.count_lbl.setText(f"📦 تم العثور على {len(matching)} طلب بتاريخ {d_str} (إجمالي التحصيل: {total_cod:,.2f} ج.م)")
        self.export_btn.setEnabled(len(matching) > 0)

    def browse_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "حفظ شيت بوسطة اليومي", self.path_edit.text(), "Excel Files (*.xlsx)"
        )
        if path:
            self.path_edit.setText(path)

    def do_export(self):
        d_str = self.date_picker.date().toString("yyyy-MM-dd")
        out_path = self.path_edit.text().strip()
        ok, msg, count = self.excel_manager.export_orders_by_date(d_str, out_path)
        if ok:
            box = QMessageBox(self)
            box.setWindowTitle("تم التصدير بنجاح")
            box.setText(msg)
            box.setIcon(QMessageBox.Icon.Information)
            open_file_btn = box.addButton("📂 فتح الملف", QMessageBox.ButtonRole.ActionRole)
            open_folder_btn = box.addButton("📁 فتح المجلد", QMessageBox.ButtonRole.ActionRole)
            box.addButton("إغلاق", QMessageBox.ButtonRole.RejectRole)
            box.exec()

            if box.clickedButton() == open_file_btn:
                try:
                    subprocess.Popen(["xdg-open", out_path])
                except Exception:
                    pass
            elif box.clickedButton() == open_folder_btn:
                try:
                    subprocess.Popen(["xdg-open", os.path.dirname(out_path)])
                except Exception:
                    pass
            self.accept()
        else:
            QMessageBox.critical(self, "فشل التصدير", msg)


class SettingsDialog(QDialog):
    def __init__(self, parent=None, config=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ الإعدادات وملف الإكسيل")
        self.setMinimumWidth(500)
        self.config = config or {}
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        excel_group = QGroupBox("📁 مسار ملف إكسيل بوسطة")
        eg_layout = QHBoxLayout(excel_group)
        self.excel_edit = QLineEdit(self.config.get("excel_path", ""))
        browse_btn = QPushButton("استعراض...")
        browse_btn.clicked.connect(self.browse_file)
        eg_layout.addWidget(self.excel_edit)
        eg_layout.addWidget(browse_btn)
        layout.addWidget(excel_group)

        ai_group = QGroupBox("✨ مفتاح الذكاء الاصطناعي (Gemini API Key)")
        ai_layout = QVBoxLayout(ai_group)
        self.api_key_edit = QLineEdit(self.config.get("gemini_api_key", ""))
        self.api_key_edit.setEchoMode(QLineEdit.EchoMode.PasswordEchoOnEdit)
        self.api_key_edit.setPlaceholderText("أدخل مفتاح Gemini API هنا (اختياري للتحليل بالذكاء الاصطناعي)")
        ai_hint = QLabel("💡 ملاحظة: المحلل المدمج فائق السرعة يعمل محلياً بدون إنترنت ومجاناً 100%. مفتاح الذكاء الاصطناعي اختياري.")
        ai_hint.setStyleSheet("color: #64748b; font-size: 11px;")
        ai_hint.setWordWrap(True)
        ai_layout.addWidget(self.api_key_edit)
        ai_layout.addWidget(ai_hint)
        layout.addWidget(ai_group)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "اختر ملف إكسيل بوسطة", self.excel_edit.text(), "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_edit.setText(file_path)

    def get_data(self):
        return {
            "excel_path": self.excel_edit.text().strip(),
            "gemini_api_key": self.api_key_edit.text().strip()
        }


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = load_config()
        self.excel_manager = ExcelManager(self.config.get("excel_path"))
        
        self.setWindowTitle("🚀 مدير أوردرات بوسطة الذكي - Bosta Smart Order Manager")
        self.resize(1200, 850)
        self.setMinimumSize(1000, 700)

        self.apply_theme()
        self.setup_ui()
        self.refresh_stats_and_table()

    def apply_theme(self):
        self.setStyleSheet("""
            QWidget {
                font-family: 'Segoe UI', 'Cairo', 'Arial', sans-serif;
                font-size: 13px;
                color: #1e293b;
            }
            QMainWindow {
                background-color: #f8fafc;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 15px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top right;
                padding: 0 10px;
                color: #0f172a;
            }
            QLineEdit, QTextEdit, QComboBox, QSpinBox, QDateEdit {
                border: 1.5px solid #cbd5e1;
                border-radius: 6px;
                padding: 7px 10px;
                background-color: #ffffff;
                selection-background-color: #ff5722;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QSpinBox:focus, QDateEdit:focus {
                border: 1.5px solid #ff5722;
                background-color: #fffbf9;
            }
            QPushButton {
                background-color: #0f172a;
                color: #ffffff;
                font-weight: bold;
                border-radius: 6px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #1e293b;
            }
            QPushButton:pressed {
                background-color: #334155;
            }
            QPushButton#btn_parse {
                background-color: #ff5722;
                color: #ffffff;
                font-size: 13px;
            }
            QPushButton#btn_parse:hover {
                background-color: #e64a19;
            }
            QPushButton#btn_ai {
                background-color: #7c3aed;
                color: #ffffff;
                font-size: 13px;
            }
            QPushButton#btn_ai:hover {
                background-color: #6d28d9;
            }
            QPushButton#btn_add {
                background-color: #16a34a;
                color: #ffffff;
                font-size: 15px;
                font-weight: bold;
                padding: 12px;
                border-radius: 8px;
            }
            QPushButton#btn_add:hover {
                background-color: #15803d;
            }
            QPushButton#btn_clear {
                background-color: #e2e8f0;
                color: #475569;
            }
            QPushButton#btn_clear:hover {
                background-color: #cbd5e1;
            }
            QPushButton#btn_export {
                background-color: #0284c7;
                color: #ffffff;
            }
            QPushButton#btn_export:hover {
                background-color: #0369a1;
            }
            QPushButton#btn_delete {
                background-color: #dc2626;
                color: #ffffff;
            }
            QPushButton#btn_delete:hover {
                background-color: #b91c1c;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                gridline-color: #f1f5f9;
                selection-background-color: #ffedd5;
                selection-color: #9a3412;
            }
            QHeaderView::section {
                background-color: #f1f5f9;
                padding: 6px;
                font-weight: bold;
                border: none;
                border-bottom: 2px solid #cbd5e1;
            }
            QStatusBar {
                background-color: #ffffff;
                border-top: 1px solid #e2e8f0;
            }
        """)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 12, 15, 12)
        main_layout.setSpacing(10)

        # 1. Top Header Bar
        header_bar = QFrame()
        header_bar.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1e293b, stop:1 #0f172a);
                border-radius: 10px;
                padding: 8px;
            }
            QLabel {
                color: #ffffff;
            }
        """)
        header_layout = QHBoxLayout(header_bar)
        header_layout.setContentsMargins(15, 8, 15, 8)

        title_label = QLabel("🚀 مدير أوردرات بوسطة الذكي")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #ffedd5;")
        header_layout.addWidget(title_label)

        header_layout.addStretch()

        self.orders_count_badge = QLabel("📊 إجمالي الطلبات: 0")
        self.orders_count_badge.setStyleSheet("background-color: #334155; padding: 6px 12px; border-radius: 6px; font-weight: bold;")
        header_layout.addWidget(self.orders_count_badge)

        self.next_code_badge = QLabel("🔢 الطلب القادم: ORD-1")
        self.next_code_badge.setStyleSheet("background-color: #ff5722; padding: 6px 12px; border-radius: 6px; font-weight: bold; color: #ffffff;")
        header_layout.addWidget(self.next_code_badge)

        # Export Button in Header
        export_btn = QPushButton("📅 تصدير ليوم محدد")
        export_btn.setObjectName("btn_export")
        export_btn.clicked.connect(self.open_export_dialog)
        header_layout.addWidget(export_btn)

        settings_btn = QPushButton("⚙️ إعدادات")
        settings_btn.setStyleSheet("background-color: #475569; padding: 6px 12px;")
        settings_btn.clicked.connect(self.open_settings)
        header_layout.addWidget(settings_btn)

        open_excel_btn = QPushButton("📂 فتح الإكسيل")
        open_excel_btn.setStyleSheet("background-color: #0284c7; padding: 6px 12px;")
        open_excel_btn.clicked.connect(self.open_excel_externally)
        header_layout.addWidget(open_excel_btn)

        main_layout.addWidget(header_bar)

        # 2. Main Work Area (Splitter: Left = Paste Box & Alerts, Right = Form Fields)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        # === LEFT PANEL: Paste Message & Validation ===
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 8, 0)
        left_layout.setSpacing(10)

        msg_group = QGroupBox("📩 لصق رسالة العميل (واتساب / فيسبوك / نص مباشر)")
        msg_layout = QVBoxLayout(msg_group)

        self.msg_text = QTextEdit()
        self.msg_text.setPlaceholderText(
            "الصق رسالة العميل هنا بأي صيغة...\n"
            "مثال:\n"
            "الاسم: أحمد مصطفى\n"
            "الرقم: 01012345678\n"
            "المحافظة: القاهرة - المعادي\n"
            "العنوان: شارع 9 عمارة 4\n"
            "السعر: 450 ج\n"
            "الاوردر: 2 كوتشي ابيض"
        )
        self.msg_text.setMinimumHeight(180)
        
        self.debounce_timer = QTimer()
        self.debounce_timer.setSingleShot(True)
        self.debounce_timer.setInterval(400)
        self.debounce_timer.timeout.connect(self.auto_parse_triggered)
        self.msg_text.textChanged.connect(lambda: self.debounce_timer.start())

        msg_layout.addWidget(self.msg_text)

        btn_bar = QHBoxLayout()
        self.btn_parse = QPushButton("⚡ تحليل سريع للرسالة")
        self.btn_parse.setObjectName("btn_parse")
        self.btn_parse.clicked.connect(self.parse_message_local)
        
        self.btn_ai = QPushButton("✨ تحليل بالذكاء الاصطناعي")
        self.btn_ai.setObjectName("btn_ai")
        self.btn_ai.clicked.connect(self.parse_message_ai)

        self.btn_clear_msg = QPushButton("🧹 مسح")
        self.btn_clear_msg.setObjectName("btn_clear")
        self.btn_clear_msg.clicked.connect(self.clear_all)

        btn_bar.addWidget(self.btn_parse)
        btn_bar.addWidget(self.btn_ai)
        btn_bar.addWidget(self.btn_clear_msg)
        msg_layout.addLayout(btn_bar)

        left_layout.addWidget(msg_group)

        # Validation & Alerts Box
        self.alert_group = QGroupBox("🔍 الفحص الذكي وكشف التكرار والأخطاء")
        alert_layout = QVBoxLayout(self.alert_group)

        self.alert_box = QLabel("جاهز... قم بلصق رسالة العميل لاستخراج البيانات وفحصها تلقائياً.")
        self.alert_box.setStyleSheet("""
            QLabel {
                background-color: #f1f5f9;
                border: 1.5px solid #cbd5e1;
                border-radius: 6px;
                padding: 10px;
                color: #475569;
                font-size: 13px;
            }
        """)
        self.alert_box.setWordWrap(True)
        self.alert_box.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.alert_box.setMinimumHeight(130)

        alert_layout.addWidget(self.alert_box)
        left_layout.addWidget(self.alert_group)

        splitter.addWidget(left_panel)

        # === RIGHT PANEL: Form Fields (Quick Editable Form) ===
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(10)

        form_group = QGroupBox("📋 بيانات الطلب المستخرجة (نظام الحقول السريع)")
        form_layout = QGridLayout(form_group)
        form_layout.setVerticalSpacing(8)
        form_layout.setHorizontalSpacing(10)

        # Row 0: Name & Phone
        form_layout.addWidget(QLabel("اسم العميل *:"), 0, 0)
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("اسم العميل كامل")
        self.edit_name.textChanged.connect(self.validate_current_form)
        form_layout.addWidget(self.edit_name, 0, 1)

        phone_header_layout = QHBoxLayout()
        phone_lbl = QLabel("رقم الهاتف *:")
        self.phone_counter_lbl = QLabel("0/11")
        self.phone_counter_lbl.setStyleSheet("font-size: 11px; font-weight: bold; color: #94a3b8;")
        phone_header_layout.addWidget(phone_lbl)
        phone_header_layout.addStretch()
        phone_header_layout.addWidget(self.phone_counter_lbl)
        form_layout.addLayout(phone_header_layout, 0, 2)

        self.edit_phone = QLineEdit()
        self.edit_phone.setPlaceholderText("مثال: 01012345678")
        self.edit_phone.textChanged.connect(self.on_phone_changed)
        form_layout.addWidget(self.edit_phone, 0, 3)

        # Row 1: Second Phone & City
        form_layout.addWidget(QLabel("هاتف ثاني:"), 1, 0)
        self.edit_phone2 = QLineEdit()
        self.edit_phone2.setPlaceholderText("اختياري (موبايل أو أرضي)")
        self.edit_phone2.textChanged.connect(self.validate_current_form)
        form_layout.addWidget(self.edit_phone2, 1, 1)

        form_layout.addWidget(QLabel("المحافظة *:"), 1, 2)
        self.combo_city = QComboBox()
        self.combo_city.addItem("-- اختر المحافظة --", "")
        for c in BOSTA_CITIES:
            self.combo_city.addItem(c, c)
        self.combo_city.currentTextChanged.connect(self.on_city_changed)
        form_layout.addWidget(self.combo_city, 1, 3)

        # Row 2: Area & Detailed Address
        form_layout.addWidget(QLabel("المنطقة:"), 2, 0)
        self.combo_area = QComboBox()
        self.combo_area.setEditable(True)
        self.combo_area.setPlaceholderText("اختر أو اكتب المنطقة")
        form_layout.addWidget(self.combo_area, 2, 1)

        form_layout.addWidget(QLabel("العنوان التفصيلي *:"), 2, 2)
        self.edit_address = QLineEdit()
        self.edit_address.setPlaceholderText("الشارع، رقم العمارة، الشقة، إلخ")
        self.edit_address.textChanged.connect(self.validate_current_form)
        form_layout.addWidget(self.edit_address, 2, 3)

        # Row 3: Landmark & Cash COD
        form_layout.addWidget(QLabel("علامة مميزة:"), 3, 0)
        self.edit_landmark = QLineEdit()
        self.edit_landmark.setPlaceholderText("بجوار مسجد، أمام مدرسة...")
        form_layout.addWidget(self.edit_landmark, 3, 1)

        form_layout.addWidget(QLabel("التحصيل النقدي (ج.م):"), 3, 2)
        self.edit_cod = QLineEdit()
        self.edit_cod.setPlaceholderText("المبلغ المطلوب تحصيله (COD)")
        self.edit_cod.textChanged.connect(self.validate_current_form)
        form_layout.addWidget(self.edit_cod, 3, 3)

        # Row 4: Items Count & Package Description
        form_layout.addWidget(QLabel("عدد القطع:"), 4, 0)
        self.spin_items = QSpinBox()
        self.spin_items.setRange(1, 999)
        self.spin_items.setValue(1)
        form_layout.addWidget(self.spin_items, 4, 1)

        form_layout.addWidget(QLabel("وصف الشحنة:"), 4, 2)
        self.edit_desc = QLineEdit()
        self.edit_desc.setPlaceholderText("نوع المنتج / المقاس / اللون")
        form_layout.addWidget(self.edit_desc, 4, 3)

        # Row 5: Notes & Order Date
        notes_header = QHBoxLayout()
        notes_header.addWidget(QLabel("كود الطلب / ملاحظات:"))
        auto_ord_btn = QPushButton("توليد كود 🔢")
        auto_ord_btn.setStyleSheet("background-color: #e2e8f0; color: #1e293b; padding: 2px 6px; font-size: 11px;")
        auto_ord_btn.clicked.connect(lambda: self.edit_notes.setText(self.excel_manager.get_next_order_code()))
        notes_header.addWidget(auto_ord_btn)
        form_layout.addLayout(notes_header, 5, 0)

        self.edit_notes = QLineEdit()
        self.edit_notes.setPlaceholderText("كود الطلب ORD-XXX أو ملاحظات إضافية")
        form_layout.addWidget(self.edit_notes, 5, 1)

        form_layout.addWidget(QLabel("تاريخ الطلب:"), 5, 2)
        self.order_date_edit = QDateEdit()
        self.order_date_edit.setCalendarPopup(True)
        self.order_date_edit.setDate(QDate.currentDate())
        self.order_date_edit.setDisplayFormat("yyyy-MM-dd")
        form_layout.addWidget(self.order_date_edit, 5, 3)

        right_layout.addWidget(form_group)

        # Action Button: ADD ORDER
        self.btn_add_order = QPushButton("💾 إضافة الطلب إلى ملف الإكسيل فوراً (Ctrl + Enter)")
        self.btn_add_order.setObjectName("btn_add")
        self.btn_add_order.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_add_order.clicked.connect(self.add_order_to_excel)
        right_layout.addWidget(self.btn_add_order)

        splitter.addWidget(right_panel)

        splitter.setSizes([500, 700])
        main_layout.addWidget(splitter, 2)

        # 3. Bottom Table: Recent Orders in Sheet
        table_group = QGroupBox("📋 سجل الطلبات المسجلة في شيت بوسطة")
        table_layout = QVBoxLayout(table_group)
        table_layout.setContentsMargins(10, 10, 10, 10)

        # Action bar above table
        action_bar = QHBoxLayout()
        
        action_bar.addWidget(QLabel("🔍 بحث:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("ابحث بالاسم أو رقم التليفون أو كود الطلب أو التاريخ...")
        self.search_input.textChanged.connect(self.filter_orders_table)
        action_bar.addWidget(self.search_input)

        # Action buttons
        export_orders_btn = QPushButton("📅 تصدير شيت ليوم محدد")
        export_orders_btn.setObjectName("btn_export")
        export_orders_btn.clicked.connect(self.open_export_dialog)
        action_bar.addWidget(export_orders_btn)

        self.delete_btn = QPushButton("🗑 حذف الطلب المحدد")
        self.delete_btn.setObjectName("btn_delete")
        self.delete_btn.clicked.connect(self.delete_selected_order)
        action_bar.addWidget(self.delete_btn)

        clear_all_btn = QPushButton("⚠️ تفريغ كافة البيانات")
        clear_all_btn.setStyleSheet("background-color: #ef4444; color: white;")
        clear_all_btn.clicked.connect(self.clear_all_orders_data)
        action_bar.addWidget(clear_all_btn)

        refresh_btn = QPushButton("🔄 تحديث")
        refresh_btn.setStyleSheet("background-color: #475569; padding: 5px 12px;")
        refresh_btn.clicked.connect(self.reload_excel_file)
        action_bar.addWidget(refresh_btn)

        table_layout.addLayout(action_bar)

        self.orders_table = QTableWidget()
        self.orders_table.setColumnCount(10)
        self.orders_table.setHorizontalHeaderLabels([
            "السطر", "كود الطلب", "التاريخ", "اسم العميل", "رقم الهاتف", "المحافظة", "المنطقة", "العنوان", "المبلغ", "القطع"
        ])
        self.orders_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.orders_table.horizontalHeader().setStretchLastSection(True)
        self.orders_table.setAlternatingRowColors(True)
        self.orders_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.orders_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.orders_table.customContextMenuRequested.connect(self.show_table_context_menu)
        self.orders_table.setMaximumHeight(200)

        table_layout.addWidget(self.orders_table)
        main_layout.addWidget(table_group, 1)

        # Shortcuts
        self.btn_add_order.setShortcut("Ctrl+Return")
        
        # Status Bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("مرحباً بك في مدير أوردرات بوسطة الذكي")

    def auto_parse_triggered(self):
        text = self.msg_text.toPlainText().strip()
        if text:
            if not self.edit_name.text() and not self.edit_phone.text():
                self.parse_message_local()

    def parse_message_local(self):
        text = self.msg_text.toPlainText().strip()
        if not text:
            self.show_alert("info", "يرجى لصق رسالة العميل أولاً في المربع أعلاه.")
            return

        parsed = parse_order_text(text)
        self.fill_form(parsed)
        self.validate_current_form()

    def parse_message_ai(self):
        text = self.msg_text.toPlainText().strip()
        if not text:
            self.show_alert("info", "يرجى لصق رسالة العميل أولاً.")
            return

        api_key = self.config.get("gemini_api_key", "")
        if not api_key:
            reply = QMessageBox.question(
                self, "مفتاح Gemini API غير متوفر",
                "لم يتم إدخال مفتاح Gemini API في الإعدادات.\nهل ترغب في فتح نافذة الإعدادات لإدخال المفتاح الآن؟\n(ملاحظة: يمكنك استخدام 'التحليل السريع' مجاناً وبدون مفتاح).",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.open_settings()
            return

        self.btn_ai.setEnabled(False)
        self.btn_ai.setText("جاري التحليل بالذكاء الاصطناعي...")
        QApplication.processEvents()

        try:
            parsed = parse_order_with_gemini(text, api_key)
            self.fill_form(parsed)
            self.validate_current_form()
        except Exception as e:
            QMessageBox.warning(self, "خطأ بالذكاء الاصطناعي", f"حدث خطأ أثناء التحليل: {str(e)}")
        finally:
            self.btn_ai.setEnabled(True)
            self.btn_ai.setText("✨ تحليل بالذكاء الاصطناعي")

    def fill_form(self, data):
        self.edit_name.setText(data.get("name", ""))
        self.edit_phone.setText(data.get("phone", ""))
        self.edit_phone2.setText(data.get("second_phone", ""))
        
        city = data.get("city", "")
        city_idx = self.combo_city.findText(city)
        if city_idx >= 0:
            self.combo_city.setCurrentIndex(city_idx)
        else:
            self.combo_city.setCurrentIndex(0)

        area = data.get("area", "")
        if area:
            self.combo_area.setEditText(area)

        self.edit_address.setText(data.get("address", ""))
        self.edit_landmark.setText(data.get("landmark", ""))
        self.edit_cod.setText(str(data.get("cod_amount", "")))
        
        try:
            items_num = int(data.get("items_count", 1))
            self.spin_items.setValue(max(1, items_num))
        except ValueError:
            self.spin_items.setValue(1)

        self.edit_desc.setText(data.get("description", ""))
        
        notes = data.get("notes", "")
        if not notes:
            notes = self.excel_manager.get_next_order_code()
        self.edit_notes.setText(notes)

        self.order_date_edit.setDate(QDate.currentDate())

    def on_city_changed(self, city_name):
        self.combo_area.clear()
        if city_name and city_name in BOSTA_CITY_AREAS:
            areas = BOSTA_CITY_AREAS[city_name]
            self.combo_area.addItems(areas)
        self.validate_current_form()

    def on_phone_changed(self, text):
        digits = re.sub(r'[^\d]', '', text)
        count = len(digits)
        
        if count == 11 and digits.startswith(('010', '011', '012', '015')):
            self.phone_counter_lbl.setText("11/11 ✅ ممتاز")
            self.phone_counter_lbl.setStyleSheet("font-size: 11px; font-weight: bold; color: #16a34a;")
            self.edit_phone.setStyleSheet("border: 1.5px solid #16a34a; background-color: #f0fdf4;")
        elif count == 10 and digits.startswith(('10', '11', '12', '15')):
            self.phone_counter_lbl.setText("10/11 ⚠️ ينقصه 0")
            self.phone_counter_lbl.setStyleSheet("font-size: 11px; font-weight: bold; color: #d97706;")
            self.edit_phone.setStyleSheet("border: 1.5px solid #d97706; background-color: #fffbeb;")
        else:
            self.phone_counter_lbl.setText(f"{count}/11 ❌ ناقص")
            self.phone_counter_lbl.setStyleSheet("font-size: 11px; font-weight: bold; color: #dc2626;")
            self.edit_phone.setStyleSheet("border: 1.5px solid #dc2626; background-color: #fef2f2;")

        self.validate_current_form()

    def get_form_data(self):
        return {
            "name": self.edit_name.text().strip(),
            "phone": self.edit_phone.text().strip(),
            "second_phone": self.edit_phone2.text().strip(),
            "city": self.combo_city.currentText().strip() if self.combo_city.currentIndex() > 0 else "",
            "area": self.combo_area.currentText().strip(),
            "address": self.edit_address.text().strip(),
            "landmark": self.edit_landmark.text().strip(),
            "cod_amount": self.edit_cod.text().strip(),
            "items_count": self.spin_items.value(),
            "description": self.edit_desc.text().strip(),
            "notes": self.edit_notes.text().strip(),
            "date": self.order_date_edit.date().toString("yyyy-MM-dd")
        }

    def validate_current_form(self):
        data = self.get_form_data()
        
        if not data["name"] and not data["phone"] and not data["address"]:
            self.show_alert("info", "جاهز... قم بلصق رسالة العميل لاستخراج البيانات وفحصها تلقائياً.")
            return True

        check = self.excel_manager.check_duplicates_and_validations(data)

        msg_parts = []

        if check["has_duplicates"]:
            for dup in check["duplicates"]:
                msg_parts.append(f"⚠️ <b>تحذير تكرار:</b> {dup['message']}")

        if check["errors"]:
            for err in check["errors"]:
                msg_parts.append(f"❌ <b>خطأ:</b> {err}")

        if check["warnings"]:
            for w in check["warnings"]:
                msg_parts.append(f"💡 <b>تنبيه:</b> {w}")

        if check["errors"]:
            self.show_alert("error", "<br>".join(msg_parts))
            return False
        elif check["has_duplicates"]:
            msg_parts.append("<br>👉 <b>يمكنك الضغط على 'إضافة الطلب' للمتابعة رغم التكرار إذا كان العميل يطلب مجدداً.</b>")
            self.show_alert("warning", "<br>".join(msg_parts))
            return True
        else:
            success_msg = "✅ <b>البيانات مكتملة وصحيحة ورقم الهاتف سليم 100%!</b><br>جاهز للإضافة إلى ملف الإكسيل الآن."
            if check["warnings"]:
                success_msg += "<br>" + "<br>".join(f"💡 {w}" for w in check["warnings"])
            self.show_alert("success", success_msg)
            return True

    def show_alert(self, alert_type, html_content):
        styles = {
            "info": """
                background-color: #f8fafc;
                border: 1.5px solid #cbd5e1;
                border-radius: 8px;
                padding: 10px;
                color: #475569;
                font-size: 13px;
            """,
            "success": """
                background-color: #f0fdf4;
                border: 1.5px solid #22c55e;
                border-radius: 8px;
                padding: 10px;
                color: #15803d;
                font-size: 13px;
            """,
            "warning": """
                background-color: #fffbeb;
                border: 1.5px solid #f59e0b;
                border-radius: 8px;
                padding: 10px;
                color: #b45309;
                font-size: 13px;
            """,
            "error": """
                background-color: #fef2f2;
                border: 1.5px solid #ef4444;
                border-radius: 8px;
                padding: 10px;
                color: #b91c1c;
                font-size: 13px;
            """
        }
        self.alert_box.setStyleSheet(styles.get(alert_type, styles["info"]))
        self.alert_box.setText(html_content)

    def add_order_to_excel(self):
        data = self.get_form_data()
        
        clean_p = re.sub(r'[^\d]', '', data["phone"])
        if len(clean_p) == 10 and clean_p.startswith(('10', '11', '12', '15')):
            data["phone"] = '0' + clean_p
            self.edit_phone.setText(data["phone"])

        check = self.excel_manager.check_duplicates_and_validations(data)

        if not check["is_valid"]:
            error_list = "\n• " + "\n• ".join(check["errors"])
            QMessageBox.critical(
                self, "لا يمكن إضافة الطلب - توجد أخطاء",
                f"يرجى تصحيح الأخطاء التالية أولاً قبل الإضافة:\n{error_list}"
            )
            return

        if check["has_duplicates"]:
            dup_details = "\n• " + "\n• ".join([d["message"] for d in check["duplicates"]])
            reply = QMessageBox.question(
                self, "⚠️ تم كشف بيانات متكررة!",
                f"تم العثور على أوردر سابق مسجل بنفس البيانات:\n{dup_details}\n\nهل أنت متأكد من رغبتك في إضافة هذا الطلب كطلب جديد؟",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        ok, msg, new_row = self.excel_manager.append_order(data)
        if ok:
            self.status_bar.showMessage(msg, 6000)
            self.show_alert("success", f"🎉 <b>تم حفظ الطلب بنجاح في السطر {new_row}!</b><br>كود الطلب: {data.get('notes')}")
            self.clear_for_next_order()
            self.refresh_stats_and_table()
            QMessageBox.information(self, "تم الحفظ بنجاح", msg)
        else:
            QMessageBox.critical(self, "فشل الحفظ", msg)

    def delete_selected_order(self):
        selected = self.orders_table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "تحديد الطلب", "يرجى تحديد السطر المراد حذفه من الجدول أولاً.")
            return

        row_idx = self.orders_table.currentRow()
        row_num_item = self.orders_table.item(row_idx, 0)
        code_item = self.orders_table.item(row_idx, 1)
        name_item = self.orders_table.item(row_idx, 3)

        if not row_num_item:
            return

        row_num = int(row_num_item.text())
        code_str = code_item.text() if code_item else ""
        name_str = name_item.text() if name_item else ""

        reply = QMessageBox.question(
            self, "تأكيد الحذف",
            f"هل أنت متأكد تماماً من حذف الطلب '{code_str}' الخاص بالعميل '{name_str}' (السطر {row_num}) من ملف الإكسيل؟",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            ok, msg = self.excel_manager.delete_order(row_num)
            if ok:
                self.refresh_stats_and_table()
                self.status_bar.showMessage(msg, 5000)
                QMessageBox.information(self, "تم الحذف", msg)
            else:
                QMessageBox.critical(self, "فشل الحذف", msg)

    def clear_all_orders_data(self):
        reply = QMessageBox.question(
            self, "⚠️ تحذير: تفريغ كافة البيانات",
            "هل أنت متأكد من رغبتك في تفريغ وحذف كافة بيانات الأوردرات من الشيت والبدء بشيت فارغ تماماً؟\n\n(سيتم الحفاظ على قالب بوسطة والعناوين الرئيسية).",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            ok, msg = self.excel_manager.clear_all_orders()
            if ok:
                self.refresh_stats_and_table()
                self.status_bar.showMessage(msg, 5000)
                QMessageBox.information(self, "تم التفريغ", msg)
            else:
                QMessageBox.critical(self, "فشل التفريغ", msg)

    def show_table_context_menu(self, pos):
        menu = QMenu(self)
        del_action = QAction("🗑 حذف هذا الطلب", self)
        del_action.triggered.connect(self.delete_selected_order)
        menu.addAction(del_action)
        menu.exec(self.orders_table.viewport().mapToGlobal(pos))

    def open_export_dialog(self):
        dlg = ExportDialog(self, self.excel_manager)
        dlg.exec()

    def clear_for_next_order(self):
        self.msg_text.clear()
        self.edit_name.clear()
        self.edit_phone.clear()
        self.edit_phone2.clear()
        self.combo_city.setCurrentIndex(0)
        self.combo_area.clear()
        self.edit_address.clear()
        self.edit_landmark.clear()
        self.edit_cod.clear()
        self.spin_items.setValue(1)
        self.edit_desc.clear()
        self.edit_notes.setText(self.excel_manager.get_next_order_code())
        self.edit_phone.setStyleSheet("")
        self.phone_counter_lbl.setText("0/11")
        self.order_date_edit.setDate(QDate.currentDate())

    def clear_all(self):
        self.clear_for_next_order()
        self.show_alert("info", "جاهز... قم بلصق رسالة العميل لاستخراج البيانات وفحصها تلقائياً.")

    def refresh_stats_and_table(self):
        count = len(self.excel_manager.orders)
        next_code = self.excel_manager.get_next_order_code()
        
        self.orders_count_badge.setText(f"📊 إجمالي الطلبات: {count}")
        self.next_code_badge.setText(f"🔢 الطلب القادم: {next_code}")
        
        if not self.edit_notes.text():
            self.edit_notes.setText(next_code)

        self.populate_table(self.excel_manager.orders)

    def populate_table(self, orders_list):
        self.orders_table.setRowCount(0)
        display_orders = list(reversed(orders_list))

        for row_idx, ord_obj in enumerate(display_orders):
            self.orders_table.insertRow(row_idx)

            items = [
                str(ord_obj.get("row", "")),
                str(ord_obj.get("notes", "")),
                str(ord_obj.get("date", "")),
                str(ord_obj.get("name", "")),
                str(ord_obj.get("phone", "")),
                str(ord_obj.get("city", "")),
                str(ord_obj.get("area", "")),
                str(ord_obj.get("address", "")),
                str(ord_obj.get("cod", "") or "-"),
                str(ord_obj.get("items", "") or "1")
            ]

            for col_idx, val in enumerate(items):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.orders_table.setItem(row_idx, col_idx, item)

    def filter_orders_table(self, query):
        q = query.strip().lower()
        if not q:
            self.populate_table(self.excel_manager.orders)
            return

        filtered = []
        for o in self.excel_manager.orders:
            if (q in o.get("name", "").lower() or
                q in o.get("phone", "").lower() or
                q in o.get("notes", "").lower() or
                q in o.get("date", "").lower() or
                q in o.get("city", "").lower() or
                q in o.get("address", "").lower()):
                filtered.append(o)

        self.populate_table(filtered)

    def reload_excel_file(self):
        self.excel_manager.reload()
        self.refresh_stats_and_table()
        self.status_bar.showMessage("تم تحديث بيانات ملف الإكسيل بنجاح.", 4000)

    def open_excel_externally(self):
        path = self.excel_manager.file_path
        if os.path.exists(path):
            try:
                subprocess.Popen(["xdg-open", path])
            except Exception as e:
                QMessageBox.warning(self, "خطأ في الفتح", f"تعذر فتح الملف تلقائياً: {str(e)}")
        else:
            QMessageBox.warning(self, "الملف غير موجود", f"لم يتم العثور على الملف في:\n{path}")

    def open_settings(self):
        dlg = SettingsDialog(self, self.config)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            new_data = dlg.get_data()
            self.config.update(new_data)
            save_config(self.config)

            if new_data["excel_path"] != self.excel_manager.file_path:
                self.excel_manager = ExcelManager(new_data["excel_path"])
                self.refresh_stats_and_table()

            self.status_bar.showMessage("تم حفظ الإعدادات بنجاح.", 4000)


def main():
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

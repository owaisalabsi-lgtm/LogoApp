import json
import os
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# مكتبة الإكسيل
try:
  import openpyxl

  HAS_EXCEL = True
except ImportError:
  HAS_EXCEL = False

# مكتبات ה-PDF العربية (ReportLab + Arabic Reshaper)
try:
  import arabic_reshaper
  from bidi.algorithm import get_display
  from reportlab.lib import colors
  from reportlab.lib.pagesizes import A4, portrait
  from reportlab.pdfgen import canvas
  from reportlab.pdfbase import pdfmetrics
  from reportlab.pdfbase.ttfonts import TTFont
  from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
  from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

  HAS_PDF = True
except ImportError:
  HAS_PDF = False


def fix_arabic(text):
  """معالجة النص العربي وتحويل اتجاهه ليظهر بشكل صحيح"""
  if not text:
    return ""
  reshaped = arabic_reshaper.reshape(str(text))
  return get_display(reshaped)


class AccountingApp:

  def __init__(self, root):
    self.root = root
    self.root.title(
        "نظام الحسابات والمطالبات المالية المتكامل - مكتب الحمدي تكنولوجي"
    )
    self.root.geometry("950x680")

    self.data_file = "accounts_data.json"
    self.trash_file = "trash_bin.json"

    self.accounts_data = self.load_data(self.data_file)
    self.trash_data = self.load_data(self.trash_file)

    self.cleanup_trash()

    self.current_account = "الحساب العام"
    if self.current_account not in self.accounts_data:
      self.accounts_data[self.current_account] = []
      self.save_data(self.data_file, self.accounts_data)

    self.selected_item_index = None
    self.setup_ui()
    self.refresh_table()

  def load_data(self, file_path):
    if os.path.exists(file_path):
      try:
        with open(file_path, "r", encoding="utf-8") as f:
          return json.load(f)
      except:
        return {}
    return {}

  def save_data(self, file_path, data):
    with open(file_path, "w", encoding="utf-8") as f:
      json.dump(data, f, ensure_ascii=False, indent=4)

  def cleanup_trash(self):
    now = datetime.now()
    updated_trash = {}
    changed = False

    for acc_name, acc_info in self.trash_data.items():
      deleted_date_str = acc_info.get("deleted_date")
      if deleted_date_str:
        deleted_date = datetime.strptime(deleted_date_str, "%Y-%m-%d %H:%M:%S")
        if now - deleted_date <= timedelta(days=30):
          updated_trash[acc_name] = acc_info
        else:
          changed = True
      else:
        updated_trash[acc_name] = acc_info

    if changed:
      self.trash_data = updated_trash
      self.save_data(self.trash_file, self.trash_data)

  def setup_ui(self):
    # 1. إطار إدارة الحسابات
    acc_frame = tk.LabelFrame(
        self.root, text=" إدارة الحسابات ", font=("Tahoma", 10, "bold")
    )
    acc_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(acc_frame, text="اختر الحساب:", font=("Tahoma", 10)).pack(
        side="right", padx=5
    )
    self.acc_combo = ttk.Combobox(
        acc_frame,
        values=list(self.accounts_data.keys()),
        state="readonly",
        font=("Tahoma", 10),
        width=20,
    )
    self.acc_combo.set(self.current_account)
    self.acc_combo.pack(side="right", padx=5)
    self.acc_combo.bind("<<ComboboxSelected>>", self.on_account_change)

    btn_new_acc = tk.Button(
        acc_frame,
        text="+ حساب جديد",
        bg="#2980b9",
        fg="white",
        font=("Tahoma", 9, "bold"),
        command=self.add_new_account,
    )
    btn_new_acc.pack(side="right", padx=5)

    btn_del_acc = tk.Button(
        acc_frame,
        text="حذف الحساب الحالي",
        bg="#c0392b",
        fg="white",
        font=("Tahoma", 9, "bold"),
        command=self.delete_current_account,
    )
    btn_del_acc.pack(side="right", padx=5)

    btn_trash = tk.Button(
        acc_frame,
        text="🗑️ سلة المحذوفات",
        bg="#7f8c8d",
        fg="white",
        font=("Tahoma", 9, "bold"),
        command=self.open_trash_window,
    )
    btn_trash.pack(side="left", padx=10)

    # 2. إطار إدخال الحركة المالية
    input_frame = tk.LabelFrame(
        self.root, text=" إضافة / تعديل حركة مالية ", font=("Tahoma", 10, "bold")
    )
    input_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(input_frame, text="المبلغ:", font=("Tahoma", 9)).grid(
        row=0, column=5, padx=5, pady=5
    )
    self.entry_amount = tk.Entry(input_frame, font=("Tahoma", 9), width=10)
    self.entry_amount.grid(row=0, column=4, padx=5, pady=5)

    tk.Label(input_frame, text="نوع الحركة:", font=("Tahoma", 9)).grid(
        row=0, column=3, padx=5, pady=5
    )
    self.type_var = tk.StringVar(value="عليه")
    rb_against = tk.Radiobutton(
        input_frame,
        text="عليه (مدين)",
        variable=self.type_var,
        value="عليه",
        font=("Tahoma", 9),
    )
    rb_against.grid(row=0, column=2, padx=2)
    rb_for = tk.Radiobutton(
        input_frame,
        text="له (دائن)",
        variable=self.type_var,
        value="له",
        font=("Tahoma", 9),
    )
    rb_for.grid(row=0, column=1, padx=2)

    tk.Label(input_frame, text="الفئة:", font=("Tahoma", 9)).grid(
        row=1, column=5, padx=5, pady=5
    )
    self.entry_cat = tk.Entry(input_frame, font=("Tahoma", 9), width=12)
    self.entry_cat.grid(row=1, column=4, padx=5, pady=5)

    tk.Label(input_frame, text="الوصف/البيان:", font=("Tahoma", 9)).grid(
        row=1, column=3, padx=5, pady=5
    )
    self.entry_desc = tk.Entry(input_frame, font=("Tahoma", 9), width=25)
    self.entry_desc.grid(row=1, column=1, columnspan=2, padx=5, pady=5)

    btn_box = tk.Frame(input_frame)
    btn_box.grid(row=2, column=0, columnspan=6, pady=5)

    self.btn_save = tk.Button(
        btn_box,
        text="إضافة الحركة",
        bg="#27ae60",
        fg="white",
        font=("Tahoma", 9, "bold"),
        command=self.save_transaction,
    )
    self.btn_save.pack(side="right", padx=5)

    self.btn_edit = tk.Button(
        btn_box,
        text="تعديل المحدد",
        bg="#f39c12",
        fg="white",
        font=("Tahoma", 9, "bold"),
        command=self.update_transaction,
    )
    self.btn_edit.pack(side="right", padx=5)

    self.btn_del = tk.Button(
        btn_box,
        text="حذف الحركة الحالية",
        bg="#d35400",
        fg="white",
        font=("Tahoma", 9, "bold"),
        command=self.delete_transaction,
    )
    self.btn_del.pack(side="right", padx=5)

    # 3. جدول العرض
    grid_frame = tk.Frame(self.root)
    grid_frame.pack(fill="both", expand=True, padx=10, pady=5)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(
        "Treeview", font=("Tahoma", 9), rowheight=25, gridlines=True
    )
    style.configure(
        "Treeview.Heading", font=("Tahoma", 9, "bold"), background="#dcdde1"
    )

    columns = ("id", "date", "type", "amount", "category", "desc")
    self.tree = ttk.Treeview(
        grid_frame, columns=columns, show="headings", selectmode="browse"
    )

    self.tree.heading("id", text="الرقم")
    self.tree.heading("date", text="التاريخ والوقت")
    self.tree.heading("type", text="النوع")
    self.tree.heading("amount", text="المبلغ والتأشير")
    self.tree.heading("category", text="الفئة")
    self.tree.heading("desc", text="البيان/الوصف")

    self.tree.column("id", width=40, anchor="center")
    self.tree.column("date", width=140, anchor="center")
    self.tree.column("type", width=70, anchor="center")
    self.tree.column("amount", width=120, anchor="center")
    self.tree.column("category", width=100, anchor="center")
    self.tree.column("desc", width=250, anchor="e")

    self.tree.tag_configure(
        "for_tag", foreground="#2980b9", font=("Tahoma", 9, "bold")
    )
    self.tree.tag_configure(
        "against_tag", foreground="#e74c3c", font=("Tahoma", 9, "bold")
    )

    scrollbar = ttk.Scrollbar(
        grid_frame, orient="vertical", command=self.tree.yview
    )
    self.tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    self.tree.pack(fill="both", expand=True)

    self.tree.bind("<<TreeviewSelect>>", self.on_select_item)

    # 4. التصدير والملخص
    bottom_frame = tk.Frame(self.root, bg="#ecefe6", height=45)
    bottom_frame.pack(fill="x", side="bottom")

    btn_excel = tk.Button(
        bottom_frame,
        text="تصدير Excel",
        bg="#1e7e34",
        fg="white",
        font=("Tahoma", 9, "bold"),
        command=self.export_excel,
    )
    btn_excel.pack(side="left", padx=10, pady=5)

    btn_pdf = tk.Button(
        bottom_frame,
        text="تصدير تقرير PDF",
        bg="#d35400",
        fg="white",
        font=("Tahoma", 9, "bold"),
        command=self.export_pdf,
    )
    btn_pdf.pack(side="left", padx=5, pady=5)

    self.lbl_summary = tk.Label(
        bottom_frame,
        text="إجمالي (له): 0.00 | إجمالي (عليه): 0.00 | صافي الرصيد: 0.00",
        font=("Tahoma", 10, "bold"),
        bg="#ecefe6",
        fg="#2c3e50",
    )
    self.lbl_summary.pack(side="right", padx=15, pady=8)

  def on_account_change(self, event):
    self.current_account = self.acc_combo.get()
    self.refresh_table()

  def add_new_account(self):
    def save_acc():
      name = entry.get().strip()
      if name:
        if name in self.accounts_data:
          messagebox.showerror("خطأ", "اسم الحساب موجود بالفعل.")
          return
        self.accounts_data[name] = []
        self.save_data(self.data_file, self.accounts_data)
        self.update_account_combobox(name)
        top.destroy()

    top = tk.Toplevel(self.root)
    top.title("إضافة حساب جديد")
    top.geometry("300x130")
    tk.Label(top, text="اسم الحساب الجديد:", font=("Tahoma", 10)).pack(pady=5)
    entry = tk.Entry(top, font=("Tahoma", 10))
    entry.pack(pady=5)
    tk.Button(
        top,
        text="حفظ الحساب",
        command=save_acc,
        bg="#27ae60",
        fg="white",
        font=("Tahoma", 9, "bold"),
    ).pack(pady=5)

  def delete_current_account(self):
    if len(self.accounts_data) <= 1:
      messagebox.showwarning(
          "تنبيه", "لا يمكن حذف الحساب الوحيد المتوفر في النظام."
      )
      return

    confirm = messagebox.askyesno(
        "تأكيد الحذف",
        f"هل أنت تأكد من نقل الحساب ('{self.current_account}') إلى سلة"
        " المحذوفات؟\n(يمكنك استرداده خلال 30 يوماً)",
    )
    if confirm:
      self.trash_data[self.current_account] = {
          "deleted_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
          "transactions": self.accounts_data[self.current_account],
      }
      self.save_data(self.trash_file, self.trash_data)

      del self.accounts_data[self.current_account]
      self.save_data(self.data_file, self.accounts_data)

      next_acc = list(self.accounts_data.keys())[0]
      self.update_account_combobox(next_acc)
      messagebox.showinfo(
          "تم الحذف", "تم نقل الحساب إلى سلة المحذوفات بنجاح."
      )

  def open_trash_window(self):
    trash_top = tk.Toplevel(self.root)
    trash_top.title("سلة المحذوفات (تحفظ الحسابات لمدة 30 يوماً)")
    trash_top.geometry("500x350")

    tk.Label(
        trash_top,
        text="الحسابات المحذوفة مؤقتاً:",
        font=("Tahoma", 10, "bold"),
    ).pack(pady=5)

    listbox = tk.Listbox(trash_top, font=("Tahoma", 10))
    listbox.pack(fill="both", expand=True, padx=10, pady=5)

    for acc_name, info in self.trash_data.items():
      del_date = info.get("deleted_date", "غير معروف")
      listbox.insert(tk.END, f"{acc_name} | (تاريخ الحذف: {del_date})")

    def restore_selected():
      selected_idx = listbox.curselection()
      if not selected_idx:
        messagebox.showwarning("تنبيه", "يرجى تحديد حساب لاسترجاعه.")
        return

      selected_text = listbox.get(selected_idx[0])
      acc_name = selected_text.split(" | ")[0]

      if acc_name in self.trash_data:
        self.accounts_data[acc_name] = self.trash_data[acc_name]["transactions"]
        self.save_data(self.data_file, self.accounts_data)

        del self.trash_data[acc_name]
        self.save_data(self.trash_file, self.trash_data)

        self.update_account_combobox(acc_name)
        messagebox.showinfo(
            "تم الاسترجاع", f"تم استعادة الحساب '{acc_name}' بنجاح!"
        )
        trash_top.destroy()

    btn_restore = tk.Button(
        trash_top,
        text="🔄 استرجاع الحساب المحدد",
        bg="#27ae60",
        fg="white",
        font=("Tahoma", 10, "bold"),
        command=restore_selected,
    )
    btn_restore.pack(pady=10)

  def update_account_combobox(self, selected_account):
    self.acc_combo["values"] = list(self.accounts_data.keys())
    self.acc_combo.set(selected_account)
    self.current_account = selected_account
    self.refresh_table()

  def save_transaction(self):
    amount_str = self.entry_amount.get().strip()
    if not amount_str:
      messagebox.showerror("خطأ", "يرجى إدخال المبلغ.")
      return
    try:
      amount = float(amount_str)
    except:
      messagebox.showerror("خطأ", "المبلغ يجب أن يكون رقماً.")
      return

    trans = {
        "id": len(self.accounts_data[self.current_account]) + 1,
        "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "type": self.type_var.get(),
        "amount": amount,
        "category": self.entry_cat.get().strip() or "عام",
        "desc": self.entry_desc.get().strip() or "-",
    }

    self.accounts_data[self.current_account].append(trans)
    self.save_data(self.data_file, self.accounts_data)
    self.clear_inputs()
    self.refresh_table()

  def on_select_item(self, event):
    selected = self.tree.selection()
    if selected:
      item = self.tree.item(selected[0])
      values = item["values"]
      self.selected_item_index = int(values[0]) - 1

      records = self.accounts_data[self.current_account]
      if 0 <= self.selected_item_index < len(records):
        rec = records[self.selected_item_index]
        self.entry_amount.delete(0, tk.END)
        self.entry_amount.insert(0, str(rec["amount"]))
        self.type_var.set(rec["type"])
        self.entry_cat.delete(0, tk.END)
        self.entry_cat.insert(0, rec["category"])
        self.entry_desc.delete(0, tk.END)
        self.entry_desc.insert(0, rec["desc"])

  def update_transaction(self):
    if self.selected_item_index is None:
      messagebox.showwarning("تنبيه", "يرجى اختيار حركة من الجدول للتعديل.")
      return

    try:
      amount = float(self.entry_amount.get().strip())
    except:
      messagebox.showerror("خطأ", "المبلغ يجب أن يكون رقماً.")
      return

    rec = self.accounts_data[self.current_account][self.selected_item_index]
    rec["amount"] = amount
    rec["type"] = self.type_var.get()
    rec["category"] = self.entry_cat.get().strip() or "عام"
    rec["desc"] = self.entry_desc.get().strip() or "-"

    self.save_data(self.data_file, self.accounts_data)
    self.clear_inputs()
    self.refresh_table()

  def delete_transaction(self):
    if self.selected_item_index is None:
      messagebox.showwarning("تنبيه", "يرجى اختيار حركة من الجدول للحذف.")
      return

    del self.accounts_data[self.current_account][self.selected_item_index]
    for idx, item in enumerate(self.accounts_data[self.current_account]):
      item["id"] = idx + 1

    self.save_data(self.data_file, self.accounts_data)
    self.clear_inputs()
    self.refresh_table()

  def clear_inputs(self):
    self.entry_amount.delete(0, tk.END)
    self.entry_cat.delete(0, tk.END)
    self.entry_desc.delete(0, tk.END)
    self.selected_item_index = None

  def refresh_table(self):
    for item in self.tree.get_children():
      self.tree.delete(item)

    total_for = 0.0
    total_against = 0.0

    records = self.accounts_data.get(self.current_account, [])
    for rec in reversed(records):
      if rec["type"] == "له":
        arrow_str = f"▲ {rec['amount']:.2f}"
        tag = "for_tag"
        total_for += rec["amount"]
      else:
        arrow_str = f"▼ {rec['amount']:.2f}"
        tag = "against_tag"
        total_against += rec["amount"]

      self.tree.insert(
          "",
          "end",
          values=(
              rec["id"],
              rec["date"],
              rec["type"],
              arrow_str,
              rec["category"],
              rec["desc"],
          ),
          tags=(tag,),
      )

    net = total_for - total_against
    self.lbl_summary.config(
        text=(
            f"إجمالي (له): {total_for:.2f} | إجمالي (عليه):"
            f" {total_against:.2f} | الصافي: {net:.2f}"
        )
    )

  def export_excel(self):
    if not HAS_EXCEL:
      messagebox.showinfo("تنبيه", "مكتبة openpyxl غير مثبتة.")
      return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
        initialfile=f"كشف_حساب_{self.current_account}.xlsx",
        title="اختر مكان حفظ ملف الإكسيل",
    )

    if not file_path:
      return

    try:
      wb = openpyxl.Workbook()
      ws = wb.active
      ws.title = self.current_account
      ws.append(["الرقم", "التاريخ", "النوع", "المبلغ", "الفئة", "البيان"])

      for rec in self.accounts_data.get(self.current_account, []):
        ws.append([
            rec["id"],
            rec["date"],
            rec["type"],
            rec["amount"],
            rec["category"],
            rec["desc"],
        ])

      wb.save(file_path)
      messagebox.showinfo("نجاح", f"تم تصدير ملف الإكسيل بنجاح إلى:\n{file_path}")
    except Exception as e:
      messagebox.showerror("خطأ", f"تعذر حفظ الملف: {e}")

  def export_pdf(self):
    if not HAS_PDF:
      messagebox.showinfo("تنبيه", "المكتبات الخاصة بالـ PDF غير مثبتة.")
      return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        initialfile=f"كشف_حساب_{self.current_account}.pdf",
        title="اختر مكان حفظ ملف PDF",
    )

    if not file_path:
      return

    try:
      # تسجيل خط Arial الويندوز لدعم النص العربي
      font_path = "C:/Windows/Fonts/arial.ttf"
      if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont("ArabicArial", font_path))
        font_name = "ArabicArial"
      else:
        font_name = "Helvetica"

      doc = SimpleDocTemplate(file_path, pagesize=portrait(A4))
      elements = []

      # عنوان التقرير
      styles = getSampleStyleSheet()
      title_style = ParagraphStyle(
          "ArabicTitle",
          parent=styles["Normal"],
          fontName=font_name,
          fontSize=16,
          alignment=1,  # سنتر
          spaceAfter=15,
      )

      title_text = fix_arabic(f"كشف حساب: {self.current_account}")
      elements.append(Paragraph(title_text, title_style))
      elements.append(Spacer(1, 10))

      # رؤوس الجدول مرتبة من اليمين إلى اليسار
      headers = [
          "البيان/الوصف",
          "الفئة",
          "المبلغ",
          "النوع",
          "التاريخ والوقت",
          "الرقم",
      ]
      table_data = [[fix_arabic(h) for h in headers]]

      # بيانات الجدول
      records = self.accounts_data.get(self.current_account, [])
      for rec in records:
        row = [
            fix_arabic(rec.get("desc", "-")),
            fix_arabic(rec.get("category", "عام")),
            str(rec.get("amount", 0.0)),
            fix_arabic(rec.get("type", "")),
            str(rec.get("date", "")),
            str(rec.get("id", "")),
        ]
        table_data.append(row)

      table = Table(
          table_data, colWidths=[160, 70, 70, 50, 110, 40]
      )  # عرض الأعمدة
      table.setStyle(
          TableStyle([
              ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dcdde1")),
              ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
              ("ALIGN", (0, 0), (-1, -1), "CENTER"),
              ("FONTNAME", (0, 0), (-1, -1), font_name),
              ("FONTSIZE", (0, 0), (-1, -1), 10),
              ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
              ("TOPPADDING", (0, 0), (-1, -1), 6),
              ("GRID", (0, 0), (-1, -1), 1, colors.HexColor("#7f8c8d")),
          ])
      )

      elements.append(table)
      doc.build(elements)

      messagebox.showinfo(
          "نجاح", f"تم تصدير كشف الحساب PDF بنجاح إلى:\n{file_path}"
      )
    except Exception as e:
      messagebox.showerror("خطأ", f"تعذر حفظ ملف PDF: {e}")


if __name__ == "__main__":
  root = tk.Tk()
  app = AccountingApp(root)
  root.mainloop()
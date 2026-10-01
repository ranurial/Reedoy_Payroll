from pathlib import Path
import ast
import re

path = Path("app.py")
if not path.exists():
    raise SystemExit("ERROR: app.py not found. Run this from the Reedoy_Payroll folder.")

text = path.read_text(encoding="utf-8")

# Locate only the LANG dictionary.
start_marker = "# TRANSLATIONS"
start = text.find(start_marker)
if start < 0:
    raise SystemExit("ERROR: TRANSLATIONS section not found.")

lang_start = text.find("LANG = {", start)
if lang_start < 0:
    raise SystemExit("ERROR: LANG dictionary not found.")

# Find the closing dictionary brace before the DATABASE TYPE section.
end_marker = "# ============================================================\n# DATABASE TYPE"
end = text.find(end_marker, lang_start)
if end < 0:
    raise SystemExit("ERROR: DATABASE TYPE marker not found.")

replacement = 'LANG = {\n    "Dashboard": "ড্যাশবোর্ড",\n    "Workers Management": "কর্মী ব্যবস্থাপনা",\n    "Attendance & Calendar": "উপস্থিতি ও ক্যালেন্ডার",\n    "Single Payslip": "একক পে-স্লিপ",\n    "Advance Salary": "অগ্রিম বেতন",\n    "Department Salary Sheet": "বিভাগভিত্তিক বেতন শীট",\n    "Settings": "সেটিংস",\n    "Worker Name": "কর্মীর নাম",\n    "Department": "বিভাগ",\n    "Designation": "পদবি",\n    "Basic Salary": "মূল বেতন",\n    "OT Rate": "OT হার",\n    "Nasta Rate": "নাস্তা হার",\n    "Present": "উপস্থিত",\n    "Absent": "অনুপস্থিত",\n    "Absent Deduction": "অনুপস্থিতির কর্তন",\n    "OT Amt": "OT টাকা",\n    "Nasta": "নাস্তা",\n    "Gross Salary": "মোট বেতন",\n    "Advance": "অগ্রিম",\n    "Net Payable": "নেট প্রদেয়",\n    "Payroll Month": "বেতন মাস",\n    "Advance Date": "অগ্রিমের তারিখ",\n    "Amount (BDT)": "পরিমাণ (টাকা)",\n    "Note": "নোট",\n    "Save": "সংরক্ষণ",\n    "Update": "আপডেট",\n    "Delete": "মুছুন",\n    "Search": "অনুসন্ধান",\n    "Refresh": "রিফ্রেশ",\n    "Generate": "তৈরি করুন",\n    "Export Excel": "এক্সেল রপ্তানি",\n    "Export PDF": "PDF রপ্তানি",\n    "Workers": "কর্মী",\n    "Attendance": "উপস্থিতি",\n    "Advances": "অগ্রিম",\n    "Payroll Report": "বেতন প্রতিবেদন",\n    "Users": "ব্যবহারকারী",\n    "Activity Log": "কার্যক্রমের লগ",\n    "Company Settings": "কোম্পানির সেটিংস",\n    "Bangla": "বাংলা",\n    "English": "ইংরেজি",\n    "Logout": "লগআউট",\n    "Quick Links": "দ্রুত লিংক",\n    "Salary Report": "বেতন প্রতিবেদন",\n    "Add Advance": "অগ্রিম যোগ করুন",\n    "Worker ID": "কর্মী আইডি",\n    "English Name": "ইংরেজি নাম",\n    "Bangla Name": "বাংলা নাম",\n    "Name": "নাম",\n    "Basic Salary (BDT)": "মূল বেতন (টাকা)",\n    "Refreshment Bill": "নাস্তার বিল",\n    "Edit": "সম্পাদনা",\n    "Actions": "কার্যক্রম",\n    "Month": "মাস",\n    "Year": "বছর",\n    "Select Worker": "কর্মী নির্বাচন করুন",\n    "Select Month": "মাস নির্বাচন করুন",\n    "Select Year": "বছর নির্বাচন করুন",\n    "Status": "অবস্থা",\n    "OT Hours": "ওটি ঘণ্টা",\n    "Total Present": "মোট উপস্থিত",\n    "Total Absent": "মোট অনুপস্থিত",\n    "Salary Summary": "বেতনের সারসংক্ষেপ",\n    "No workers found.": "কোনো কর্মী পাওয়া যায়নি।",\n    "Save Worker": "কর্মী সংরক্ষণ করুন",\n    "Add Worker": "কর্মী যোগ করুন",\n    "Update Worker": "কর্মীর তথ্য আপডেট করুন",\n    "Worker name is required.": "কর্মীর নাম আবশ্যক।",\n    "Worker saved successfully.": "কর্মীর তথ্য সফলভাবে সংরক্ষিত হয়েছে।",\n}'

new_text = text[:lang_start] + replacement + "\n\n\n" + text[end:]

path.write_text(new_text, encoding="utf-8")
print("LANGUAGE BLOCK FIXED SUCCESSFULLY")
print("Only the LANG translation block was replaced.")

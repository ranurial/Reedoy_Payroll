from pathlib import Path
import re

p = Path("app.py")
text = p.read_text(encoding="utf-8")

new_block = '''DEPT_BN = {
    "All Departments": "সকল বিভাগ",
    "General": "সাধারণ",
    "Printing": "প্রিন্টিং",
    "Jigar": "জিগার",
    "Wash": "ওয়াশ",
    "Loop": "লুপ",
    "Stanter": "স্ট্যান্টার",
    "Electrical": "ইলেকট্রিক্যাল",
    "Accounts": "অ্যাকাউন্টস",
    "Design": "ডিজাইন",
}'''

pattern = r'DEPT_BN\s*=\s*\{.*?\n\}'

text2, count = re.subn(pattern, new_block, text, count=1, flags=re.S)

if count != 1:
    print("ERROR: DEPT_BN block পাওয়া যায়নি। কোনো পরিবর্তন করা হয়নি।")
else:
    p.write_text(text2, encoding="utf-8")
    print("DEPT_BN FIXED SUCCESSFULLY")

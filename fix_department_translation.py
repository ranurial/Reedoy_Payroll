from pathlib import Path
path=Path('app.py')
text=path.read_text(encoding='utf-8')
replacements={
'"Printing": "à¦ªà§à¦°à¦¿à¦¨à§à¦Ÿà¦¿à¦‚"':'"Printing": "প্রিন্টিং"',
'"Jigar": "à¦œà¦¿à¦—à¦¾à¦°"':'"Jigar": "জিগার"',
'"Wash": "à¦“à¦¯à¦¼à¦¾à¦¶"':'"Wash": "ওয়াশ"',
'"Loop": "à¦²à§à¦ª"':'"Loop": "লুপ"',
'"Stanter": "à¦¸à§à¦Ÿà§à¦¯à¦¾à¦¨à§à¦Ÿà¦¾à¦°"':'"Stanter": "স্ট্যান্টার"',}
changed=0
for old,new in replacements.items():
    if old in text:
        text=text.replace(old,new); changed+=1
path.write_text(text,encoding='utf-8')
print('DEPARTMENT TRANSLATION FIXED')
print('Changed:',changed,'of',len(replacements))
if changed != len(replacements): print('WARNING: Some entries were not found.')

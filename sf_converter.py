#!/usr/bin/env python3
"""
Smart Farmer Data Converter
แปลง Excel/CSV จากแบบฟอร์มติดตามผล → 3 ไฟล์ CSV พร้อม Import Google Sheets
วิธีใช้: python3 sf_converter.py ชื่อไฟล์.xlsx 2568
"""

import sys
import os
import re
import pandas as pd

# ── แปลงเลขบัตรประชาชน ──
def fix_id(val):
    if val is None or (isinstance(val, float) and pd.isna(val)): return ''
    s = str(val).strip()
    # เลขไทย
    th = {'๐':'0','๑':'1','๒':'2','๓':'3','๔':'4','๕':'5','๖':'6','๗':'7','๘':'8','๙':'9'}
    s = ''.join(th.get(c,c) for c in s)
    # ลบขีด ช่องว่าง
    s = re.sub(r'[-\s]','',s)
    # Scientific Notation
    if 'e+' in s.lower():
        try: s = str(int(float(s)))
        except: pass
    s = re.sub(r'\.0+$','',s)
    s = re.sub(r'[^0-9]','',s)
    return s

def safe(val):
    if val is None or (isinstance(val, float) and pd.isna(val)): return ''
    s = str(val).strip()
    return '' if s in ('nan','None','NaN','') else s

def to_mee(val):
    s = safe(val)
    if not s: return ''
    try:
        if float(s) >= 1: return 'มี'
    except: pass
    return 'มี' if s in ('1','TRUE','true','มี') else ''

def to_use(val):
    s = safe(val)
    if not s: return ''
    try:
        if float(s) >= 1: return 'นำไปใช้'
    except: pass
    return ''

def get_result(r):
    if safe(r[105]) and _to_float(r[105]) >= 1: return 'ต้นแบบ'
    if safe(r[104]) and _to_float(r[104]) >= 1: return 'Existing'
    if safe(r[103]) and _to_float(r[103]) >= 1: return 'Developing'
    return 'Developing'

def _to_float(val):
    try: return float(safe(val))
    except: return 0

# ── หัวคอลัมน์ ──
MASTER_H = ['เลขบัตรประชาชน','คำนำหน้า','ชื่อ','นามสกุล','วันเกิด',
    'บ้านเลขที่','หมู่ที่','ตำบล (ตามบัตร)','อำเภอ (ตามบัตร)','จังหวัด',
    'สมาชิกครัวเรือน','แรงงานเกษตร','ระดับการศึกษา','วุฒิการศึกษา','สาขา',
    'เป็นเจ้าของแปลง','ความสัมพันธ์กับเจ้าของแปลง','ชื่อเจ้าของแปลง',
    'เลขบัตรเจ้าของแปลง','ขึ้นทะเบียนเกษตรอำเภอ','ขึ้นทะเบียนการยาง',
    'สมาชิกสหกรณ์','สมาชิกแปลงใหญ่','วิสาหกิจชุมชน','สถานะ SF']

PLOT_H = ['เลขบัตรประชาชน (FK)','ชื่อ-นามสกุล (VLOOKUP)',
    'พื้นที่ทั้งหมด-ไร่','พื้นที่ทั้งหมด-งาน','พื้นที่ทั้งหมด-วา',
    'พื้นที่ร่วม ส.ป.ก.-ไร่','พื้นที่ร่วม ส.ป.ก.-งาน','พื้นที่ร่วม ส.ป.ก.-วา',
    'แปลงเลขที่','กลุ่ม/ระวาง ส.ป.ก.','หมู่ที่ (ที่ตั้งแปลง)',
    'ตำบล (ที่ตั้งแปลง)','อำเภอ (ที่ตั้งแปลง)','จังหวัด (ที่ตั้งแปลง)',
    'แหล่งน้ำ','ไฟฟ้า','สัญญาณโทรศัพท์','ถนน','การคมนาคม']

HIST_H = ['เลขบัตรประชาชน (FK)','ชื่อ-นามสกุล (VLOOKUP)','ปีงบประมาณ',
    'ทำนา','ปลูกผัก','ผลไม้','สมุนไพร','พืชไร่','ปศุสัตว์','แปรรูป','อื่นๆ',
    'พื้นที่ไร่','ผลผลิตจำนวน','ผลผลิตหน่วย',
    'ช่องทาง-พ่อค้า/ล้ง','ช่องทาง-ตลาดชุมชน','ช่องทาง-ห้าง',
    'ช่องทาง-ออนไลน์','ชื่อเพจ',
    'รายได้เกษตรก่อน','รายได้เกษตรหลัง','สรุปรายได้','ผลต่าง(บาท)','ผลต่าง(%)',
    'ความรู้-บริหารแปลง','ความรู้-พัฒนาผลิตภัณฑ์','ความรู้-ช่องทางขาย',
    'ความรู้-วางแผนอนาคต','ความรู้-เผยแพร่','ไม่ได้นำไปใช้เนื่องจาก',
    'เปลี่ยน-รายได้เพิ่ม','เปลี่ยน-รายจ่ายลด','เปลี่ยน-ต้นทุนลด',
    'เปลี่ยน-วางแผนผลิต','เปลี่ยน-วางแผนการตลาด','เปลี่ยน-ผลผลิตเพิ่ม',
    'เปลี่ยน-ผลผลิตมีคุณภาพ','เปลี่ยน-อื่นๆ',
    'ปัญหา/อุปสรรค','ผลการประเมิน SF','ประเภทเกษตรกร',
    'ผู้จัดเก็บข้อมูล','เบอร์ติดต่อ','วันที่ลงข้อมูล']

def convert(filepath, year='2568', sheet_name='รวมข้อมูล'):
    ext = os.path.splitext(filepath)[1].lower()
    print(f'อ่านไฟล์: {filepath}')

    if ext == '.xlsx':
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=0, dtype=str)
    elif ext == '.csv':
        # ลองหาแถวที่เป็นข้อมูลจริง (col0 เป็นตัวเลข)
        raw = pd.read_csv(filepath, header=None, dtype=str, encoding='utf-8-sig')
        header_row = 0
        for i, row in raw.iterrows():
            val = str(row.iloc[0]).strip()
            if re.match(r'^\d+$', val):
                header_row = i - 1 if i > 0 else 0
                break
        df = pd.read_csv(filepath, header=header_row, dtype=str, encoding='utf-8-sig')
    else:
        raise ValueError(f'รองรับเฉพาะ .xlsx และ .csv')

    rows = df.values.tolist()
    print(f'พบแถวข้อมูล: {len(rows)} แถว')

    master_rows, plot_rows, hist_rows = [], [], []
    skipped, warns = 0, []

    for idx, r in enumerate(rows):
        # padding ให้ครบ 113 col
        while len(r) < 113: r.append('')

        pid = fix_id(r[6])
        if not pid or len(pid) < 10:
            print(f'  ข้ามแถว {idx+2}: pid=[{safe(r[6])}]')
            skipped += 1
            continue
        if len(pid) != 13:
            warns.append(f'แถว {idx+2}: เลขบัตร "{pid}" ความยาว {len(pid)} หลัก')

        result = get_result(r)
        bday = f'{safe(r[7])}-{safe(r[8])}-{safe(r[9])}' if safe(r[7]) and safe(r[8]) and safe(r[9]) else ''

        master_rows.append([
            pid, safe(r[2]), safe(r[3]), safe(r[4]), bday,
            safe(r[10]), safe(r[11]), safe(r[12]), safe(r[13]), safe(r[14]),
            safe(r[15]), safe(r[16]),
            safe(r[23]), safe(r[26]), safe(r[27]),
            safe(r[17]), safe(r[18]), safe(r[20]), fix_id(r[21]),
            'มี' if safe(r[28]) in ('1','1.0') else '',
            'มี' if safe(r[29]) in ('1','1.0') else '',
            safe(r[33]), safe(r[35]), safe(r[36]), result
        ])

        plot_rows.append([
            pid, '',
            safe(r[38]), safe(r[39]), safe(r[40]),
            safe(r[41]), safe(r[42]), safe(r[43]),
            safe(r[44]), safe(r[45]),
            safe(r[46]), safe(r[47]), safe(r[48]), safe(r[49]),
            safe(r[50]), safe(r[51]), safe(r[52]), safe(r[53]), safe(r[54])
        ])

        hist_rows.append([
            pid, '', year,
            safe(r[69]), safe(r[70]), safe(r[71]), safe(r[72]),
            safe(r[73]), safe(r[74]), safe(r[75]), safe(r[76]),
            safe(r[77]), safe(r[80]), safe(r[81]),
            to_mee(r[82]), to_mee(r[83]), safe(r[84]), safe(r[85]), safe(r[86]),
            safe(r[64]), safe(r[65]), safe(r[66]), safe(r[67]), safe(r[68]),
            to_use(r[88]), to_use(r[89]), to_use(r[90]), to_use(r[91]), to_use(r[92]),
            safe(r[93]),
            to_mee(r[94]), to_mee(r[95]), to_mee(r[96]), to_mee(r[97]),
            to_mee(r[98]), to_mee(r[99]), to_mee(r[100]), safe(r[101]),
            safe(r[102]), result, safe(r[110]),
            safe(r[107]), safe(r[108]), safe(r[109])
        ])
        print(f'  ✓ {safe(r[2])}{safe(r[3])} {safe(r[4])} [{pid}] → {result}')

    # บันทึกไฟล์
    base = os.path.splitext(os.path.basename(filepath))[0]
    out_dir = os.path.dirname(os.path.abspath(filepath))

    for name, headers, data in [
        (f'{base}_import_Master.csv', MASTER_H, master_rows),
        (f'{base}_import_Plot.csv',   PLOT_H,   plot_rows),
        (f'{base}_import_History.csv',HIST_H,   hist_rows),
    ]:
        out_path = os.path.join(out_dir, name)
        pd.DataFrame(data, columns=headers).to_csv(out_path, index=False, encoding='utf-8-sig')
        print(f'  บันทึก: {name} ({len(data)} แถว)')

    print(f'\n✅ สำเร็จ: {len(master_rows)} ราย | ข้าม: {skipped} แถว')
    if warns:
        print(f'⚠ ควรตรวจสอบ {len(warns)} รายการ:')
        for w in warns: print(f'  - {w}')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('วิธีใช้: python3 sf_converter.py ชื่อไฟล์.xlsx [ปีงบประมาณ] [ชื่อ Sheet]')
        print('ตัวอย่าง: python3 sf_converter.py data.xlsx 2568 รวมข้อมูล')
        sys.exit(1)
    filepath = sys.argv[1]
    year = sys.argv[2] if len(sys.argv) > 2 else '2568'
    sheet = sys.argv[3] if len(sys.argv) > 3 else 'รวมข้อมูล'
    convert(filepath, year, sheet)

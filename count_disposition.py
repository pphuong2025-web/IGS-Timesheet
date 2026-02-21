"""Count entries with disposition date 1/22 in columns I and N of sheet VR-TS1."""
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# Namespaces in xlsx
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def get_col_letter_index(col_letter):
    """Convert column letter to 0-based index: A=0, I=8, N=13."""
    n = 0
    for c in col_letter.upper():
        n = n * 26 + (ord(c) - ord('A') + 1)
    return n - 1

def col_index_from_ref(cell_ref):
    """From cell ref like 'A1' or 'N42', return 0-based column index."""
    col = ''
    for c in cell_ref:
        if c.isalpha():
            col += c
        else:
            break
    return get_col_letter_index(col)

def _excel_serial_to_md(serial):
    """Convert Excel serial to (month, day) or None. Excel epoch 1899-12-30."""
    from datetime import datetime, timedelta
    try:
        n = float(serial)
        if n < 1 or n > 100000:
            return None
        # Excel: 1 = 1900-01-01. Epoch 1899-12-30.
        d = datetime(1899, 12, 30) + timedelta(days=int(n))
        return (d.month, d.day)
    except (ValueError, TypeError, OverflowError):
        return None

def is_date_1_22(val, cell_type=None):
    """Return True if value looks like disposition date 1/22."""
    # Excel serial date
    if isinstance(val, (int, float)) and 1 < val < 100000:
        md = _excel_serial_to_md(val)
        if md and md[0] == 1 and md[1] == 22:
            return True
    if val is None:
        return False
    if isinstance(val, float) and (val != val or val == 0.0):
        return False
    s = str(val).strip()
    if not s:
        return False
    # Try Excel serial when it's a numeric string
    try:
        n = float(s)
        if 1 < n < 100000:
            md = _excel_serial_to_md(n)
            if md and md[0] == 1 and md[1] == 22:
                return True
    except ValueError:
        pass
    # Match 1/22, 1/22/2025, 01/22, etc.
    if '1/22' in s:
        return True
    parts = s.replace(',', ' ').split()
    for p in parts:
        p = p.strip()
        if '/' in p:
            pp = p.split('/')
            if len(pp) >= 2:
                try:
                    m, d = int(pp[0]), int(pp[1])
                    if m == 1 and d == 22:
                        return True
                except ValueError:
                    pass
    return False

def main():
    import shutil
    import tempfile
    src = Path(r'c:\Users\phuong.pham\OneDrive - Foxconn Industrial Internet in North America\Desktop\test_dir\NV_IGS_VR144_Bonepile (6).xlsx')
    path = src
    try:
        with zipfile.ZipFile(path, 'r') as z:
            pass
    except PermissionError:
        # Copy to temp if original is locked (e.g. open in Excel)
        path = Path(tempfile.gettempdir()) / 'NV_IGS_VR144_Bonepile_temp.xlsx'
        shutil.copy2(str(src), str(path))
    col_i_idx = 8   # I
    col_n_idx = 13  # N

    # Find sheet and shared strings
    with zipfile.ZipFile(path, 'r') as z:
        # Load workbook to get sheet id for VR-TS1
        wb = z.read('xl/workbook.xml')
        wb_root = ET.fromstring(wb)
        sheets = wb_root.findall('.//main:sheet', NS) or wb_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')
        if not sheets:
            sheets = [e for e in wb_root.iter() if 'sheet' in (e.tag or '') and e.get('name')]

        # Resolve namespace in tag
        def find_sheets(root):
            for e in root.iter():
                t = e.tag
                if t and 'sheet' in t.lower() and e.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'):
                    yield e
                if t and 'Sheet' in t:
                    yield e

        sheet_elem = None
        sheet_rid = None
        for e in wb_root.iter():
            name = e.get('name') or e.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}name')
            if name and 'VR-TS1' in str(name):
                sheet_elem = e
                sheet_rid = e.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') or e.get('r:id')
                break

        if sheet_elem is None:
            # Try by name in any attribute
            for e in wb_root.iter():
                for a, v in (e.attrib or {}).items():
                    if 'VR-TS1' in str(v):
                        sheet_rid = e.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') or e.get('r:id')
                        break

        # Get rId to path: read [Content_Types].xml and xl/_rels/workbook.xml.rels
        rels = z.read('xl/_rels/workbook.xml.rels')
        rels_root = ET.fromstring(rels)
        sheet_path = None
        for r in rels_root.iter():
            rid = r.get('Id') or r.get('{http://schemas.openxmlformats.org/package/2006/relationships}Id')
            if rid == sheet_rid:
                target = r.get('Target') or r.get('{http://schemas.openxmlformats.org/package/2006/relationships}Target')
                if target:
                    sheet_path = 'xl/' + target.replace('..', '')
                break

        if not sheet_path:
            # List sheets by rId from workbook
            import re
            wb_s = wb.decode('utf-8', errors='replace')
            # find sheet name="VR-TS1" r:id="rId2" -> rId2
            m = re.search(r'VR-TS1[^>]*r:id="([^"]+)"', wb_s) or re.search(r'name="VR-TS1"[^>]*r:id="([^"]+)"', wb_s)
            if m:
                sheet_rid = m.group(1)
            for r in rels_root.iter():
                rid = r.get('Id') or r.get('{http://schemas.openxmlformats.org/package/2006/relationships}Id')
                if rid == sheet_rid:
                    target = r.get('Target') or r.get('{http://schemas.openxmlformats.org/package/2006/relationships}Target')
                    if target:
                        sheet_path = 'xl/' + target.lstrip('./')
                    break

        if not sheet_path:
            # Fallback: common name
            names = [n for n in z.namelist() if 'sheet' in n and 'xl/worksheets' in n]
            # Get sheet index from workbook: we need VR-TS1's order
            import re
            wb_s = wb.decode('utf-8', errors='replace')
            idx = 0
            for i, part in enumerate(re.split(r'<[^>]*sheet[^>]*>', wb_s, flags=re.I)):
                if 'VR-TS1' in part or (i > 0 and 'VR-TS1' in wb_s[:wb_s.find(part)]):
                    # find which sheet element
                    break
            # Simpler: try sheet2.xml, sheet3.xml, etc. if workbook has multiple
            for n in sorted(names):
                if 'sheet1' in n:
                    sheet_path = n
                    break
                if 'sheet2' in n:
                    sheet_path = n
                    break
            if not sheet_path and names:
                sheet_path = names[0]

        # Prefer finding by sheet name in workbook
        wb_s = wb.decode('utf-8', errors='replace')
        import re
        # <sheet name="VR-TS1" sheetId="2" r:id="rId2"/>
        m = re.search(r'<[^>]*\s+name="VR-TS1"[^>]*\s+r:id="(rId\d+)"', wb_s)
        if not m:
            m = re.search(r'r:id="(rId\d+)"[^>]*\s+name="VR-TS1"', wb_s)
        if m:
            wanted_rid = m.group(1)
            rels_s = rels.decode('utf-8', errors='replace')
            # <Relationship Id="rId2" Type="..." Target="worksheets/sheet2.xml"/>
            m2 = re.search(r'Id="' + re.escape(wanted_rid) + r'"[^>]*Target="([^"]+)"', rels_s)
            if not m2:
                m2 = re.search(r'Target="([^"]+)"[^>]*Id="' + re.escape(wanted_rid) + '"', rels_s)
            if m2:
                t = m2.group(1).replace('..', '')
                sheet_path = 'xl/' + t if not t.startswith('xl/') else t

        if not sheet_path:
            # list and try each
            for n in z.namelist():
                if 'worksheets/sheet' in n and n.endswith('.xml'):
                    sheet_path = n
                    break

        if not sheet_path:
            print('Could not find VR-TS1 sheet path')
            return

        # Shared strings: each <si> holds one string (possibly with <r><t>...</t></r> rich runs)
        shared = []
        try:
            ss_xml = z.read('xl/sharedStrings.xml')
            ss_root = ET.fromstring(ss_xml)
            for si in ss_root.iter():
                if (si.tag or '').endswith('}si'):
                    shared.append(''.join(si.itertext()))
        except Exception:
            pass

        # Read worksheet
        try:
            ws = z.read(sheet_path)
        except KeyError:
            # try without xl/
            try:
                ws = z.read(sheet_path.replace('xl/xl/', 'xl/'))
            except Exception:
                ws = z.read('xl/worksheets/sheet1.xml')

        ws_root = ET.fromstring(ws)
        # Get rows: <row r="1">...</row>
        # Cell: <c r="I5" t="s"><v>123</v></c>  t=s means shared string

        count_i = 0
        count_n = 0

        for row in ws_root.iter():
            tag = (row.tag or '')
            if 'row' not in tag.lower():
                continue
            for c in row:
                ct = (c.tag or '')
                if 'c' not in ct or ct[-2:] == 'c/' or ct[-1] != 'c':
                    continue
                # it's a cell
                r = c.get('r') or c.get('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}r')
                if not r:
                    continue
                col_idx = col_index_from_ref(r)
                if col_idx != col_i_idx and col_idx != col_n_idx:
                    continue

                # get value: <v> or inline <is><t>
                v_el = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                if v_el is None:
                    v_el = c.find('v')
                if v_el is not None:
                    val = v_el.text or ''
                else:
                    is_el = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is')
                    if is_el is not None:
                        val = ''.join(is_el.itertext())
                    else:
                        val = ''

                is_shared = (c.get('t') or c.get('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')) == 's'
                if is_shared and shared and val.isdigit():
                    idx = int(val)
                    if 0 <= idx < len(shared):
                        val = shared[idx]

                if is_date_1_22(val):
                    if col_idx == col_i_idx:
                        count_i += 1
                    else:
                        count_n += 1

    print('Column I (disposition date 1/22):', count_i)
    print('Column N (disposition date 1/22):', count_n)

if __name__ == '__main__':
    main()

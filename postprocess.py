#!/usr/bin/env python3
"""Excel 后处理程序 — 坐落解析 + 证件替换 + 排序 + 双表输出"""

import json
import os
import re
import openpyxl
from openpyxl.utils import get_column_letter


# ===================== 配置 =====================
INPUT_FILE = "/Users/nickzu/Desktop/月亮湾/副本月亮湾业主清册2.xlsx"
JSON_DIR = "/Users/nickzu/Desktop/deepseek测试"
OUTPUT_DIR = "/Users/nickzu/Desktop/月亮湾"

PRESET_POOL_FILE = os.path.join(JSON_DIR, "postprocess.json")
MAPPING_FILE = os.path.join(JSON_DIR, "mapping.json")


# ===================== 坐落解析 =====================
def parse_zuoluo(text):
    text = text.strip()
    if not text:
        return 0, ""
    m = re.match(r'^(\d+)\幢(\d+)$', text)
    if m:
        return int(m.group(1)), m.group(2)
    m = re.match(r'^地下室(\d[\d\-A-Za-z、]+)$', text)
    if m:
        return 0, m.group(1)
    m = re.match(r'^(\d+)\幢(.+)$', text)
    if m:
        return int(m.group(1)), m.group(2).strip()
    m = re.match(r'^地下室(.+)$', text)
    if m:
        return 0, m.group(1).strip()
    return 0, text


# ===================== 面积提取 =====================
def extract_area(text):
    m = re.search(r'房屋建筑面积([\d.]+)平方米', str(text))
    return float(m.group(1)) if m else 0.0


# ===================== 排序键 =====================
def get_sort_key(building, room):
    if building > 0 and re.match(r'^\d+$', room):
        return (0, building, int(room))
    elif building == 0 and re.match(r'^\d', room):
        pfx = int(re.match(r'^(\d+)', room).group(1))
        return (1, pfx, room)
    else:
        return (2, building, room)


# ===================== 输出列生成 =====================
def building_col(building, room):
    if building > 0:
        return f"{building}幢"
    if re.match(r'^\d', room):
        return "地下室（车位）"
    return "地下室"


def room_col(room):
    return room


# ===================== 证件号判断 =====================
import datetime

def is_mainland_id(s):
    s = s.strip()
    if not re.match(r'^\d{17}[\dX]$', s):
        return False
    try:
        birth = s[6:14]
        dt = datetime.datetime.strptime(birth, '%Y%m%d')
        return 1900 <= dt.year <= 2099
    except ValueError:
        return False


# ===================== JSON 读写 =====================
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ===================== 核心处理 =====================
def process_excel(input_path, output_path, pool, mapping):
    wb = openpyxl.load_workbook(input_path)
    ws = wb[wb.sheetnames[0]]

    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        seq, holder, zuoluo, cert_ids, usage, area = row

        building, room = parse_zuoluo(str(zuoluo) if zuoluo else "")

        # 解析证件号 → 替换非 18 位
        id_list = [x.strip() for x in str(cert_ids).split(",")] if cert_ids else []
        new_ids = []
        for id_num in id_list:
            if not id_num:
                new_ids.append(id_num)
            elif is_mainland_id(id_num):
                new_ids.append(id_num)
            elif id_num in mapping:
                new_ids.append(mapping[id_num])
            else:
                if not pool:
                    raise RuntimeError("虚拟号池已耗尽！")
                assigned = pool.pop(0)
                mapping[id_num] = assigned
                new_ids.append(assigned)

        new_cert = ",".join(new_ids) if new_ids else ""

        rows.append({
            "seq": seq,
            "holder": str(holder) if holder else "",
            "zuoluo": str(zuoluo) if zuoluo else "",
            "cert_ids": new_cert,
            "usage": str(usage) if usage else "",
            "area": extract_area(area),
            "building": building,
            "room": room,
            "sort_key": get_sort_key(building, room),
        })

    rows.sort(key=lambda r: r["sort_key"])

    # 仅保留住宅（sort_key[0]==0，即幢号+纯数字房号）
    rows = [r for r in rows if r["sort_key"][0] == 0]

    # ===== 构建输出 =====
    out_wb = openpyxl.Workbook()

    fc_ws = out_wb.active
    fc_ws.title = "fc"
    fc_ws.append(["楼栋", "房号", "权证面积"])

    yz_ws = out_wb.create_sheet("yz")
    yz_ws.append(["楼栋号", "房号", "权利人"])

    for r in rows:
        ld = building_col(r["building"], r["room"])
        fh = room_col(r["room"])

        fc_ws.append([ld, fh, r["area"]])

        # 权利人列：姓名(证件号) 拼接
        names = [n.strip() for n in r["holder"].split(",")] if r["holder"] else []
        ids = [x.strip() for x in r["cert_ids"].split(",")] if r["cert_ids"] else []
        combined = ""
        for j, name in enumerate(names):
            cid = ids[j] if j < len(ids) else ""
            combined += f"{name}({cid})" if cid else name

        yz_ws.append([ld, fh, combined])

    # 列宽自适应
    for sht in [fc_ws, yz_ws]:
        for ci in range(1, sht.max_column + 1):
            mx = 0
            for cell in sht[get_column_letter(ci)]:
                if cell.value:
                    mx = max(mx, len(str(cell.value)))
            sht.column_dimensions[get_column_letter(ci)].width = min(mx + 2, 40)

    # 权证面积 2 位小数
    for cell in fc_ws["C"][1:]:
        cell.number_format = "0.00"

    out_wb.save(output_path)
    return rows


# ===================== 入口 =====================
def main():
    raw_pool = load_json(PRESET_POOL_FILE)
    mapping = load_json(MAPPING_FILE) if os.path.exists(MAPPING_FILE) else {}

    # 从池中移除已被映射占用的虚拟号
    used = set(mapping.values())
    pool = [v for v in raw_pool if v not in used]

    print(f"虚拟号池: {len(pool)} 个 (原始 {len(raw_pool)}，已占用 {len(used)})")
    print(f"已有映射: {len(mapping)} 条")

    basename = os.path.splitext(os.path.basename(INPUT_FILE))[0]
    out_path = os.path.join(OUTPUT_DIR, f"{basename}_纯住宅.xlsx")

    print(f"\n处理中: {INPUT_FILE}")
    rows = process_excel(INPUT_FILE, out_path, pool, mapping)

    save_json(MAPPING_FILE, mapping)
    save_json(PRESET_POOL_FILE, pool)

    print(f"输出: {out_path}")
    print(f"  fc: {len(rows)} 行")
    print(f"  yz: {len(rows)} 行")
    print(f"更新映射: {len(mapping)} 条")
    print(f"剩余池: {len(pool)} 个")

    a = sum(1 for r in rows if r["sort_key"][0] == 0)
    b = sum(1 for r in rows if r["sort_key"][0] == 1)
    c = sum(1 for r in rows if r["sort_key"][0] == 2)
    print(f"分类 — 住宅: {a} | 车位: {b} | 公共设施: {c}")


if __name__ == "__main__":
    main()

from __future__ import annotations

import json
import re
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import pdfplumber
from docx import Document
from openpyxl import load_workbook


PROJECT = Path(__file__).resolve().parent
ROOT = PROJECT.parent
DIST = PROJECT / "dist"
FILES_OUT = DIST / "files"

HEADERS = [
    "stt", "code", "field", "stages", "group", "name", "indicators",
    "owner", "source_type", "format", "purpose", "frequency",
]

FIELD_MAP = {
    "1": "Ngân sách trung ương của bộ, cơ quan trung ương",
    "2": "Ngân sách địa phương",
    "3": "Quyết toán ngân sách nhà nước",
    "4": "Dự án đầu tư xây dựng",
    "5": "Doanh nghiệp và các tổ chức tài chính, ngân hàng",
    "6": "Chuyên đề",
    "7": "Phục vụ chung toàn Ngành",
}

STAGE_MAP = {
    "A1": "Chuẩn bị kiểm toán",
    "A2": "Thực hiện kiểm toán",
    "A3": "Lập và gửi báo cáo kiểm toán",
    "A4": "Theo dõi thực hiện kết luận, kiến nghị kiểm toán",
    "B1": "Xây dựng kế hoạch kiểm toán trung hạn và hằng năm",
    "B2": "Kiểm soát chất lượng kiểm toán",
    "B3": "Tổng hợp, báo cáo kết quả kiểm toán",
    "B4": "Theo dõi thực hiện kết luận, kiến nghị ở cấp Ngành",
    "B5": "Công khai kết quả kiểm toán",
    "B6": "Thanh tra",
}

SOURCE_TYPE_MAP = {
    "1": "Hệ thống của đơn vị được kiểm toán hoặc đơn vị liên quan",
    "2": "Cơ sở dữ liệu quốc gia, chuyên ngành",
    "3": "Hồ sơ, tài liệu rời do đơn vị quản lý",
    "4": "Dữ liệu từ cơ quan thanh tra, kiểm tra, giám sát",
    "5": "Dữ liệu công khai, dữ liệu mở",
    "6": "Dữ liệu phải mua hoặc thuê quyền truy cập",
}

FORMAT_MAP = {
    "1": "Bản giấy",
    "2": "Bản scan hoặc PDF không bóc tách được",
    "3": "Tệp văn bản hoặc bảng tính rời",
    "4": "Dữ liệu trong hệ thống thông tin, cơ sở dữ liệu",
    "5": "Dữ liệu công bố trên môi trường mạng",
}

FREQUENCY_MAP = {
    "1": "Thường xuyên",
    "2": "Định kỳ",
    "3": "Không thường xuyên",
}


def clean(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def expand_code_list(value, mapping, alpha=False):
    text = clean(value)
    if not text:
        return ""
    parts = [part.strip() for part in re.split(r"[,;/]+", text) if part.strip()]
    expanded = []
    for part in parts:
        key = part.upper() if alpha else part.lstrip("0") or "0"
        if key in mapping:
            expanded.append(mapping[key])
            continue
        field_with_note = re.fullmatch(r"0?([1-7])\s*[-–]\s*(.+)", part)
        if mapping is FIELD_MAP and field_with_note:
            expanded.append(f"{mapping[field_with_note.group(1)]}: {field_with_note.group(2)}")
            continue
        return text
    return "; ".join(dict.fromkeys(expanded))


def normalize_frequency(value):
    text = clean(value)
    if not text:
        return ""
    if re.fullmatch(r"[0-9\s,;/]+", text):
        return expand_code_list(text, FREQUENCY_MAP)
    lowered = text.casefold()
    levels = []
    if "không thường xuyên" in lowered:
        levels.append("Không thường xuyên")
        lowered = lowered.replace("không thường xuyên", "")
    if "thường xuyên" in lowered:
        levels.append("Thường xuyên")
    if "định kỳ" in lowered or "hằng năm" in lowered or "hàng năm" in lowered:
        levels.append("Định kỳ")
    return "; ".join(dict.fromkeys(levels)) or text


def to_record(values, unit):
    values = [clean(v) for v in list(values)[:12]]
    values += [""] * (12 - len(values))
    record = dict(zip(HEADERS, values))
    record["field"] = expand_code_list(record["field"], FIELD_MAP)
    record["stages"] = expand_code_list(record["stages"], STAGE_MAP, alpha=True)
    record["source_type"] = expand_code_list(record["source_type"], SOURCE_TYPE_MAP)
    record["format"] = expand_code_list(record["format"], FORMAT_MAP)
    record["frequency"] = normalize_frequency(record["frequency"])
    record["unit"] = unit
    return record


def is_column_index_row(values, count):
    normalized = [clean(value) for value in list(values)[:count]]
    return normalized == [str(index) for index in range(1, count + 1)]


def content_key(value):
    return re.sub(r"[^\w]+", " ", clean(value).casefold(), flags=re.UNICODE).strip()


def resolve_formula(wb, value):
    if not isinstance(value, str) or not value.startswith("="):
        return value
    match = re.fullmatch(r"='([^']+)'!([A-Z]+\d+)", value)
    if not match:
        match = re.fullmatch(r"=([^!]+)!([A-Z]+\d+)", value)
    if match:
        return wb[match.group(1)][match.group(2)].value
    return value


def read_excel(path, unit, detail_sheet, detail_start, priority_sheet, priority_start):
    wb = load_workbook(path, data_only=False, read_only=False)
    ws = wb[detail_sheet]
    details = []
    for row in ws.iter_rows(min_row=detail_start, max_col=12, values_only=True):
        marker = row[0]
        is_data_marker = isinstance(marker, (int, float)) or (isinstance(marker, str) and marker.startswith("="))
        if is_data_marker and not is_column_index_row(row, 12) and clean(row[1]):
            details.append(to_record(row, unit))

    ws = wb[priority_sheet]
    priorities = []
    rank = 0
    for row in ws.iter_rows(min_row=priority_start, max_col=4, values_only=True):
        marker = row[0]
        if not (isinstance(marker, (int, float)) or (isinstance(marker, str) and marker.startswith("="))):
            continue
        if is_column_index_row(row, 4):
            continue
        code = clean(resolve_formula(wb, row[1]))
        name = clean(resolve_formula(wb, row[2]))
        if not code or not name:
            continue
        rank += 1
        priorities.append({"rank": rank, "code": code, "name": name, "reason": clean(row[3]), "unit": unit})
    return details, priorities


def read_kv4(path):
    detail_pages = {2, 3, 4, 7, 8, 11, 12, 15}
    priority_pages = {5, 9, 13, 16}
    details, priorities = [], []
    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables() or []
            for table in tables:
                if page_no in detail_pages:
                    for row in table:
                        if row and re.fullmatch(r"\d{1,2}", clean(row[0])) and clean(row[1]) not in {"2", "Mã nguồn"}:
                            details.append(to_record(row, "KV4"))
                elif page_no in priority_pages:
                    for row in table:
                        if row and re.fullmatch(r"\d+", clean(row[0])):
                            priorities.append({
                                "rank": len(priorities) + 1,
                                "section_rank": int(clean(row[0])),
                                "code": clean(row[1]),
                                "name": clean(row[2]),
                                "reason": clean(row[3]),
                                "unit": "KV4",
                            })
    return details, priorities


def read_docx_forms(path, unit):
    doc = Document(path)
    details, priorities = [], []
    if doc.tables:
        for row in doc.tables[0].rows[2:]:
            values = [cell.text for cell in row.cells]
            if values and re.fullmatch(r"\d+", clean(values[0])) and not is_column_index_row(values, 12) and clean(values[1]):
                details.append(to_record(values, unit))
    if len(doc.tables) > 1:
        for row in doc.tables[1].rows[1:]:
            values = [clean(cell.text) for cell in row.cells]
            if values and re.fullmatch(r"\d+", values[0]) and not is_column_index_row(values, 4) and values[1] and values[2]:
                priorities.append({
                    "rank": int(values[0]),
                    "code": values[1],
                    "name": values[2],
                    "reason": values[3],
                    "unit": unit,
                })
    return details, priorities


def copy_files(unit, folder):
    destination = FILES_OUT / unit
    destination.mkdir(parents=True, exist_ok=True)
    result = []
    for path in sorted(folder.rglob("*")):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        target = destination / path.name
        source_stat = path.stat()
        if not target.exists() or target.stat().st_size != source_stat.st_size:
            shutil.copy2(path, target)
        result.append({
            "name": path.name,
            "download": f"files/{unit}/{path.name}",
            "type": path.suffix.lower().lstrip(".").upper(),
            "size": source_stat.st_size,
            "updated": datetime.fromtimestamp(source_stat.st_mtime).isoformat(timespec="minutes"),
        })
    return result


def main():
    DIST.mkdir(parents=True, exist_ok=True)
    FILES_OUT.mkdir(parents=True, exist_ok=True)

    kv4_details, kv4_priorities = read_kv4(ROOT / "KV4" / "KTNN KV4 _CV998.pdf")
    kv7_details, kv7_priorities = read_excel(
        ROOT / "KV7" / "Files_30-9-2026_09-28-05" / "1. Phụ lục CV_998_CNTT.xlsx",
        "KV7", "Biểu 1", 7, "Biểu 2", 6,
    )
    kv8_details, kv8_priorities = read_excel(
        ROOT / "KV8" / "Phuluc_CV998_Layykien.xlsx",
        "KV8", "BIEU 01", 7, "BIEU 02", 7,
    )
    k12_details, k12_priorities = read_excel(
        ROOT / "K12" / "KV12. PL nguon du lieu phuc vu hoat dong kiem toan.xlsx",
        "KV12", "PL01", 8, "PL02", 6,
    )
    cn5_details, cn5_priorities = read_docx_forms(
        ROOT / "CN5" / "Phu_bieu_01_02_03_nguon_du_lieu_KTNN_CNV.docx", "CN5"
    )

    details = kv4_details + kv7_details + kv8_details + k12_details + cn5_details
    priorities = kv4_priorities + kv7_priorities + kv8_priorities + k12_priorities + cn5_priorities

    file_folders = {
        "KV4": ROOT / "KV4",
        "KV7": ROOT / "KV7",
        "KV8": ROOT / "KV8",
        "KV12": ROOT / "K12",
        "CN5": ROOT / "CN5",
    }
    files = {unit: copy_files(unit, folder) for unit, folder in file_folders.items()}
    request_files = copy_files("yeu-cau-goc", ROOT / "CV yêu cầu gốc")

    unit_defs = [
        ("KV4", "Kiểm toán nhà nước khu vực IV", kv4_details, kv4_priorities, "Một PDF 16 trang; tách nội dung theo 4 lĩnh vực kiểm toán.", "Một phần"),
        ("KV7", "Kiểm toán nhà nước khu vực VII", kv7_details, kv7_priorities, "Văn bản trả lời và phụ lục Excel; tập trung ngân sách địa phương, thuế, hải quan và đầu tư.", "Chưa có"),
        ("KV8", "Kiểm toán nhà nước khu vực VIII", kv8_details, kv8_priorities, "Có đủ 3 biểu; phạm vi mã nguồn rộng nhất trong các hồ sơ hiện có.", "Đầy đủ"),
        ("KV12", "Kiểm toán nhà nước khu vực XII", k12_details, k12_priorities, "Văn bản trả lời và phụ lục Excel; nhiều nguồn thuế, TABMIS, đầu tư và dữ liệu bổ sung.", "Chưa có"),
        ("CN5", "Kiểm toán nhà nước chuyên ngành V", cn5_details, cn5_priorities, "Có đủ 3 biểu; tập trung vào hai Bộ, doanh nghiệp nhà nước và các dự án đầu tư thuộc phạm vi kiểm toán.", "Đầy đủ"),
    ]
    units = []
    for code, name, unit_details, unit_priorities, note, form3 in unit_defs:
        units.append({
            "code": code,
            "name": name,
            "detail_count": len(unit_details),
            "priority_count": len(unit_priorities),
            "form1": True,
            "form2": True,
            "form3": form3,
            "note": note,
            "files": files[code],
        })

    priority_presence = Counter()
    for unit in ["KV4", "KV7", "KV8", "KV12", "CN5"]:
        codes = {
            p["code"] for p in priorities
            if p["unit"] == unit and p["code"] and not re.fullmatch(r"BS-\d+", p["code"], flags=re.IGNORECASE)
        }
        priority_presence.update(codes)
    bs_matches = {}
    for priority in priorities:
        code = priority["code"].upper()
        if not re.fullmatch(r"BS-\d+", code):
            continue
        key = (code, content_key(priority["name"]))
        bs_matches.setdefault(key, set()).add(priority["unit"])
    for (code, _), matched_units in bs_matches.items():
        if len(matched_units) >= 2:
            priority_presence[code] = max(priority_presence[code], len(matched_units))
    common = [
        {"code": code, "units": count}
        for code, count in sorted(priority_presence.items(), key=lambda item: (-item[1], item[0]))
    ][:12]

    issues = [
        {"severity": "Cao", "unit": "KV8", "title": "Mã BS-01 chưa thống nhất", "detail": "Biểu 02 xếp BS-01 là báo cáo nguồn và nhu cầu cải cách tiền lương, trong khi Biểu 01 dùng BS-01 cho giấy phép khai thác khoáng sản; báo cáo cải cách tiền lương mang mã BS-02."},
        {"severity": "Cao", "unit": "KV4", "title": "Danh sách ưu tiên vượt giới hạn", "detail": "Hồ sơ nêu 19 ưu tiên theo 4 nhóm lĩnh vực, trong khi yêu cầu gốc quy định tối đa 10 nguồn cho một đơn vị."},
        {"severity": "Cao", "unit": "KV7", "title": "Danh sách ưu tiên vượt giới hạn", "detail": "Biểu 02 có 20 nguồn ưu tiên; cần chọn lại tối đa 10 nguồn và sắp xếp theo mức độ ưu tiên giảm dần."},
        {"severity": "Cao", "unit": "KV12", "title": "Danh sách ưu tiên vượt giới hạn", "detail": "Biểu 02 có 28 nguồn ưu tiên; cần chọn lại tối đa 10 nguồn theo yêu cầu của Công văn."},
        {"severity": "Trung bình", "unit": "KV8", "title": "Mã bổ sung và số thứ tự bị lặp", "detail": "Công văn nêu 1 nguồn bổ sung nhưng phụ lục xuất hiện BS-01 đến BS-08 và nhiều đoạn đánh lại số thứ tự; cần hợp nhất trước khi tổng hợp toàn ngành."},
        {"severity": "Trung bình", "unit": "KV7", "title": "Lý do ưu tiên còn chung chung", "detail": "Nhiều dòng chỉ nêu phục vụ khảo sát/lập kế hoạch, chưa chỉ rõ hậu quả nếu thiếu dữ liệu hoặc dữ liệu giữ nguyên hình thức hiện tại."},
        {"severity": "Trung bình", "unit": "KV7", "title": "Thiếu Biểu số 03 và email người lập", "detail": "Hồ sơ hiện có 2 biểu; trường email người lập để trống và chưa có báo cáo khó khăn, vướng mắc, kiến nghị."},
        {"severity": "Trung bình", "unit": "KV12", "title": "Thiếu Biểu số 03", "detail": "Hồ sơ hiện có Biểu 01 và Biểu 02 nhưng chưa có báo cáo khó khăn, vướng mắc và kiến nghị."},
        {"severity": "Theo dõi", "unit": "KV4", "title": "Biểu số 03 chưa bao phủ đủ 4 lĩnh vực", "detail": "Có báo cáo vướng mắc cho đầu tư, chi NSNN và thu NSNN; chưa thấy phần tương ứng cho doanh nghiệp và tổ chức tài chính, ngân hàng."},
    ]

    data = {
        "meta": {
            "title": "Tổng hợp nguồn dữ liệu phục vụ hoạt động kiểm toán",
            "reference": "Công văn 998/CNTT-UDDLS ngày 24/9/2026",
            "updated": datetime.now().isoformat(timespec="minutes"),
            "unit_count": len(units),
            "detail_count": len(details),
            "priority_count": len(priorities),
            "file_count": sum(len(v) for v in files.values()),
        },
        "request": {
            "objective": "Xây dựng danh mục dữ liệu ưu tiên, kế hoạch kết nối, chia sẻ dữ liệu và kho dữ liệu dùng chung của Ngành; giảm việc yêu cầu đơn vị được kiểm toán cung cấp lại dữ liệu đã có.",
            "required": [
                "Biểu 01: Danh mục nguồn dữ liệu gắn với yêu cầu nghiệp vụ cụ thể",
                "Biểu 02: Tối đa 10 nguồn ưu tiên kết nối, khai thác trước",
                "Biểu 03: Khó khăn, vướng mắc và kiến nghị theo từng mã nguồn",
                "Yêu cầu chất lượng: đúng, đủ, sạch, sống",
                "Nêu cơ quan quản lý, phương thức, tần suất, căn cứ và điều kiện bảo mật",
            ],
            "deadline": "30/9/2026",
            "files": request_files,
        },
        "units": units,
        "details": details,
        "priorities": priorities,
        "common_priorities": common,
        "issues": issues,
    }
    (DIST / "data.js").write_text("window.DASHBOARD_DATA = " + json.dumps(data, ensure_ascii=False) + ";\n", encoding="utf-8")
    print(json.dumps(data["meta"], ensure_ascii=True))


if __name__ == "__main__":
    main()

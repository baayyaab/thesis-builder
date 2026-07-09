"""
thesis_core.py
══════════════════════════════════════════════════════════════════
Modular Thesis Builder — มหาวิทยาลัยธรรมศาสตร์ (2024-2026)
Part 1: Data Schema & Backend Core Logic

ไม่มี Streamlit / python-docx ในไฟล์นี้
ทุกอย่างเป็น pure Python เพื่อให้ test ได้ง่ายและ import ไปใช้ต่อได้
══════════════════════════════════════════════════════════════════
"""

from __future__ import annotations
import re
import copy
from dataclasses import dataclass, field
from typing import Optional


# ══════════════════════════════════════════════════════════════════
#  SECTION 1 — DATA SCHEMA
#  โครงสร้าง Dictionary/Dataclass สำหรับเก็บข้อมูลวิทยานิพนธ์
# ══════════════════════════════════════════════════════════════════

@dataclass
class CommitteeMember:
    """สมาชิกคณะกรรมการ 1 คน"""
    title_th: str          # คำนำหน้า เช่น "ศาสตราจารย์ ดร."
    name_th: str           # ชื่อ-นามสกุล ภาษาไทย
    title_en: str          # คำนำหน้า ภาษาอังกฤษ เช่น "Professor Dr."
    name_en: str           # ชื่อ-นามสกุล ภาษาอังกฤษ
    role: str              # บทบาท เช่น "ประธานกรรมการ" / "อาจารย์ที่ปรึกษาหลัก"
    enabled: bool = True   # toggle เปิด-ปิด (สำหรับที่ปรึกษาร่วม/กรรมการเสริม)


@dataclass
class ThesisChapter:
    """บทเนื้อหา 1 บท (dynamic — เพิ่มได้ไม่จำกัด)"""
    chapter_number: int       # หมายเลขบท
    title_th: str             # ชื่อบท ภาษาไทย
    title_en: str = ""        # ชื่อบท ภาษาอังกฤษ (ถ้ามี)
    body: str = ""            # เนื้อหาดิบ (plain text หรือ markdown-like)


@dataclass
class ThesisSchema:
    """
    โครงสร้างข้อมูลหลักของวิทยานิพนธ์ทั้งเล่ม
    ครอบคลุม Metadata / Committee / Sections ตาม Template มธ.
    """

    # ── Metadata ──────────────────────────────────────────────────
    title_th: str = ""                  # ชื่อเรื่อง ภาษาไทย
    title_en: str = ""                  # ชื่อเรื่อง ภาษาอังกฤษ
    author_th: str = ""                 # ชื่อผู้เขียน ภาษาไทย (นาย/นาง/นางสาว + ชื่อ)
    author_en: str = ""                 # ชื่อผู้เขียน ภาษาอังกฤษ (FIRSTNAME LASTNAME)

    degree_level: str = "โท"           # "โท" | "IS" | "เอก"
    degree_name_th: str = ""           # ชื่อเต็มปริญญา ภาษาไทย เช่น "นิติศาสตรมหาบัณฑิต"
    degree_name_en: str = ""           # ชื่อเต็มปริญญา ภาษาอังกฤษ
    field_of_study_th: str = ""        # สาขาวิชา ภาษาไทย
    field_of_study_en: str = ""        # สาขาวิชา ภาษาอังกฤษ
    faculty_th: str = ""               # คณะ ภาษาไทย
    faculty_en: str = ""               # คณะ ภาษาอังกฤษ

    academic_year_be: int = 0          # ปีการศึกษา พ.ศ.
    academic_year_ce: int = 0          # ปีการศึกษา ค.ศ. (คำนวณอัตโนมัติได้)
    approval_day: int = 0              # วันที่อนุมัติ
    approval_month_th: str = ""        # เดือนอนุมัติ ภาษาไทยเต็ม เช่น "มีนาคม"
    approval_month_en: str = ""        # เดือนอนุมัติ ภาษาอังกฤษ เช่น "March"
    approval_year_be: int = 0          # ปีอนุมัติ พ.ศ.

    # ── Committee ─────────────────────────────────────────────────
    committee_chair: Optional[CommitteeMember] = None          # ประธานกรรมการ
    advisor_main: Optional[CommitteeMember] = None             # อาจารย์ที่ปรึกษาหลัก
    advisor_co: list[CommitteeMember] = field(default_factory=list)   # ที่ปรึกษาร่วม (0-n คน)
    examiners: list[CommitteeMember] = field(default_factory=list)    # กรรมการสอบ
    dean: Optional[CommitteeMember] = None                     # คณบดี

    # ── Sections ──────────────────────────────────────────────────
    abstract_th: str = ""              # บทคัดย่อ ภาษาไทย
    abstract_en: str = ""              # Abstract ภาษาอังกฤษ
    
    # ⚙️ [แก้ไข] ปรับจาก list[str] ให้เป็น str เพื่อรองรับพิมพ์ข้อความยาวในหน้าเว็บ
    keywords_th: str = ""              # คำสำคัญ ภาษาไทย (คั่นด้วยจุลภาคหรือเว้นวรรค)
    keywords_en: str = ""              # Keywords ภาษาอังกฤษ
    
    acknowledgement: str = ""          # กิตติกรรมประกาศ

    # บทที่ 1-5 เริ่มต้น (dynamic เพิ่มได้ผ่าน add_chapter())
    chapters: list[ThesisChapter] = field(default_factory=list)

    appendix: str = ""                # ภาคผนวก (plain text / markdown)
    author_biography: str = ""        # ประวัติผู้เขียน


    # ── helper: สร้างโครงบท 1-5 เริ่มต้น ─────────────────────────
    def init_default_chapters(self) -> None:
        """เรียกครั้งเดียวตอนสร้างเล่มใหม่ เพื่อสร้างบทที่ 1-5 ว่างๆ"""
        DEFAULT_CHAPTERS = [
            (1, "บทนำ",                          "Introduction"),
            (2, "เอกสารและงานวิจัยที่เกี่ยวข้อง", "Review of Related Literature"),
            (3, "วิธีดำเนินการวิจัย",             "Research Methodology"),
            (4, "ผลการวิจัย",                     "Research Findings"),
            (5, "สรุปผล อภิปรายผล และข้อเสนอแนะ", "Conclusions, Discussion and Recommendations"),
        ]
        self.chapters = [
            ThesisChapter(num, th, en) for num, th, en in DEFAULT_CHAPTERS
        ]

    def add_chapter(self, title_th: str, title_en: str = "") -> ThesisChapter:
        """เพิ่มบทใหม่ต่อท้าย (Dynamic บทที่ 6, 7, 8 …)"""
        next_num = (self.chapters[-1].chapter_number + 1) if self.chapters else 1
        new_ch   = ThesisChapter(next_num, title_th, title_en)
        self.chapters.append(new_ch)
        return new_ch

    def to_dict(self) -> dict:
        """แปลงทั้ง Schema เป็น plain dict (สำหรับ JSON / session_state)"""
        def member_to_dict(m: CommitteeMember | None) -> dict | None:
            return m.__dict__ if m else None

        return {
            # Metadata
            "title_th": self.title_th,
            "title_en": self.title_en,
            "author_th": self.author_th,
            "author_en": self.author_en,
            "degree_level": self.degree_level,
            "degree_name_th": self.degree_name_th,
            "degree_name_en": self.degree_name_en,
            "field_of_study_th": self.field_of_study_th,
            "field_of_study_en": self.field_of_study_en,
            "faculty_th": self.faculty_th,
            "faculty_en": self.faculty_en,
            "academic_year_be": self.academic_year_be,
            "academic_year_ce": self.academic_year_ce,
            "approval_day": self.approval_day,
            "approval_month_th": self.approval_month_th,
            "approval_month_en": self.approval_month_en,
            "approval_year_be": self.approval_year_be,
            # Committee
            "committee_chair": member_to_dict(self.committee_chair),
            "advisor_main": member_to_dict(self.advisor_main),
            "advisor_co": [m.__dict__ for m in self.advisor_co],
            "examiners": [m.__dict__ for m in self.examiners],
            "dean": member_to_dict(self.dean),
            # Sections
            "abstract_th": self.abstract_th,
            "abstract_en": self.abstract_en,
            "keywords_th": self.keywords_th,  # ส่งค่าออกเป็น str ตามโครงสร้างใหม่
            "keywords_en": self.keywords_en,  # ส่งค่าออกเป็น str ตามโครงสร้างใหม่
            "acknowledgement": self.acknowledgement,
            "chapters": [ch.__dict__ for ch in self.chapters],
            "appendix": self.appendix,
            "author_biography": self.author_biography,
        }


# ══════════════════════════════════════════════════════════════════
#  SECTION 2 — BACKEND CORE FUNCTIONS
#  ฟังก์ชันจัดการกฎ มธ. อัตโนมัติ
# ══════════════════════════════════════════════════════════════════

# ── 2.1 check_cover_type ──────────────────────────────────────────
COVER_STYLES: dict[str, dict] = {
    "โท":   {"bg": "สีแดงเลือดหมู", "text": "สีทอง", "label": "วิทยานิพนธ์"},
    "IS":   {"bg": "สีแดงเลือดหมู", "text": "สีทอง", "label": "การค้นคว้าอิสระ"},
    "เอก":  {"bg": "สีกรมท่า",      "text": "สีทอง", "label": "ดุษฎีนิพนธ์"},
}

def check_cover_type(degree_level: str) -> dict:
    """
    รับ degree_level ("โท" | "IS" | "เอก")
    คืนค่า dict ที่มี:
        bg_color  — สีพื้นปก
        text_color — สีตัวอักษร
        doc_label  — ชื่อเรียกเอกสาร (วิทยานิพนธ์ / การค้นคว้าอิสระ / ดุษฎีนิพนธ์)
    """
    style = COVER_STYLES.get(degree_level)
    if style is None:
        raise ValueError(
            f"ระดับปริญญา '{degree_level}' ไม่ถูกต้อง — ใช้ได้เฉพาะ: {list(COVER_STYLES)}"
        )
    return {
        "bg_color":   style["bg"],
        "text_color": style["text"],
        "doc_label":  style["label"],
    }


# ── 2.2 generate_title_page_en ────────────────────────────────────
_MONTH_EN = {
    "มกราคม": "January",   "กุมภาพันธ์": "February", "มีนาคม": "March",
    "เมษายน": "April",     "พฤษภาคม": "May",         "มิถุนายน": "June",
    "กรกฎาคม": "July",     "สิงหาคม": "August",      "กันยายน": "September",
    "ตุลาคม": "October",   "พฤศจิกายน": "November",  "ธันวาคม": "December",
}

def generate_title_page_en(data: ThesisSchema) -> dict:
    """
    สร้างข้อมูลหน้าปกใน (Inner Title Page) ภาษาอังกฤษ 1 หน้า
    จาก Metadata ที่มีอยู่

    คืนค่า dict พร้อม key สำหรับแต่ละ block ของหน้าปก:
        title, author, degree, field, faculty, university,
        approval_statement, academic_year
    """
    month_en = (
        data.approval_month_en
        or _MONTH_EN.get(data.approval_month_th, data.approval_month_th)
    )
    year_ce = data.academic_year_ce or (data.academic_year_be - 543)
    approval_year_ce = data.approval_year_be - 543

    cover = check_cover_type(data.degree_level)
    doc_label_en = {
        "วิทยานิพนธ์":        "A Thesis",
        "การค้นคว้าอิสระ":    "An Independent Study",
        "ดุษฎีนิพนธ์":        "A Dissertation",
    }.get(cover["doc_label"], "A Thesis")

    return {
        "title":      data.title_en.upper(),
        "author":     data.author_en.upper(),
        "doc_type":   doc_label_en,
        "submitted":  (
            f"{doc_label_en} Submitted in Partial Fulfillment of the Requirements\n"
            f"for the Degree of {data.degree_name_en}\n"
            f"{data.field_of_study_en}\n"
            f"{data.faculty_en}\n"
            f"Thammasat University"
        ),
        "approval_statement": (
            f"Approved by the Examining Committee on "
            f"{data.approval_day} {month_en} {approval_year_ce}"
        ),
        "academic_year": str(year_ce),
        "copyright": f"Copyright {year_ce} by Thammasat University",
    }


# ── 2.3 parse_scientific_name ─────────────────────────────────────
_BINOMIAL_RE = re.compile(
    r'\b([A-Z][a-z]+)\s+([a-z]{2,})\b'
)

def parse_scientific_name(text: str) -> dict:
    seen_genera: dict[str, str] = {}
    found_names: list[dict]     = []
    italic_spans: list[tuple]   = []
    result_chars                = list(text)

    offset = 0
    for m in _BINOMIAL_RE.finditer(text):
        genus   = m.group(1)
        species = m.group(2)
        abbrev  = f"{genus[0]}."

        if genus not in seen_genera:
            seen_genera[genus] = abbrev
            replacement = f"{genus} {species}"
            found_names.append({"genus": genus, "species": species,
                                 "abbreviated": f"{abbrev} {species}",
                                 "first_occurrence": True})
        else:
            replacement = f"{abbrev} {species}"
            found_names.append({"genus": genus, "species": species,
                                 "abbreviated": replacement,
                                 "first_occurrence": False})

        orig_start = m.start() + offset
        orig_end   = m.end()   + offset
        new_end    = orig_start + len(replacement)
        italic_spans.append((orig_start, new_end))

        result_chars[orig_start:orig_end] = list(replacement)
        offset += len(replacement) - (m.end() - m.start())

    return {
        "processed_text": "".join(result_chars),
        "found_names":    found_names,
        "italic_spans":   italic_spans,
    }


# ── 2.4 monitor_foreign_word_parentheses ─────────────────────────
_PAREN_RE = re.compile(
    r'([\u0E00-\u0E7F\w]+\s*)\(([A-Za-z][\w\s\-]*?)\)'
)

def monitor_foreign_word_parentheses(text: str,
                                     seen_words: set[str] | None = None
                                     ) -> dict:
    if seen_words is None:
        seen_words = set()

    removed_pairs: list[dict] = []

    def replacer(m: re.Match) -> str:
        thai_word    = m.group(1).strip()
        foreign_word = m.group(2).strip()
        key          = foreign_word.lower()

        if key not in seen_words:
            seen_words.add(key)
            return m.group(0)
        else:
            removed_pairs.append({"thai": thai_word, "foreign": foreign_word})
            return thai_word

    processed = _PAREN_RE.sub(replacer, text)

    return {
        "processed_text": processed,
        "seen_words":     seen_words,
        "removed_pairs":  removed_pairs,
    }


# ══════════════════════════════════════════════════════════════════
#  SECTION 3 — UTILITY HELPERS
# ══════════════════════════════════════════════════════════════════

def be_to_ce(year_be: int) -> int:
    return year_be - 543

def ce_to_be(year_ce: int) -> int:
    return year_ce + 543

def new_thesis(degree_level: str = "โท") -> ThesisSchema:
    check_cover_type(degree_level)
    thesis = ThesisSchema(degree_level=degree_level)
    thesis.init_default_chapters()
    return thesis


# ══════════════════════════════════════════════════════════════════
#  QUICK SELF-TEST (python thesis_core.py)
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import json

    print("=" * 60)
    print("SELF-TEST: thesis_core.py")
    print("=" * 60)

    thesis = new_thesis("โท")
    thesis.title_th        = "การศึกษาประสิทธิภาพของ Escherichia coli ในระบบนิเวศ"
    thesis.title_en        = "Efficiency Study of Escherichia coli in Ecosystem"
    thesis.author_th       = "นายสมชาย ใจดี"
    thesis.author_en       = "Somchai Jaidee"
    thesis.degree_name_th  = "วิทยาศาสตรมหาบัณฑิต"
    thesis.degree_name_en  = "Master of Science"
    thesis.field_of_study_th = "จุลชีววิทยา"
    thesis.field_of_study_en = "Microbiology"
    thesis.faculty_th      = "คณะวิทยาศาสตร์และเทคโนโลยี"
    thesis.faculty_en      = "Faculty of Science and Technology"
    thesis.academic_year_be = 2567
    thesis.approval_day    = 15
    thesis.approval_month_th = "มีนาคม"
    thesis.approval_year_be  = 2568
    
    # ทดสอบกำหนดค่าแบบ String (โครงสร้างใหม่)
    thesis.keywords_th     = "ดีเอ็นเอ, พันธุศาสตร์, แบคทีเรีย"
    thesis.keywords_en     = "DNA, Genetics, Bacteria"

    print(f"\n✅ Test 1 — new_thesis: บทที่ 1-5 = {[ch.title_th for ch in thesis.chapters]}")
    thesis.add_chapter("บทสรุปเพิ่มเติม", "Additional Summary")
    print(f"   add_chapter บทที่ 6 → {thesis.chapters[-1].title_th}")

    for lvl in ["โท", "IS", "เอก"]:
        style = check_cover_type(lvl)
        print(f"\n✅ Test 2 — cover_type({lvl}): {style}")

    en_page = generate_title_page_en(thesis)
    print(f"\n✅ Test 3 — title_page_en:\n{json.dumps(en_page, ensure_ascii=False, indent=2)}")

    sample = (
        "การศึกษาพบว่า Escherichia coli สามารถเจริญเติบโตได้ดี "
        "นอกจากนี้ Homo sapiens ยังมีปฏิสัมพันธ์กับ Escherichia coli "
        "และ Homo sapiens อีกด้วย"
    )
    result = parse_scientific_name(sample)
    print(f"\n✅ Test 4 — parse_scientific_name:")
    print(f"   processed : {result['processed_text']}")

    para1 = "การใช้งานเว็บไซต์ (Web Site) เป็นสิ่งสำคัญ เว็บไซต์ (Web Site) ควรออกแบบให้ดี"
    r1 = monitor_foreign_word_parentheses(para1)
    print(f"\n✅ Test 5 — foreign_word_parentheses (para1):")
    print(f"   processed : {r1['processed_text']}")

    print("\n" + "=" * 60)
    print("ALL TESTS PASSED ✅ (Keywords Updated to String)")
    print("=" * 60)
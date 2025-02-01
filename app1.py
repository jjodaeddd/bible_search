import sys
import json
import re
from PyQt5.QtWidgets import QApplication, QMainWindow, QLineEdit, QPushButton, QTextEdit, QVBoxLayout, QWidget
from pptx import Presentation
from pptx.util import Inches

# ✅ 성경 데이터 불러오기
def load_bible():
    with open("C:/Users/admin/Desktop/dh/bible_search/bible.json", "r", encoding="utf-8-sig") as f:
        return json.load(f)

bible_data = load_bible()  # 성경 데이터 로드

# ✅ 책 이름 약어 변환 딕셔너리 (전체)
book_abbreviations = {
    "창세기": "창", "창세": "창", "창": "창", "ㅊ": "창",
    "출애굽기": "출", "출애": "출", "출": "출", "ㅊㅇ": "출",
    "레위기": "레", "레위": "레", "ㄹㅇ": "레",
    "민수기": "민", "민수": "민", "ㅁㅅ": "민",
    "신명기": "신", "신명": "신", "ㅅㅁ": "신",
    "여호수아": "수", "여호": "수", "ㅇㅎ": "수",
    "사사기": "삿", "사사": "삿", "ㅅㅅ": "삿",
    "룻기": "룻", "ㄹ": "룻",
    "사무엘상": "삼상", "ㅅㅁㅇㅅ": "삼상",
    "사무엘하": "삼하", "ㅅㅁㅇㅎ": "삼하",
    "열왕기상": "왕상", "ㅇㅇㄱㅅ": "왕상",
    "열왕기하": "왕하", "ㅇㅇㄱㅎ": "왕하",
    "역대상": "대상", "ㅇㄷㅅ": "대상",
    "역대하": "대하", "ㅇㄷㅎ": "대하",
    "에스라": "스", "스라": "스", "ㅇㅅㄹ": "스",
    "느헤미야": "느", "느헤": "느", "ㄴㅎ": "느",
    "에스더": "에", "더": "에", "ㅇㅅㄷ": "에",
    "욥기": "욥", "ㅇ": "욥",
    "시편": "시", "ㅅㅍ": "시",
    "잠언": "잠", "ㅈㅇ": "잠",
    "전도서": "전", "전도": "전", "ㅈㄷ": "전",
    "아가": "아", "ㅇㄱ": "아",
    "이사야": "사", "사야": "사", "ㅇㅅㅇ": "사",
    "예레미야": "렘", "예레": "렘", "ㅇㄹㅁㅇ": "렘",
    "예레미야애가": "애", "애가": "애", "ㅇㄹㅁㅇㅇㄱ": "애",
    "에스겔": "겔", "ㅇㅅㄱ": "겔",
    "다니엘": "단", "니엘": "단", "ㄷㄴㅇ": "단",
    "호세아": "호", "호세": "호", "ㅎㅅㅇ": "호",
    "요엘": "욜", "엘": "욜", "ㅇㅇ": "욜",
    "아모스": "암", "모스": "암", "ㅇㅁㅅ": "암",
    "오바댜": "옵", "바댜": "옵", "ㅇㅂㄷ": "옵",
    "요나": "욘", "나": "욘", "ㅇㄴ": "욘",
    "미가": "미", "가": "미", "ㅁㄱ": "미",
    "나훔": "나", "훔": "나", "ㄴㅎ": "나",
    "하박국": "합", "박국": "합", "ㅎㅂㄱ": "합",
    "스바냐": "습", "바냐": "습", "ㅅㅂㄴ": "습",
    "학개": "학", "개": "학", "ㅎㄱ": "학",
    "스가랴": "슥", "가랴": "슥", "ㅅㄱㄹ": "슥",
    "말라기": "말", "라기": "말", "ㅁㄹㄱ": "말",
    "마태복음": "마", "마태": "마", "ㅁㅌ": "마",
    "마가복음": "막", "마가": "막", "ㅁㄱ": "막",
    "누가복음": "눅", "누가": "눅", "ㄴㄱ": "눅",
    "요한복음": "요", "요한": "요", "ㅇㅎ": "요",
    "사도행전": "행", "행전": "행", "ㅅㄷㅎㅈ": "행",
    "로마서": "롬", "로마": "롬", "ㄹㅁ": "롬",
    "고린도전서": "고전", "ㄱㄹㄷㅈ": "고전",
    "고린도후서": "고후", "ㄱㄹㄷㅎ": "고후",
    "갈라디아서": "갈", "갈라": "갈", "ㄱㄹㄷ": "갈",
    "에베소서": "엡", "에베": "엡", "ㅇㅂㅅ": "엡",
    "빌립보서": "빌", "빌립": "빌", "ㅂㄹㅂ": "빌",
    "골로새서": "골", "골로": "골", "ㄱㄹㅅ": "골",
    "데살로니가전서": "살전", "ㄷㅅㄹㄴㄱㅈ": "살전",
    "데살로니가후서": "살후", "ㄷㅅㄹㄴㄱㅎ": "살후",
    "디모데전서": "딤전", "ㄷㅁㄷㅈ": "딤전",
    "디모데후서": "딤후", "ㄷㅁㄷㅎ": "딤후",
    "디도서": "딛", "도서": "딛", "ㄷㄷ": "딛",
    "빌레몬서": "몬", "레몬": "몬", "ㅂㄹㅁ": "몬",
    "히브리서": "히", "히브": "히", "ㅎㅂㄹ": "히",
    "야고보서": "약", "고보": "약", "ㅇㄱㅂ": "약",
    "베드로전서": "벧전", "ㅂㄷㄹㅈ": "벧전",
    "베드로후서": "벧후", "ㅂㄷㄹㅎ": "벧후",
    "요한일서": "요일", "ㅇㅎㅇ": "요일",
    "요한이서": "요이", "ㅇㅎㅇ": "요이",
    "요한삼서": "요삼", "ㅇㅎㅅ": "요삼",
    "유다서": "유", "유다": "유", "ㅇㄷ": "유",
    "요한계시록": "계", "계시": "계", "ㅇㅎㄱㅅㄹ": "계"
}

# ✅ 검색어 변환 함수 (수정됨)
def parse_query(query):
    query = re.sub(r'\s+|장|절', '', query).replace(',', ':')
    pattern = r'^([가-힣ㄱ-ㅎ]+)?\s*(\d+):(\d+)(?:-(\d+))?$'
    match = re.match(pattern, query)

    if not match:
        return None, None, None, None

    book, chapter, start_verse, end_verse = match.groups()
    
    if book:
        book = book_abbreviations.get(book, book)
    
    return book, chapter, start_verse, end_verse or start_verse


# ✅ PPT 파일 생성 함수
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

def create_ppt(book, chapter, verse, text):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank slide
    
    # Set background color (black)
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = f"[{book} {chapter}:{verse}]"
    title_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 0)  # Yellow
    title_frame.paragraphs[0].font.size = Pt(40)
    
    # Add content
    content_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(5))
    content_frame = content_box.text_frame
    
    verses = text.split('\n')[2:]  # Extract only verse content
    for verse in verses:
        p = content_frame.add_paragraph()
        p.text = verse.strip()
        p.font.color.rgb = RGBColor(255, 255, 255)  # White
        p.font.size = Pt(28)
    
    file_name = f"{book}_{chapter}_{verse}.pptx"
    prs.save(file_name)
    return file_name


class BibleSearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('성경 검색 프로그램')
        self.setGeometry(100, 100, 600, 400)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.search_input = QLineEdit()
        self.search_button = QPushButton('검색')
        self.result_text = QTextEdit()
        self.ppt_button = QPushButton('PPT 생성')

        layout.addWidget(self.search_input)
        layout.addWidget(self.search_button)
        layout.addWidget(self.result_text)
        layout.addWidget(self.ppt_button)

        self.search_button.clicked.connect(self.search_bible)
        self.ppt_button.clicked.connect(self.generate_ppt)

    def search_bible(self):
        search_query = self.search_input.text()
        book, chapter, start_verse, end_verse = parse_query(search_query)

        if book and chapter and start_verse:
            start_key = f"{book}{chapter}:{start_verse}"
            end_key = f"{book}{chapter}:{end_verse}" if end_verse else start_key
            
            verses = {}
            for k, v in bible_data.items():
                if k.startswith(f"{book}{chapter}:"):
                    verse_num = int(k.split(':')[1])
                    if int(start_verse) <= verse_num <= (int(end_verse) if end_verse else int(start_verse)):
                        verses[k] = v
            
            if verses:
                book_name = next((k for k, v in book_abbreviations.items() if v == book), book)
                result = f"{book_name} {chapter}장\n\n"
                
                for k, v in sorted(verses.items(), key=lambda x: int(x[0].split(':')[1])):
                    verse_num = k.split(':')[1]
                    result += f"{verse_num}: {v}\n"
                
                self.result_text.setText(result)
                self.ppt_file = create_ppt(book_name, chapter, f"{start_verse}-{end_verse}", result)
            else:
                self.result_text.setText("해당 구절을 찾을 수 없습니다.")
        else:
            self.result_text.setText("올바른 검색 형식이 아닙니다.")

    def generate_ppt(self):
        # PPT 생성 로직 구현 (이미 search_bible에서 생성됨)
        pass

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = BibleSearchApp()
    ex.show()
    sys.exit(app.exec_())

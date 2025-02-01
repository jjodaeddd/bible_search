import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLineEdit, QPushButton, QTextEdit, QVBoxLayout, QWidget
from PyQt5.QtGui import QIcon

# 기존 코드에서 필요한 함수들을 가져옵니다
from your_original_file import load_bible, parse_query, create_ppt, book_abbreviations, abbrev_to_korean

class BibleSearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.bible_data = load_bible()

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
        query = self.search_input.text()
        book, chapter, start_verse, end_verse = parse_query(query)
        
        if book and chapter and start_verse:
            book_data = next((item for item in self.bible_data if item["abbrev"] == book), None)
            if book_data:
                chapter_index = int(chapter) - 1
                if 0 <= chapter_index < len(book_data["chapters"]):
                    korean_book_name = abbrev_to_korean.get(book, book)
                    start_verse_index = int(start_verse) - 1
                    end_verse_index = int(end_verse) - 1 if end_verse else start_verse_index
                    
                    if 0 <= start_verse_index <= end_verse_index < len(book_data["chapters"][chapter_index]):
                        verses = book_data["chapters"][chapter_index][start_verse_index:end_verse_index+1]
                        result = f"{korean_book_name} {chapter}:{start_verse}"
                        if end_verse:
                            result += f"-{end_verse}"
                        result += ":\n\n"
                        for i, verse in enumerate(verses, start=int(start_verse)):
                            result += f"{i}절: {verse}\n"
                        self.result_text.setText(result)
                    else:
                        self.result_text.setText("해당 절 범위를 찾을 수 없습니다.")
                else:
                    self.result_text.setText("해당 장을 찾을 수 없습니다.")
            else:
                self.result_text.setText("해당 책을 찾을 수 없습니다.")
        else:
            self.result_text.setText("올바른 검색 형식이 아닙니다.")

    def generate_ppt(self):
        # PPT 생성 로직 구현
        # create_ppt 함수를 호출하여 PPT 파일 생성
        pass

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = BibleSearchApp()
    ex.show()
    sys.exit(app.exec_())

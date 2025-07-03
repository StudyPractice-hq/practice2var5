import sys
import os
import shutil
import sqlite3
import pandas as pd

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLineEdit, QLabel, QTextEdit, QTableWidget, QTableWidgetItem,
    QFileDialog, QMessageBox, QHeaderView, QDialog, QScrollArea
)
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt

import fitz  # PyMuPDF
import matplotlib.pyplot as plt

DB_NAME = 'books.db'
PDF_DIR = 'pdfs'


class PDFViewer(QDialog):
    def __init__(self, pdf_path):
        super().__init__()
        self.setWindowTitle("Чтение книги")
        self.setMinimumSize(800, 1000)

        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.total_pages = len(self.doc)
        self.current_page = 0

        self.layout = QVBoxLayout()

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.image_label)

        self.nav_layout = QHBoxLayout()
        self.prev_btn = QPushButton("← Назад")
        self.next_btn = QPushButton("Вперёд →")
        self.page_info = QLabel()

        self.prev_btn.clicked.connect(self.show_prev_page)
        self.next_btn.clicked.connect(self.show_next_page)

        self.nav_layout.addWidget(self.prev_btn)
        self.nav_layout.addWidget(self.page_info)
        self.nav_layout.addWidget(self.next_btn)

        self.layout.addWidget(self.scroll_area)
        self.layout.addLayout(self.nav_layout)
        self.setLayout(self.layout)

        self.render_page()

    def render_page(self):
        page = self.doc.load_page(self.current_page)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(img)
        self.image_label.setPixmap(pixmap)
        self.page_info.setText(f"Страница {self.current_page + 1} / {self.total_pages}")
        self.prev_btn.setEnabled(self.current_page > 0)
        self.next_btn.setEnabled(self.current_page < self.total_pages - 1)

    def show_prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.render_page()

    def show_next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.render_page()


class LibraryApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Магазин книг")
        self.setGeometry(200, 200, 1000, 700)

        if not os.path.exists(PDF_DIR):
            os.makedirs(PDF_DIR)

        self.conn = sqlite3.connect(DB_NAME)
        self.init_db()

        self.create_ui()
        self.load_books()

    def init_db(self):
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS books (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                author TEXT NOT NULL,
                price REAL,
                description TEXT,
                pdf_path TEXT,
                quantity INTEGER DEFAULT 0
            )
        ''')
        try:
            cursor.execute("ALTER TABLE books ADD COLUMN quantity INTEGER DEFAULT 0")
        except sqlite3.OperationalError:
            pass

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                book_id INTEGER,
                date TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        self.conn.commit()

    def create_ui(self):
        layout = QVBoxLayout()

        self.title_input = QLineEdit()
        self.author_input = QLineEdit()
        self.price_input = QLineEdit()
        self.quantity_input = QLineEdit()
        self.desc_input = QTextEdit()
        self.pdf_path = None

        layout.addWidget(QLabel("Название:"))
        layout.addWidget(self.title_input)
        layout.addWidget(QLabel("Автор:"))
        layout.addWidget(self.author_input)
        layout.addWidget(QLabel("Цена:"))
        layout.addWidget(self.price_input)
        layout.addWidget(QLabel("Количество:"))
        layout.addWidget(self.quantity_input)
        layout.addWidget(QLabel("Описание:"))
        layout.addWidget(self.desc_input)

        select_pdf_btn = QPushButton("Выбрать PDF")
        select_pdf_btn.clicked.connect(self.select_pdf)
        layout.addWidget(select_pdf_btn)

        add_btn = QPushButton("Добавить книгу")
        add_btn.clicked.connect(self.add_book)

        edit_btn = QPushButton("Редактировать")
        edit_btn.clicked.connect(self.edit_book)

        delete_btn = QPushButton("Удалить книгу")
        delete_btn.clicked.connect(self.delete_book)

        sell_btn = QPushButton("Продать книгу")
        sell_btn.clicked.connect(self.sell_book)

        open_btn = QPushButton("Открыть PDF")
        open_btn.clicked.connect(self.open_pdf_internal)

        stats_btn = QPushButton("Статистика продаж")
        stats_btn.clicked.connect(self.show_statistics)

        export_btn = QPushButton("Экспорт в Excel")
        export_btn.clicked.connect(self.export_statistics)
        self.export_btn = export_btn

        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по названию или автору")
        self.search_input.textChanged.connect(self.load_books)
        filter_layout.addWidget(QLabel("Фильтр:"))
        filter_layout.addWidget(self.search_input)
        layout.addLayout(filter_layout)

        btn_row = QHBoxLayout()
        btn_row.addWidget(add_btn)
        btn_row.addWidget(delete_btn)
        btn_row.addWidget(sell_btn)
        btn_row.addWidget(open_btn)
        btn_row.addWidget(stats_btn)
        btn_row.addWidget(edit_btn)
        btn_row.addWidget(export_btn)
        layout.addLayout(btn_row)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["ID", "Название", "Автор", "Цена", "Описание", "Кол-во"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        self.setLayout(layout)
        self.export_btn.setEnabled(False)

    def select_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выбрать PDF-файл", "", "PDF Files (*.pdf)")
        if path:
            filename = os.path.basename(path)
            new_path = os.path.join(PDF_DIR, filename)
            shutil.copy(path, new_path)
            self.pdf_path = new_path

    def add_book(self):
        title = self.title_input.text().strip()
        author = self.author_input.text().strip()
        price = self.price_input.text().strip()
        quantity = self.quantity_input.text().strip()
        desc = self.desc_input.toPlainText().strip()

        if not title or not author or not self.pdf_path:
            QMessageBox.warning(self, "Ошибка", "Заполните все поля и выберите PDF.")
            return

        try:
            price_val = float(price) if price else 0.0
            quantity_val = int(quantity) if quantity else 0
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Цена и количество должны быть числами.")
            return

        cursor = self.conn.cursor()
        cursor.execute(
            "INSERT INTO books (title, author, price, description, pdf_path, quantity) VALUES (?, ?, ?, ?, ?, ?)",
            (title, author, price_val, desc, self.pdf_path, quantity_val)
        )
        self.conn.commit()

        self.title_input.clear()
        self.author_input.clear()
        self.price_input.clear()
        self.desc_input.clear()
        self.quantity_input.clear()
        self.pdf_path = None

        self.load_books()

    def edit_book(self):
        book_id = self.get_selected_book_id()
        if book_id is None:
            QMessageBox.warning(self, "Ошибка", "Выберите книгу для редактирования.")
            return

        title = self.title_input.text().strip()
        author = self.author_input.text().strip()
        price = self.price_input.text().strip()
        desc = self.desc_input.toPlainText().strip()
        quantity = self.quantity_input.text().strip()

        try:
            price_val = float(price) if price else 0.0
            quantity_val = int(quantity) if quantity else 0
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Цена и количество должны быть числами.")
            return

        cursor = self.conn.cursor()
        cursor.execute(
            "UPDATE books SET title=?, author=?, price=?, description=?, quantity=? WHERE id=?",
            (title, author, price_val, desc, quantity_val, book_id)
        )
        self.conn.commit()
        self.load_books()

    def load_books(self):
        filter_text = self.search_input.text().lower()
        cursor = self.conn.cursor()
        if filter_text:
            query = '''
                SELECT id, title, author, price, description, quantity
                FROM books
                WHERE LOWER(title) LIKE ? OR LOWER(author) LIKE ? OR CAST(price AS TEXT) LIKE ?
                OR CAST(quantity AS TEXT) LIKE ? OR LOWER(description) LIKE ?
            '''
            params = tuple(f'%{filter_text}%' for _ in range(5))
            cursor.execute(query, params)
        else:
            cursor.execute("SELECT id, title, author, price, description, quantity FROM books")

        rows = cursor.fetchall()
        self.table.setRowCount(len(rows))
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(value)))

    def get_selected_book_id(self):
        row = self.table.currentRow()
        if row == -1:
            return None
        return int(self.table.item(row, 0).text())

    def delete_book(self):
        book_id = self.get_selected_book_id()
        if book_id is None:
            QMessageBox.warning(self, "Ошибка", "Выберите книгу для удаления.")
            return
        confirm = QMessageBox.question(self, "Подтвердите", "Удалить книгу?", QMessageBox.Yes | QMessageBox.No)
        if confirm == QMessageBox.Yes:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM books WHERE id=?", (book_id,))
            self.conn.commit()
            self.load_books()

    def open_pdf_internal(self):
        book_id = self.get_selected_book_id()
        if book_id is None:
            QMessageBox.warning(self, "Ошибка", "Выберите книгу.")
            return

        cursor = self.conn.cursor()
        cursor.execute("SELECT pdf_path FROM books WHERE id=?", (book_id,))
        result = cursor.fetchone()
        if result and os.path.exists(result[0]):
            viewer = PDFViewer(result[0])
            viewer.exec_()
        else:
            QMessageBox.warning(self, "Ошибка", "Файл PDF не найден.")

    def sell_book(self):
        book_id = self.get_selected_book_id()
        if book_id is None:
            QMessageBox.warning(self, "Ошибка", "Выберите книгу.")
            return

        cursor = self.conn.cursor()
        cursor.execute("INSERT INTO sales (book_id) VALUES (?)", (book_id,))
        self.conn.commit()
        QMessageBox.information(self, "Продажа", "Книга успешно продана.")

    def show_statistics(self):
        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT b.title, COUNT(s.id), SUM(COALESCE(b.price, 0))
            FROM sales s
            JOIN books b ON s.book_id = b.id
            GROUP BY s.book_id
        ''')
        data = cursor.fetchall()

        if not data:
            QMessageBox.information(self, "Статистика", "Продаж пока нет.")
            self.export_btn.setEnabled(False)
            return

        self.export_btn.setEnabled(True)
        titles, sold_counts, revenues = zip(*data)

        self.sales_df = pd.DataFrame({
            "Название книги": titles,
            "Количество продаж": sold_counts,
            "Выручка (₽)": revenues
        })

        avg_check = sum(revenues) / sum(sold_counts)

        plt.figure(figsize=(10, 6))

        plt.subplot(2, 1, 1)
        plt.bar(titles, revenues)
        plt.title("Выручка по книгам")
        plt.ylabel("₽")
        plt.xticks(rotation=45, ha='right')

        plt.subplot(2, 1, 2)
        plt.bar(titles, sold_counts)
        plt.title("Количество проданных экземпляров")
        plt.ylabel("Шт.")
        plt.xticks(rotation=45, ha='right')

        plt.tight_layout()
        plt.suptitle(f"📈 Средний чек: {avg_check:.2f} ₽", fontsize=10, y=1.03)
        plt.show(block=False)

    def export_statistics(self):
        if not hasattr(self, 'sales_df') or self.sales_df.empty:
            QMessageBox.information(self, "Экспорт", "Нет данных для экспорта.")
            return

        path, _ = QFileDialog.getSaveFileName(self, "Сохранить Excel-файл", "sales_report.xlsx", "Excel Files (*.xlsx)")
        if path:
            try:
                self.sales_df.to_excel(path, index=False)
                QMessageBox.information(self, "Экспорт", f"Файл успешно сохранён:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл:\n{e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LibraryApp()
    window.show()
    sys.exit(app.exec_())

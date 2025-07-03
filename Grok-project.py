import sys
import os
import sqlite3
from datetime import datetime
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QTableWidget, QTableWidgetItem, QLineEdit, QPushButton,
                             QFormLayout, QFileDialog, QMessageBox, QHeaderView, QTabWidget,
                             QDialog, QScrollArea, QLabel)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QImage
import fitz  # PyMuPDF для рендеринга PDF
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

class LibraryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bookstore Management System")
        self.setGeometry(100, 100, 1000, 600)

        # Инициализация базы данных
        self.conn = sqlite3.connect('library.db')
        self.create_tables()

        # Основной виджет с вкладками
        self.main_widget = QTabWidget()
        self.setCentralWidget(self.main_widget)

        # Вкладка "Книги"
        self.books_widget = QWidget()
        self.books_layout = QVBoxLayout(self.books_widget)

        # Форма для добавления книг
        self.form_layout = QFormLayout()
        self.title_input = QLineEdit()
        self.author_input = QLineEdit()
        self.price_input = QLineEdit()
        self.quantity_input = QLineEdit()
        self.description_input = QLineEdit()
        self.pdf_path_input = QLineEdit()
        self.pdf_path_input.setReadOnly(True)
        self.browse_button = QPushButton("Browse PDF")
        self.browse_button.clicked.connect(self.browse_pdf)
        self.add_button = QPushButton("Add Book")
        self.add_button.clicked.connect(self.add_book)
        self.edit_button = QPushButton("Edit Selected Book")
        self.edit_button.clicked.connect(self.edit_book)
        self.delete_button = QPushButton("Delete Selected Book")
        self.delete_button.clicked.connect(self.delete_book)
        self.sell_button = QPushButton("Sell Selected Book")
        self.sell_button.clicked.connect(self.sell_book)

        self.form_layout.addRow("Title:", self.title_input)
        self.form_layout.addRow("Author:", self.author_input)
        self.form_layout.addRow("Price:", self.price_input)
        self.form_layout.addRow("Quantity:", self.quantity_input)
        self.form_layout.addRow("Description:", self.description_input)
        self.form_layout.addRow("PDF Path:", self.pdf_path_input)
        self.form_layout.addRow("", self.browse_button)
        self.form_layout.addRow("", self.add_button)
        self.form_layout.addRow("", self.edit_button)
        self.form_layout.addRow("", self.delete_button)
        self.form_layout.addRow("", self.sell_button)

        # Поле универсального фильтра
        self.filter_layout = QHBoxLayout()
        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText("Filter by title, author, price, quantity, or description")
        self.filter_input.textChanged.connect(self.filter_books)
        self.filter_layout.addWidget(self.filter_input)

        # Таблица для отображения книг
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["ID", "Title", "Author", "Price", "Quantity", "Description"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.doubleClicked.connect(self.open_pdf)

        # Добавление элементов в layout вкладки "Книги"
        self.books_layout.addLayout(self.form_layout)
        self.books_layout.addLayout(self.filter_layout)
        self.books_layout.addWidget(self.table)

        # Вкладка "Статистика"
        self.stats_widget = QWidget()
        self.stats_layout = QVBoxLayout(self.stats_widget)
        self.stats_button = QPushButton("Show Sales Statistics")
        self.stats_button.clicked.connect(self.show_statistics)
        self.export_button = QPushButton("Export Statistics to Excel")
        self.export_button.clicked.connect(self.export_to_excel)
        self.stats_layout.addWidget(self.stats_button)
        self.stats_layout.addWidget(self.export_button)
        self.stats_canvas = None

        # Изначально отключение кнопки экспорта
        self.update_export_button_state()

        # Добавление вкладок
        self.main_widget.addTab(self.books_widget, "Books")
        self.main_widget.addTab(self.stats_widget, "Statistics")

        # Загрузка начальных данных
        self.load_books()

    def create_tables(self):
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS books (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                author TEXT NOT NULL,
                price REAL,
                quantity INTEGER,
                description TEXT,
                pdf_path TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                book_id INTEGER,
                sale_date TEXT,
                amount REAL,
                FOREIGN KEY (book_id) REFERENCES books(id)
            )
        ''')
        self.conn.commit()

    def add_book(self):
        title = self.title_input.text()
        author = self.author_input.text()
        price = self.price_input.text()
        quantity = self.quantity_input.text()
        description = self.description_input.text()
        pdf_path = self.pdf_path_input.text()

        if not title or not author or not pdf_path:
            QMessageBox.warning(self, "Input Error", "Title, Author, and PDF Path are required!")
            return

        try:
            price = float(price) if price else 0.0
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Price must be a valid number!")
            return

        try:
            quantity = int(quantity) if quantity else 0
            if quantity < 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Quantity must be a non-negative integer!")
            return

        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO books (title, author, price, quantity, description, pdf_path)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (title, author, price, quantity, description, pdf_path))
        self.conn.commit()

        self.title_input.clear()
        self.author_input.clear()
        self.price_input.clear()
        self.quantity_input.clear()
        self.description_input.clear()
        self.pdf_path_input.clear()

        self.load_books()
        self.update_export_button_state()

    def edit_book(self):
        selected = self.table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select a book to edit!")
            return

        book_id = self.table.item(self.table.currentRow(), 0).text()
        cursor = self.conn.cursor()
        cursor.execute("SELECT title, author, price, quantity, description, pdf_path FROM books WHERE id = ?", (book_id,))
        book_data = cursor.fetchone()

        edit_dialog = QDialog(self)
        edit_dialog.setWindowTitle("Edit Book")
        edit_dialog.setGeometry(200, 200, 400, 300)
        edit_layout = QFormLayout(edit_dialog)

        title_edit = QLineEdit(book_data[0])
        author_edit = QLineEdit(book_data[1])
        price_edit = QLineEdit(str(book_data[2]))
        quantity_edit = QLineEdit(str(book_data[3]))
        description_edit = QLineEdit(book_data[4])
        pdf_path_edit = QLineEdit(book_data[5])
        pdf_path_edit.setReadOnly(True)
        browse_edit_button = QPushButton("Browse PDF")
        browse_edit_button.clicked.connect(lambda: self.browse_pdf(pdf_path_edit))

        edit_layout.addRow("Title:", title_edit)
        edit_layout.addRow("Author:", author_edit)
        edit_layout.addRow("Price:", price_edit)
        edit_layout.addRow("Quantity:", quantity_edit)
        edit_layout.addRow("Description:", description_edit)
        edit_layout.addRow("PDF Path:", pdf_path_edit)
        edit_layout.addRow("", browse_edit_button)

        save_button = QPushButton("Save Changes")
        save_button.clicked.connect(lambda: self.save_book_changes(book_id, title_edit.text(), author_edit.text(),
                                                                 price_edit.text(), quantity_edit.text(),
                                                                 description_edit.text(), pdf_path_edit.text(),
                                                                 edit_dialog))
        edit_layout.addRow("", save_button)

        edit_dialog.exec_()

    def save_book_changes(self, book_id, title, author, price, quantity, description, pdf_path, dialog):
        if not title or not author or not pdf_path:
            QMessageBox.warning(self, "Input Error", "Title, Author, and PDF Path are required!")
            return

        try:
            price = float(price) if price else 0.0
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Price must be a valid number!")
            return

        try:
            quantity = int(quantity) if quantity else 0
            if quantity < 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "Input Error", "Quantity must be a non-negative integer!")
            return

        cursor = self.conn.cursor()
        cursor.execute('''
            UPDATE books
            SET title = ?, author = ?, price = ?, quantity = ?, description = ?, pdf_path = ?
            WHERE id = ?
        ''', (title, author, price, quantity, description, pdf_path, book_id))
        self.conn.commit()
        self.load_books()
        self.update_export_button_state()
        dialog.accept()

    def delete_book(self):
        selected = self.table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select a book to delete!")
            return

        book_id = self.table.item(self.table.currentRow(), 0).text()
        reply = QMessageBox.question(self, "Confirm Deletion",
                                   f"Are you sure you want to delete book ID {book_id}?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM books WHERE id = ?", (book_id,))
            cursor.execute("DELETE FROM sales WHERE book_id = ?", (book_id,))
            self.conn.commit()
            self.load_books()
            self.update_export_button_state()

    def sell_book(self):
        selected = self.table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select a book to sell!")
            return

        book_id = self.table.item(self.table.currentRow(), 0).text()
        quantity = int(self.table.item(self.table.currentRow(), 4).text())
        if quantity <= 0:
            QMessageBox.warning(self, "Stock Error", "No copies available to sell!")
            return

        price = float(self.table.item(self.table.currentRow(), 3).text())
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO sales (book_id, sale_date, amount)
            VALUES (?, ?, ?)
        ''', (book_id, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), price))
        cursor.execute("UPDATE books SET quantity = quantity - 1 WHERE id = ?", (book_id,))
        self.conn.commit()
        QMessageBox.information(self, "Success", f"Book ID {book_id} sold for {price}!")
        self.load_books()
        self.update_export_button_state()

    def browse_pdf(self, input_field=None):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf)")
        if file_path:
            if input_field:
                input_field.setText(file_path)
            else:
                self.pdf_path_input.setText(file_path)

    def load_books(self, filter_text=""):
        cursor = self.conn.cursor()
        query = "SELECT id, title, author, price, quantity, description FROM books"
        params = []
        if filter_text:
            query += """
                WHERE title LIKE ? 
                OR author LIKE ? 
                OR description LIKE ?
                OR CAST(price AS TEXT) LIKE ?
                OR CAST(quantity AS TEXT) LIKE ?
            """
            params = [f'%{filter_text}%'] * 5

        cursor.execute(query, params)
        rows = cursor.fetchall()
        self.table.setRowCount(len(rows))

        for row_idx, row_data in enumerate(rows):
            for col_idx, data in enumerate(row_data):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(data)))

    def filter_books(self):
        filter_text = self.filter_input.text()
        self.load_books(filter_text)

    def open_pdf(self, index):
        row = index.row()
        cursor = self.conn.cursor()
        cursor.execute("SELECT pdf_path, title FROM books WHERE id = ?", (self.table.item(row, 0).text(),))
        result = cursor.fetchone()
        pdf_path, title = result

        if pdf_path and os.path.exists(pdf_path):
            try:
                pdf_document = fitz.open(pdf_path)
                pdf_dialog = QDialog(self)
                pdf_dialog.setWindowTitle(f"PDF Viewer - {title}")
                pdf_dialog.setGeometry(150, 150, 600, 400)
                layout = QVBoxLayout(pdf_dialog)

                scroll_area = QScrollArea()
                scroll_area.setWidgetResizable(True)
                container = QWidget()
                container_layout = QVBoxLayout(container)

                for page_num in range(pdf_document.page_count):
                    page = pdf_document.load_page(page_num)
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
                    pixmap = QPixmap.fromImage(img)
                    label = QLabel()
                    label.setPixmap(pixmap)
                    container_layout.addWidget(label)

                scroll_area.setWidget(container)
                layout.addWidget(scroll_area)
                pdf_dialog.exec_()
                pdf_document.close()
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open PDF: {str(e)}")
        else:
            QMessageBox.warning(self, "Error", "PDF file not found!")

    def show_statistics(self):
        button_index = self.stats_layout.indexOf(self.stats_button)
        export_button_index = self.stats_layout.indexOf(self.export_button)
        for i in reversed(range(self.stats_layout.count())):
            if i != button_index and i != export_button_index:
                widget = self.stats_layout.itemAt(i).widget()
                if widget:
                    widget.deleteLater()

        cursor = self.conn.cursor()
        cursor.execute("SELECT SUM(amount), COUNT(id) FROM sales")
        result = cursor.fetchone()
        total_revenue, total_sales = result if result else (0.0, 0)
        avg_check = total_revenue / total_sales if total_sales > 0 else 0.0

        cursor.execute('''
            SELECT b.title, COUNT(s.id) as sales_count, SUM(s.amount) as total_revenue
            FROM books b
            LEFT JOIN sales s ON b.id = s.book_id
            GROUP BY b.id
            ORDER BY sales_count DESC
            LIMIT 5
        ''')
        stats = cursor.fetchall()

        if total_sales == 0:
            QMessageBox.information(self, "Statistics", "No sales data available.")
            return

        if not stats or all(sales_count == 0 for _, sales_count, _ in stats):
            QMessageBox.information(self, "Statistics", "No sales data available for graphing.")
            return

        titles = [row[0] for row in stats]
        sales_counts = [row[1] for row in stats]
        revenues = [row[2] or 0 for row in stats]

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
        ax1.bar(titles, sales_counts)
        ax1.set_title("Top 5 Books Sold")
        ax1.set_xlabel("Book Title")
        ax1.set_ylabel("Number of Sales")
        ax1.tick_params(axis='x', rotation=45)

        ax2.pie(revenues, labels=titles, autopct='%1.1f%%')
        ax2.set_title("Revenue Share by Book")

        plt.tight_layout()
        self.stats_canvas = FigureCanvas(fig)
        self.stats_layout.addWidget(self.stats_canvas)

        summary = (
            f"Total Revenue: ${total_revenue:.2f}\n"
            f"Total Sales: {total_sales}\n"
            f"Average Check: ${avg_check:.2f}\n"
            f"Top Book: {titles[0] if titles else 'N/A'}"
        )
        summary_label = QLabel(summary)
        summary_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        self.stats_layout.addWidget(summary_label)

    def update_export_button_state(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT COUNT(id) FROM sales")
        total_sales = cursor.fetchone()[0]
        self.export_button.setEnabled(total_sales > 0)

    def export_to_excel(self):
        if not self.export_button.isEnabled():
            QMessageBox.warning(self, "Export Error", "No sales data available to export.")
            return

        cursor = self.conn.cursor()
        cursor.execute('''
            SELECT b.title, s.sale_date, s.amount
            FROM sales s
            JOIN books b ON s.book_id = b.id
            ORDER BY s.sale_date DESC
        ''')
        sales_data = cursor.fetchall()

        if not sales_data:
            QMessageBox.warning(self, "Export Error", "No sales data available to export.")
            return

        df = pd.DataFrame(sales_data, columns=['Title', 'Sale Date', 'Amount'])
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel File", f"sales_statistics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", "Excel Files (*.xlsx)")
        if file_path:
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "Success", f"Statistics exported to {file_path}")

    def closeEvent(self, event):
        self.conn.close()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = LibraryApp()
    window.show()
    sys.exit(app.exec_())
import sys
import sqlite3
import os
import shutil
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
                             QFileDialog, QMessageBox, QAbstractItemView, QTabWidget,
                             QTextEdit, QDialog, QFormLayout, QSpinBox, QComboBox,
                             QGroupBox, QGridLayout, QHeaderView)
from PyQt5.QtCore import Qt
import fitz  # PyMuPDF
from PyQt5.QtGui import QPixmap, QImage
import pandas as pd


class PDFViewer(QDialog):
    def __init__(self, pdf_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Просмотр PDF")
        self.setGeometry(200, 200, 800, 900)
        self.pdf_path = pdf_path
        self.current_page = 0
        self.doc = None
        self.total_pages = 0

        self.init_ui()
        self.load_pdf()

    def init_ui(self):
        layout = QVBoxLayout()

        # Навигация
        nav_layout = QHBoxLayout()
        self.prev_btn = QPushButton("← Назад")
        self.prev_btn.clicked.connect(self.prev_page)
        nav_layout.addWidget(self.prev_btn)

        self.next_btn = QPushButton("Вперед →")
        self.next_btn.clicked.connect(self.next_page)
        nav_layout.addWidget(self.next_btn)

        self.page_label = QLabel("Страница: 0/0")
        nav_layout.addWidget(self.page_label)
        layout.addLayout(nav_layout)

        # Просмотр PDF
        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        layout.addWidget(self.view)

        self.setLayout(layout)

    def load_pdf(self):
        try:
            self.doc = fitz.open(self.pdf_path)
            self.total_pages = len(self.doc)
            self.display_page(0)
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить PDF: {str(e)}")

    def display_page(self, page_num):
        if not self.doc or page_num < 0 or page_num >= self.total_pages:
            return

        page = self.doc.load_page(page_num)
        pix = page.get_pixmap()
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(img)

        self.scene.clear()
        self.scene.addPixmap(pixmap)
        self.view.fitInView(self.scene.sceneRect(), Qt.KeepAspectRatio)

        self.current_page = page_num
        self.page_label.setText(f"Страница: {self.current_page + 1}/{self.total_pages}")
        self.update_buttons()

    def update_buttons(self):
        self.prev_btn.setEnabled(self.current_page > 0)
        self.next_btn.setEnabled(self.current_page < self.total_pages - 1)

    def prev_page(self):
        if self.current_page > 0:
            self.display_page(self.current_page - 1)

    def next_page(self):
        if self.current_page < self.total_pages - 1:
            self.display_page(self.current_page + 1)


class BookEditor(QDialog):
    def __init__(self, book_data=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Добавить/Редактировать книгу" if book_data else "Добавить книгу")
        self.setGeometry(300, 300, 500, 400)
        self.book_data = book_data

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        form = QFormLayout()

        self.title_input = QLineEdit()
        form.addRow("Название:", self.title_input)

        self.author_input = QLineEdit()
        form.addRow("Автор:", self.author_input)

        self.price_input = QLineEdit()
        form.addRow("Цена:", self.price_input)

        self.quantity_input = QSpinBox()
        self.quantity_input.setRange(0, 10000)
        form.addRow("Количество:", self.quantity_input)

        self.desc_input = QTextEdit()
        form.addRow("Описание:", self.desc_input)

        # PDF файл
        pdf_layout = QHBoxLayout()
        self.pdf_path_input = QLineEdit()
        self.pdf_path_input.setReadOnly(True)
        pdf_layout.addWidget(self.pdf_path_input)

        browse_btn = QPushButton("Выбрать")
        browse_btn.clicked.connect(self.browse_pdf)
        pdf_layout.addWidget(browse_btn)

        form.addRow("PDF файл:", pdf_layout)

        if self.book_data:
            self.title_input.setText(self.book_data.get('title', ''))
            self.author_input.setText(self.book_data.get('author', ''))
            self.price_input.setText(str(self.book_data.get('price', '')))
            self.quantity_input.setValue(self.book_data.get('quantity', 0))
            self.desc_input.setPlainText(self.book_data.get('description', ''))
            self.pdf_path_input.setText(self.book_data.get('pdf_path', ''))

        layout.addLayout(form)

        # Кнопки
        btn_layout = QHBoxLayout()

        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.accept)
        btn_layout.addWidget(save_btn)

        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def browse_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите PDF", "", "PDF Files (*.pdf)")
        if path:
            self.pdf_path_input.setText(path)

    def get_data(self):
        return {
            'title': self.title_input.text().strip(),
            'author': self.author_input.text().strip(),
            'price': self.price_input.text().strip(),
            'quantity': self.quantity_input.value(),
            'description': self.desc_input.toPlainText().strip(),
            'pdf_path': self.pdf_path_input.text().strip()
        }


class BookStoreApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Управление книжным магазином")
        self.setGeometry(100, 100, 1200, 800)

        self.init_db()
        self.init_ui()
        self.load_books()

    def init_db(self):
        self.conn = sqlite3.connect('bookstore.db')
        self.cursor = self.conn.cursor()

        # Создаем таблицы, если они не существуют
        self.cursor.execute("""
                            CREATE TABLE IF NOT EXISTS books
                            (
                                id
                                INTEGER
                                PRIMARY
                                KEY
                                AUTOINCREMENT,
                                title
                                TEXT
                                NOT
                                NULL,
                                author
                                TEXT
                                NOT
                                NULL,
                                price
                                REAL
                                NOT
                                NULL,
                                description
                                TEXT,
                                pdf_path
                                TEXT,
                                quantity
                                INTEGER
                                DEFAULT
                                0,
                                added_date
                                TEXT
                                DEFAULT
                                CURRENT_TIMESTAMP
                            )
                            """)

        self.cursor.execute("""
                            CREATE TABLE IF NOT EXISTS sales
                            (
                                id
                                INTEGER
                                PRIMARY
                                KEY
                                AUTOINCREMENT,
                                book_id
                                INTEGER,
                                book_title
                                TEXT,
                                date
                                TEXT,
                                quantity
                                INTEGER,
                                price
                                REAL,
                                total
                                REAL,
                                FOREIGN
                                KEY
                            (
                                book_id
                            ) REFERENCES books
                            (
                                id
                            )
                                )
                            """)

        self.conn.commit()

        if not os.path.exists('book_pdfs'):
            os.makedirs('book_pdfs')

    def init_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        layout = QVBoxLayout()
        self.central_widget.setLayout(layout)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.setup_books_tab()
        self.setup_sales_tab()

    def setup_books_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Книги")

        layout = QVBoxLayout()
        tab.setLayout(layout)

        # Кнопки управления
        btn_layout = QHBoxLayout()

        add_btn = QPushButton("Добавить книгу")
        add_btn.clicked.connect(self.add_book)
        btn_layout.addWidget(add_btn)

        edit_btn = QPushButton("Редактировать")
        edit_btn.clicked.connect(self.edit_book)
        btn_layout.addWidget(edit_btn)

        del_btn = QPushButton("Удалить")
        del_btn.clicked.connect(self.delete_book)
        btn_layout.addWidget(del_btn)

        view_btn = QPushButton("Просмотр PDF")
        view_btn.clicked.connect(self.view_pdf)
        btn_layout.addWidget(view_btn)

        layout.addLayout(btn_layout)

        # Таблица книг
        self.books_table = QTableWidget()
        self.books_table.setColumnCount(7)
        self.books_table.setHorizontalHeaderLabels(
            ["ID", "Название", "Автор", "Цена", "Кол-во", "Дата добавления", "PDF"]
        )
        self.books_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.books_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.books_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.books_table)

    def setup_sales_tab(self):
        tab = QWidget()
        self.tabs.addTab(tab, "Продажи")

        layout = QVBoxLayout()
        tab.setLayout(layout)

        # Форма продажи
        form = QFormLayout()

        self.sale_combo = QComboBox()
        form.addRow("Книга:", self.sale_combo)

        self.sale_qty = QSpinBox()
        self.sale_qty.setRange(1, 100)
        form.addRow("Количество:", self.sale_qty)

        sell_btn = QPushButton("Оформить продажу")
        sell_btn.clicked.connect(self.sell_book)
        form.addRow(sell_btn)

        layout.addLayout(form)

        # Таблица продаж
        self.sales_table = QTableWidget()
        self.sales_table.setColumnCount(6)
        self.sales_table.setHorizontalHeaderLabels(
            ["Дата", "ID книги", "Название", "Цена", "Кол-во", "Сумма"]
        )
        self.sales_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.sales_table)

        # Кнопка статистики
        stats_btn = QPushButton("Статистика продаж")
        stats_btn.clicked.connect(self.show_stats)
        layout.addWidget(stats_btn)

    def load_books(self):
        try:
            self.cursor.execute("""
                                SELECT id,
                                       title,
                                       author,
                                       price,
                                       quantity,
                                       strftime('%d.%m.%Y', added_date) as added_date,
                                       pdf_path
                                FROM books
                                ORDER BY title
                                """)
            books = self.cursor.fetchall()

            self.books_table.setRowCount(len(books))
            for row_idx, book in enumerate(books):
                for col_idx, value in enumerate(book):
                    item = QTableWidgetItem(str(value))
                    self.books_table.setItem(row_idx, col_idx, item)

            self.update_sales_combo()

        except sqlite3.Error as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить книги: {str(e)}")

    def update_sales_combo(self):
        self.sale_combo.clear()
        self.cursor.execute("""
                            SELECT id, title, quantity, price
                            FROM books
                            WHERE quantity > 0
                            ORDER BY title
                            """)

        for book in self.cursor.fetchall():
            self.sale_combo.addItem(
                f"{book[1]} (ID: {book[0]}, {book[2]} шт., {book[3]} руб.)",
                book[0]
            )

    def add_book(self):
        dialog = BookEditor()
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()

            if not data['title'] or not data['author']:
                QMessageBox.warning(self, "Ошибка", "Заполните название и автора!")
                return

            try:
                price = float(data['price'])
                if price <= 0:
                    raise ValueError
            except ValueError:
                QMessageBox.warning(self, "Ошибка", "Некорректная цена!")
                return

            # Копируем PDF
            new_pdf_path = ""
            if data['pdf_path']:
                try:
                    filename = f"{data['title'][:50]}_{data['author'][:50]}.pdf"
                    filename = "".join(c for c in filename if c.isalnum() or c in (' ', '.', '_')).rstrip()
                    new_pdf_path = os.path.join('book_pdfs', filename)
                    shutil.copy2(data['pdf_path'], new_pdf_path)
                except Exception as e:
                    QMessageBox.warning(self, "Ошибка", f"Не удалось скопировать PDF: {str(e)}")
                    return

            try:
                self.cursor.execute("""
                                    INSERT INTO books
                                        (title, author, price, description, pdf_path, quantity)
                                    VALUES (?, ?, ?, ?, ?, ?)
                                    """, (data['title'], data['author'], price,
                                          data['description'], new_pdf_path, data['quantity']))

                self.conn.commit()
                self.load_books()
                QMessageBox.information(self, "Успех", "Книга успешно добавлена!")

            except sqlite3.Error as e:
                QMessageBox.warning(self, "Ошибка", f"Ошибка базы данных: {str(e)}")

    def edit_book(self):
        selected = self.books_table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Ошибка", "Выберите книгу для редактирования!")
            return

        book_id = self.books_table.item(selected[0].row(), 0).text()

        self.cursor.execute("""
                            SELECT title, author, price, quantity, description, pdf_path
                            FROM books
                            WHERE id = ?
                            """, (book_id,))

        book = self.cursor.fetchone()
        if not book:
            QMessageBox.warning(self, "Ошибка", "Книга не найдена!")
            return

        book_data = {
            'title': book[0],
            'author': book[1],
            'price': book[2],
            'quantity': book[3],
            'description': book[4],
            'pdf_path': book[5]
        }

        dialog = BookEditor(book_data)
        if dialog.exec_() == QDialog.Accepted:
            new_data = dialog.get_data()

            if not new_data['title'] or not new_data['author']:
                QMessageBox.warning(self, "Ошибка", "Заполните название и автора!")
                return

            try:
                price = float(new_data['price'])
                if price <= 0:
                    raise ValueError
            except ValueError:
                QMessageBox.warning(self, "Ошибка", "Некорректная цена!")
                return

            # Обновляем PDF если он был изменен
            new_pdf_path = book_data['pdf_path']
            if new_data['pdf_path'] and new_data['pdf_path'] != book_data['pdf_path']:
                try:
                    if os.path.exists(book_data['pdf_path']):
                        os.remove(book_data['pdf_path'])

                    filename = f"{new_data['title'][:50]}_{new_data['author'][:50]}.pdf"
                    filename = "".join(c for c in filename if c.isalnum() or c in (' ', '.', '_')).rstrip()
                    new_pdf_path = os.path.join('book_pdfs', filename)
                    shutil.copy2(new_data['pdf_path'], new_pdf_path)
                except Exception as e:
                    QMessageBox.warning(self, "Ошибка", f"Не удалось обновить PDF: {str(e)}")
                    return

            try:
                self.cursor.execute("""
                                    UPDATE books
                                    SET title       = ?,
                                        author      = ?,
                                        price       = ?,
                                        quantity    = ?,
                                        description = ?,
                                        pdf_path    = ?
                                    WHERE id = ?
                                    """, (new_data['title'], new_data['author'], price,
                                          new_data['quantity'], new_data['description'],
                                          new_pdf_path, book_id))

                self.conn.commit()
                self.load_books()
                QMessageBox.information(self, "Успех", "Данные книги обновлены!")

            except sqlite3.Error as e:
                QMessageBox.warning(self, "Ошибка", f"Ошибка базы данных: {str(e)}")

    def delete_book(self):
        selected = self.books_table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Ошибка", "Выберите книгу для удаления!")
            return

        book_id = self.books_table.item(selected[0].row(), 0).text()
        book_title = self.books_table.item(selected[0].row(), 1).text()

        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Вы уверены, что хотите удалить книгу:\n{book_title} (ID: {book_id})?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            try:
                # Удаляем PDF файл
                pdf_path = self.books_table.item(selected[0].row(), 6).text()
                if pdf_path and os.path.exists(pdf_path):
                    os.remove(pdf_path)

                # Удаляем книгу из БД
                self.cursor.execute("DELETE FROM books WHERE id = ?", (book_id,))
                self.conn.commit()

                self.load_books()
                QMessageBox.information(self, "Успех", "Книга успешно удалена!")

            except sqlite3.Error as e:
                QMessageBox.warning(self, "Ошибка", f"Ошибка базы данных: {str(e)}")

    def view_pdf(self):
        selected = self.books_table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Ошибка", "Выберите книгу для просмотра!")
            return

        pdf_path = self.books_table.item(selected[0].row(), 6).text()
        if not pdf_path:
            QMessageBox.warning(self, "Ошибка", "Для этой книги нет PDF файла!")
            return

        if not os.path.exists(pdf_path):
            QMessageBox.warning(self, "Ошибка", "PDF файл не найден!")
            return

        viewer = PDFViewer(pdf_path, self)
        viewer.exec_()

    def sell_book(self):
        if self.sale_combo.count() == 0:
            QMessageBox.warning(self, "Ошибка", "Нет доступных книг для продажи!")
            return

        book_id = self.sale_combo.currentData()
        qty = self.sale_qty.value()

        # Получаем данные о книге
        self.cursor.execute("""
                            SELECT title, price, quantity
                            FROM books
                            WHERE id = ?
                            """, (book_id,))

        book = self.cursor.fetchone()
        if not book:
            QMessageBox.warning(self, "Ошибка", "Книга не найдена!")
            return

        title, price, available = book

        if qty > available:
            QMessageBox.warning(
                self, "Ошибка",
                f"Недостаточно книг в наличии! Доступно: {available}"
            )
            return

        # Оформляем продажу
        total = price * qty
        date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        try:
            # Добавляем запись о продаже
            self.cursor.execute("""
                                INSERT INTO sales
                                    (book_id, book_title, date, quantity, price, total)
                                VALUES (?, ?, ?, ?, ?, ?)
                                """, (book_id, title, date, qty, price, total))

            # Обновляем количество книг
            self.cursor.execute("""
                                UPDATE books
                                SET quantity = quantity - ?
                                WHERE id = ?
                                """, (qty, book_id))

            self.conn.commit()

            # Обновляем интерфейс
            self.load_books()
            self.load_sales()

            QMessageBox.information(
                self, "Успех",
                f"Продажа оформлена!\n{title}\n{qty} шт. × {price} руб. = {total} руб."
            )

        except sqlite3.Error as e:
            QMessageBox.warning(self, "Ошибка", f"Ошибка базы данных: {str(e)}")

    def load_sales(self):
        try:
            self.cursor.execute("""
                                SELECT strftime('%d.%m.%Y %H:%M', date) as date,
                    book_id,
                    book_title,
                    price,
                    quantity,
                    total
                                FROM sales
                                ORDER BY date DESC
                                    LIMIT 100
                                """)
            sales = self.cursor.fetchall()

            self.sales_table.setRowCount(len(sales))
            for row_idx, sale in enumerate(sales):
                for col_idx, value in enumerate(sale):
                    item = QTableWidgetItem(str(value))
                    if col_idx in (3, 5):  # Цена и сумма
                        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.sales_table.setItem(row_idx, col_idx, item)

        except sqlite3.Error as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить продажи: {str(e)}")

    def show_stats(self):
        try:
            # Получаем основные метрики
            self.cursor.execute("""
                                SELECT COUNT(*)      as total_sales,
                                       SUM(quantity) as total_books,
                                       SUM(total)    as total_amount,
                                       AVG(total)    as avg_check
                                FROM sales
                                """)
            stats = self.cursor.fetchone()

            # Получаем топ-5 книг
            self.cursor.execute("""
                                SELECT book_title, SUM(quantity) as total_qty, SUM(total) as total_sum
                                FROM sales
                                GROUP BY book_title
                                ORDER BY total_qty DESC LIMIT 5
                                """)
            top_books = self.cursor.fetchall()

            # Создаем диалог для отображения статистики
            dialog = QDialog(self)
            dialog.setWindowTitle("Статистика продаж")
            dialog.setGeometry(200, 200, 800, 600)

            layout = QVBoxLayout()

            # Основные метрики
            metrics = QGroupBox("Ключевые показатели")
            metrics_layout = QFormLayout()

            metrics_layout.addRow("Общая выручка:", QLabel(f"{stats[2]:.2f} руб."))
            metrics_layout.addRow("Продано книг:", QLabel(str(stats[1])))
            metrics_layout.addRow("Средний чек:", QLabel(f"{stats[3]:.2f} руб."))

            top_text = "\n".join([f"{i + 1}. {book[0]} - {book[1]} шт. ({book[2]:.2f} руб.)"
                                  for i, book in enumerate(top_books)])
            metrics_layout.addRow("Топ-5 книг:", QLabel(top_text if top_books else "Нет данных"))

            metrics.setLayout(metrics_layout)
            layout.addWidget(metrics)

            # Графики
            fig = plt.figure(figsize=(10, 5))
            canvas = FigureCanvas(fig)

            # График продаж по дням
            self.cursor.execute("""
                                SELECT date (date) as day, SUM (total) as daily_total
                                FROM sales
                                GROUP BY day
                                ORDER BY day
                                """)
            sales_data = self.cursor.fetchall()

            if sales_data:
                dates = [row[0] for row in sales_data]
                amounts = [row[1] for row in sales_data]

                ax = fig.add_subplot(111)
                ax.bar(dates, amounts)
                ax.set_title('Продажи по дням')
                ax.set_xlabel('Дата')
                ax.set_ylabel('Сумма (руб)')
                plt.xticks(rotation=45)
                fig.tight_layout()

            layout.addWidget(canvas)

            # Кнопка экспорта
            if stats[0] > 0:  # Если есть данные для экспорта
                export_btn = QPushButton("Экспорт в Excel")
                export_btn.clicked.connect(lambda: self.export_stats_to_excel(dialog))
                layout.addWidget(export_btn)

            dialog.setLayout(layout)
            dialog.exec_()

        except sqlite3.Error as e:
            QMessageBox.warning(self, "Ошибка", f"Ошибка загрузки статистики: {str(e)}")

    def export_stats_to_excel(self, parent):
        try:
            # Получаем все данные о продажах
            self.cursor.execute("""
                                SELECT
                                    date as "Дата продажи", book_id as "ID книги", book_title as "Название книги", price as "Цена", quantity as "Количество", total as "Сумма"
                                FROM sales
                                ORDER BY date DESC
                                """)
            sales_data = self.cursor.fetchall()
            columns = [desc[0] for desc in self.cursor.description]

            # Создаем DataFrame
            df = pd.DataFrame(sales_data, columns=columns)

            # Выбираем файл для сохранения
            file_path, _ = QFileDialog.getSaveFileName(
                parent, "Сохранить как", "Статистика_продаж.xlsx",
                "Excel Files (*.xlsx)"
            )

            if file_path:
                # Сохраняем в Excel
                writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Продажи', index=False)

                # Добавляем сводные таблицы
                workbook = writer.book
                worksheet = writer.sheets['Продажи']

                # Форматирование
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'border': 1
                })

                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, max(len(value), 12))

                writer.save()
                QMessageBox.information(parent, "Успех", "Данные успешно экспортированы")

        except Exception as e:
            QMessageBox.warning(parent, "Ошибка", f"Ошибка экспорта: {str(e)}")

    def closeEvent(self, event):
        self.conn.close()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BookStoreApp()
    window.show()
    sys.exit(app.exec_())
import sys
import sqlite3
import csv
import xlsxwriter

from PyQt6 import uic
import os
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import QApplication, QComboBox, QMessageBox
from PyQt6.QtWidgets import QMainWindow, QInputDialog, QTableWidgetItem, QFileDialog
from PIL import Image


class FirstForm(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('FirstForm.ui', self)
        # передаем цвета виджета на экране:
        self.tableWidget.setStyleSheet("background-color: #FFF8DC;")
        self.btn_add.setStyleSheet("background-color: white; color: black;")
        self.btn_delete.setStyleSheet("background-color: white; color: black;")
        self.btn_new.setStyleSheet("background-color: white; color: black;")
        self.btn_save.setStyleSheet("background-color: white; color: black;")
        # передаем кнопкам символы с помощью кодировки Unicode
        self.btn_new.setText('\U0001F5D8')
        self.btn_delete.setText('\U0001F5D1')

        # помещаем картинки на экран
        pixmap, pixmap_2 = QPixmap('image/Фон.jpg'), QPixmap('image/Круг.jpg')
        self.label_photo.setPixmap(pixmap)
        self.label.setPixmap(pixmap_2)

        self.table()  # вызываем функцию что-бы заполнить таблицу на главной экране
        # Устанавливаем ширину столцов
        self.tableWidget.setColumnWidth(0, 150)  # первый столбец
        self.tableWidget.setColumnWidth(1, 150)  # второй столбец
        self.tableWidget.setColumnWidth(2, 350)  # третий столбец

        self.tableWidget.cellClicked.connect(self.watch)  # при нажатии на ячейку таблицы вызываем функци watch
        self.btn_add.clicked.connect(self.open_add_form)  # при нажатии на кнопку открываем форму AddForm
        self.btn_delete.clicked.connect(self.open_delete_form)  # при нажатии на кнопку открываем форму DeleteForm
        self.btn_new.clicked.connect(self.table)  # при нажатии на кнопку вызываем функцию для перезагрузки таблицы
        self.btn_save.clicked.connect(self.open_save_form)

    def table(self):
        # заполняем таблицу:
        conn = sqlite3.connect('project_2.db')
        cursor = conn.cursor()

        cursor.execute('SELECT name, alias, description FROM gods')
        rows = cursor.fetchall()

        # Устанавливаем количество строк и столбцов в QTableWidget
        self.tableWidget.setRowCount(len(rows))
        self.tableWidget.setColumnCount(3)  # Количество столбцов

        # Устанавливаем заголовки столбцов
        self.tableWidget.setHorizontalHeaderLabels(['Name', 'Alias', 'Description'])

        # Заполняем таблицу данными
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(value)))

        conn.close()

    def watch(self, row, column):
        # получаем название столбца ячейки, на которую мы нажали
        column_name = self.tableWidget.horizontalHeaderItem(column).text()
        if column_name == 'Name':  # если этот слобец с именем существа, то вызываем форму Gods_Form
            name = self.tableWidget.item(row, column).text()
            self.gods_form = Gods_Form(name, 'просмотр')  # Передаем name в форму Gods_Form
            self.gods_form.show()
        else:  # иначе просим пользователя нажать на ячейку, находящуюся в столбце с именем
            QMessageBox.information(self, 'Ошибка', 'Пожалуйста, нажмите на имя существа, чтобы его просмотреть.')

    def open_add_form(self):  # открываем форму AddForm
        self.add_form = AddForm()
        self.add_form.show()

    def open_delete_form(self):  # открываем форму DeleteForm
        self.delete_form = DeleteForm()
        self.delete_form.show()

    def open_save_form(self):  # открываем форму SaveForm
        self.save_form = SaveForm()
        self.save_form.show()


class Gods_Form(QMainWindow):
    def __init__(self, name, do):
        super().__init__()
        uic.loadUi('watch.ui', self)
        # окрашиваем форму
        self.setStyleSheet("background-color: #FFDAB9;")
        self.textEdit_name.setStyleSheet("background-color: white; color: black;")
        self.textEdit_alias.setStyleSheet("background-color: white; color: black;")
        self.textEdit_description.setStyleSheet("background-color: white; color: black;")
        self.textEdit_ability.setStyleSheet("background-color: white; color: black;")
        self.textEdit_time_of_stay.setStyleSheet("background-color: white; color: black;")
        self.textEdit_book.setStyleSheet("background-color: white; color: black;")
        self.btn_change.setStyleSheet("background-color: white; color: black;")
        # добавляем картинку
        pixmap = QPixmap('image/Фонарь.jpg')
        self.label.setPixmap(pixmap)

        self.load_god_info(name)  # заполняем информацией экран
        if do == 'просмотр':  # если мы пришли сюда из главной формы, при нажатии на кнопку изменить, вызываем функцию
            self.btn_change.clicked.connect(self.open_change_form)  # открываем форму при нажатии на кнопку изменить
        else:  # если пришли из фукции добавления, то рячем кнопку изменить
            self.btn_change.hide()

    def load_god_info(self, name):
        conn = sqlite3.connect('project_2.db')
        cursor = conn.cursor()

        try:
            cursor.execute('''SELECT gods.name AS god_name, gods.alias, gods.description, gods.ability, 
            gods.time_of_stay, gods.image, books.name AS book_name FROM gods_books
                    JOIN gods ON gods_books.god_id = gods.ID
                    JOIN books ON gods_books.book_id = books.ID
                     WHERE gods.name = ?''', (name,))
            r = cursor.fetchone()  # Получаем запись о существе

            if r:  # Проверяем, есть ли результат
                # Устанавливаем текст в текстовое поле
                self.textEdit_name.setPlainText(r[0])  # name
                self.textEdit_alias.setPlainText(r[1])  # alias
                self.textEdit_description.setPlainText(r[2])  # description
                self.textEdit_ability.setPlainText(r[3])  # ability
                self.textEdit_time_of_stay.setPlainText(r[4])  # time_of_stay
                self.textEdit_book.setPlainText(r[6])  # book
                if r[5]:
                    im = Image.open(r[5])  # image
                    im2 = im.resize((200, 150))  # сжимаем картинку под размер виджета QLabel
                    im2.save('picture.jpg')  # сохраняем получившуюся картинку
                    pixmap = QPixmap('picture.jpg')
                    self.label_image.setPixmap(pixmap)  # показываем на экран
                    os.remove('picture.jpg')  # удаляем сжатую картинку
            else:
                self.textEdit_name.setPlainText("Существо не найдено.")

        except sqlite3.Error as e:
            # Обработка ошибок базы данных
            self.textEdit_name.setPlainText(f"Ошибка базы данных: {e}")

        finally:
            conn.close()  # Закрываем соединение

    def open_change_form(self):  # открываем форму ChangeForm
        self.close()  # закрываем текущую форму
        self.change_form = ChangeForm(self.textEdit_name.toPlainText())
        self.change_form.show()


class AddForm(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('add.ui', self)
        # окрашиваем форму
        self.setStyleSheet("background-color: #FFF8DC;")
        self.btn_photo.setStyleSheet("background-color: white; color: black;")
        self.btn_book.setStyleSheet("background-color: white; color: black;")
        self.btn_save.setStyleSheet("background-color: white; color: black;")
        self.btn_watch.setStyleSheet("background-color: white; color: black;")
        self.btn_new.setText('\U0001F5D8')
        self.btn_new.setStyleSheet("background-color: white; color: black;")

        # размещаем виджет для выбора книжки
        self.combo_box = QComboBox(self)
        self.combo_box.setGeometry(130, 570, 298, 20)
        self.reboot()

        self.combo_box.setStyleSheet("background-color: white; color: black;")  # делаем белым виджет QComboBox

        self.image = None

        self.btn_save.clicked.connect(self.save)  # вызываем функцию save и так сохраняем существо
        self.btn_watch.clicked.connect(self.watch)  # с помощью этой функции открываем форму с просмотром существа
        self.btn_photo.clicked.connect(self.choose_photo)  # выбираем картинку
        self.btn_book.clicked.connect(self.add_book)  # добавляем книжку
        # перезагружаем виджет QComboBox, чтобы увидеть новую книгу (если мы ее добавили)
        self.btn_new.clicked.connect(self.reboot)

    def reboot(self):
        self.combo_box.addItems(self.books())  # передаем виджету QComboBox список, полученый из функции books

    def add_book(self):
        book, ok_pressed = QInputDialog.getText(self, "Добавьте книжку", "Введите название книги:")

        if ok_pressed and book:  # Проверяем, что пользователь нажал ОК и ввел название
            conn = sqlite3.connect('project_2.db')
            cursor = conn.cursor()

            # Получаем максимальный ID в таблице с книжками
            cursor.execute('SELECT max(id) FROM books')
            r = cursor.fetchone()
            max_id_book = r[0] if r[0] is not None else 0  # Проверка на наличие результата

            new_id_book = max_id_book + 1

            cursor.execute('INSERT INTO books (ID, name) VALUES (?, ?)', (new_id_book, book))

            conn.commit()  # Подтверждаем изменения
            conn.close()  # Закрываем соединение

    def books(self):
        conn = sqlite3.connect('project_2.db')
        cursor = conn.cursor()

        cursor.execute('''
                SELECT name FROM books
                ''')
        r = cursor.fetchall()  # получаем названия всех имеющихся книжек
        books_list = [i[0] for i in r]

        return books_list  # передаем список с книжками

    def save(self):
        # получаем информацию о существе для сохранения
        name = self.textEdit_name.toPlainText()
        alias = self.textEdit_alias.toPlainText()
        description = self.textEdit_description.toPlainText()
        ability = self.textEdit_ability.toPlainText()
        time_of_stay = self.textEdit_time_of_stay.toPlainText()
        book = self.combo_box.currentText()
        image = ''  # будущий путь к картинке

        if self.image is not None:  # если мы выбрали картинку, то сохранянем ее в папку к остальным картинкам
            directory = 'image'
            # Определяем новый путь для сохранения
            image_path = os.path.join(directory, f"{name}.jpg")
            image = image_path
            # Сохраняем изображение
            self.image.save(image_path, "JPEG")

        conn = sqlite3.connect('project_2.db')
        cursor = conn.cursor()

        f = f'''
                SELECT ID FROM books
                WHERE name = '{book}'
                    '''
        cursor.execute(f)
        r_book = cursor.fetchone()
        id_book = r_book[0] if r_book[0] is not None else 0  # получаем id книги

        cursor.execute('''
                SELECT max(id) FROM gods
            ''')
        r_gods = cursor.fetchone()
        max_id = r_gods[0] if r_gods[0] is not None else 0  # Проверка на наличие результата
        new_id_god = max_id + 1  # получаем новой id для существа

        cursor.execute('''
                        SELECT max(id) FROM gods_books
                    ''')
        r_gods_books = cursor.fetchone()
        max_id_ID = r_gods_books[0] if r_gods_books[0] is not None else 0  # Проверка на наличие результата
        new_id_gods_books = max_id_ID + 1  # получаем новый id для таблицы gods_books в базу данных

        # добавляем
        c = f'''
                INSERT INTO gods (ID, name, alias, description, ability, time_of_stay, image) 
                VALUES ({new_id_god}, '{name}', '{alias}', '{description}', '{ability}', 
                '{time_of_stay}', '{image}')
            '''
        cursor.execute(c)

        a = f'''INSERT INTO gods_books (ID, god_id, book_id) 
                           VALUES ({new_id_gods_books}, {new_id_god}, {id_book})'''
        cursor.execute(a)

        conn.commit()
        conn.close()
        QMessageBox.information(self, "Информация", f"Существо успешно добавлено.")  # сообщем об успешном сохранение

    def watch(self):
        name = self.textEdit_name.toPlainText()  # получаем имя существа, которое хотим посмотреть
        self.close()  # закрываем текущее окно
        self.gods_form = Gods_Form(name, 'добавление')  # Передаем имя существа в конструктор
        self.gods_form.show()

    def choose_photo(self):
        # выбераем картинку
        fname = QFileDialog.getOpenFileName(self, 'Выбрать картинку', '',
                                            'Картинка (*.jpg);;Картинка (*.png);;Все файлы (*)')[0]
        if fname:
            self.image = Image.open(fname)  # открываем картинку для дальнейшей работы с ней
        self.btn_photo.setText('Вы выбрали картинку.')


class DeleteForm(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('delete.ui', self)
        self.setStyleSheet("background-color: #FFF8DC;")
        # окрашиваем форму
        self.textEdit.setStyleSheet("background-color: white; color: black;")
        self.tableWidget_delete.setStyleSheet("background-color: white; color: black;")
        self.pushButton_delete.setStyleSheet("background-color: white; color: black;")
        self.pushButton_watch.setStyleSheet("background-color: white; color: black;")

        # при нажатии на кнопки вызываем функции
        self.pushButton_delete.clicked.connect(self.d)
        self.pushButton_watch.clicked.connect(self.d_watch)

    def d(self):  # функция удаления
        conn = sqlite3.connect('project_2.db')
        cursor = conn.cursor()

        n_id = self.textEdit.toPlainText().strip()  # Убираем лишние пробелы
        try:
            a = int(n_id)  # Преобразуем в int
            # удаляем картинку из общей папки
            r = cursor.execute(f'SELECT image FROM gods WHERE ID = {a}').fetchone()[0]
            os.remove(r)
            # удаляем существо из таблицы gods
            cursor.execute(f'DELETE FROM gods WHERE ID = {a}')
            # удаляем ствроку из таблицы gods_books
            cursor.execute(f'DELETE FROM gods_books WHERE god_id = {a}')
            conn.commit()  # Подтверждение изменений

            conn.close()  # Закрываем соединение

            QMessageBox.information(self, "Обновление", f"Существо усешно удаленно.")

        except:
            QMessageBox.warning(self, "Ошибка", f"Введите число!")

    def d_watch(self):  # функция просмотра существ в таблице, чтобы проверить, что мы удалили
        conn = sqlite3.connect('project_2.db')
        cursor = conn.cursor()

        cursor.execute('''
                    SELECT ID, name, alias FROM gods
                ''')
        rows = cursor.fetchall()  # получаем данные для заполнения таблицы

        # Очистка таблицы перед загрузкой новых данных
        self.tableWidget_delete.clear()
        self.tableWidget_delete.setRowCount(len(rows))
        self.tableWidget_delete.setColumnCount(3)  # количество колонок
        self.tableWidget_delete.setHorizontalHeaderLabels(['ID', 'name', 'alias'])  # названия колонок
        # заполняем таблицу
        for i, row in enumerate(rows):
            for j, elem in enumerate(row):
                self.tableWidget_delete.setItem(i, j, QTableWidgetItem(str(elem)))

        # Настройка размеров колонок
        self.tableWidget_delete.resizeColumnsToContents()
        conn.close()  # Закрыть соединение


class ChangeForm(QMainWindow):
    def __init__(self, name):
        super().__init__()
        uic.loadUi('change.ui', self)

        self.btn_gods.setText('Внести изменения в записи о ' + name)
        self.btn_inf.clicked.connect(self.inf)

        self.btn_change.clicked.connect(self.choose)
        self.btn_watch.clicked.connect(self.d_watch)

    def inf(self):  # выбираем какой пункт хотим изменить
        inf, ok_pressed = QInputDialog.getItem(
            self, "Выбор существа", "Выберите суцество",
            ('name - имя', 'alias - название', 'description - описание', 'ability - способности',
             'time_of_stay - время первого упоминания', 'image - картинка',
             'book - источник первого упоминания'), 1, False)

        if ok_pressed and inf:
            self.btn_inf.setText('Вы выбрали: ' + inf)
            if inf.split(' - ')[0] == 'image':  # если мы выбрали изменить картинку, то выбираем новую
                fname = QFileDialog.getOpenFileName(self, 'Выбрать картинку', '',
                                                    'Картинка (*.jpg);;Картинка (*.png);;Все файлы (*)')[0]
                self.btn_inf.setText('Вы выбрали: ' + inf)
                self.textEdit.setPlainText(fname)
            elif inf.split(' - ')[0] == 'book':  # если мы выбрали книжку, то выбираем новую книгу
                conn = sqlite3.connect('project_2.db')
                cursor = conn.cursor()

                cursor.execute('''
                                SELECT name FROM books
                                ''')
                r = cursor.fetchall()  # получаем книжки
                book, ok_pressed = QInputDialog.getItem(
                    self, "Выбор книжки", "Выберите книжку",
                    ([i[0] for i in r]), 1, False)
                if ok_pressed and book:
                    self.textEdit.setPlainText(book)
            else:  # если выбрали что-то другое
                conn = sqlite3.connect('project_2.db')
                cursor = conn.cursor()
                god = self.btn_gods.text().split()[-1]  # получаем существо
                inf = inf.split(' - ')[0]  # подучаем критерий для изменения
                a = f'''
                    SELECT {inf} FROM gods
                    WHERE name == '{god}'
                    '''
                cursor.execute(a)
                inf_t = cursor.fetchall()[0]  # получаем значение
                for i in inf_t:
                    self.textEdit.setPlainText(i)

    def choose(self):
        conn = sqlite3.connect('project_2.db')
        cursor = conn.cursor()

        god = self.btn_gods.text().split()[-1]  # получаем существо
        inf = self.btn_inf.text().split()[2]  # получаем информацию
        text = self.textEdit.toPlainText()  # получаем на что изменить

        if inf == 'book':
            cursor.execute(f'''
                SELECT ID FROM gods
                WHERE name == '{god}'
                ''')
            id_god = cursor.fetchall()[0]  # получаем id существа
            cursor.execute(f'''
                        SELECT ID FROM books
                        WHERE name == '{text}'
                    ''')
            id_book = cursor.fetchall()[0]  # получаем id книги
            query = f'''
                UPDATE gods_books
                SET  book_id = ?
                WHERE god_id = ?
                '''
            cursor.execute(query, (id_book[0], id_god[0]))  # изменяем
        elif inf == 'name':
            new_name = self.textEdit.toPlainText()
            if new_name:  # если пользователь ввел имя
                cursor.execute(f'''
                UPDATE gods
                SET  name = '{new_name}'
                WHERE name = '{god}'
                ''')  # измением имя
                self.btn_gods.setText('Ввнести изменения в записи о ' + new_name)  # меняем имя на экране
            else:
                QMessageBox.information(self, "Ошибка", f"Введите имя. Нельзя оставлять имя пустым.")
                return
        elif inf == 'image':
            image = Image.open(text)  # открываем изображение
            directory = 'image'
            # Определяем новый путь для сохранения
            image_path = os.path.join(directory, f"{god}.jpg")
            # Сохраняем изображение
            image.save(image_path, "JPEG")
            cursor.execute(f'''
                    UPDATE gods
                    SET image = '{image_path}'
                    WHERE name = '{god}'
                ''')
        else:
            query = f'''
                    UPDATE gods
                    SET {inf} = ?
                    WHERE name = ?
                '''
            cursor.execute(query, (text, god))  # сохраняем изменения

        conn.commit()  # Подтверждение изменений

        QMessageBox.information(self, "Обновление", f"Изменения спешно внесены.")

        conn.close()  # Закрываем соединение

        self.textEdit.clear()  # очищаем текстовое поле

    def d_watch(self):
        god = self.btn_gods.text().split()[-1]  # получаем существо для просмотра
        self.close()
        self.gods_form = Gods_Form(god, 'изменение')
        self.gods_form.show()


class SaveForm(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('save.ui', self)
        self.setStyleSheet("background-color: #FFF8DC;")

        self.btn_chek.clicked.connect(self.save)

    def save(self):
        # получаем значения виджетов QCheckBox
        all_god = self.checkBox_all.isChecked()
        one_god = self.checkBox_one.isChecked()
        save_exe = self.checkBox_exe.isChecked()
        save_csv = self.checkBox_csv.isChecked()
        # проверяем на правильность отмечаных данных
        if all_god and one_god:
            QMessageBox.warning(self, 'Ошибка', 'Выберите только один из вариантов сохранение существ!')
        elif save_csv and save_exe:
            QMessageBox.warning(self, 'Ошибка', 'Выберите один формат файла!')
        elif all_god:
            # выбераем папку
            target_folder = QFileDialog.getExistingDirectory(self, "Выберите папку")

            conn = sqlite3.connect('project_2.db')
            cursor = conn.cursor()

            cursor.execute('''SELECT gods.name AS god_name, gods.alias, gods.description,
                             gods.ability, gods.time_of_stay, books.name AS book_name FROM gods_books
                            JOIN gods ON gods_books.god_id = gods.ID
                            JOIN books ON gods_books.book_id = books.ID''')
            rows = cursor.fetchall()
            conn.close()
            # получаем имя файла
            csv_file_name, ok_pressed = QInputDialog.getText(self, "Введите имя", "Введите название файла:")
            if ok_pressed and csv_file_name:
                # Получение названий столбцов
                column_names = [description[0] for description in cursor.description]
            if save_csv:
                # Запись в CSV файл
                csv_file_path = os.path.join(target_folder, f"{csv_file_name}.csv")  # формируем путь файла
                with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(column_names)  # Записываем заголовки
                    writer.writerows(rows)  # записываем значения
                QMessageBox.information(self, "Информация", f"Файл успешно сохранен.")
            elif save_exe:
                # Запись в exe файл
                csv_file_path = os.path.join(target_folder, f"{csv_file_name}.xlsx")  # формируем путь файла
                workbook = xlsxwriter.Workbook(csv_file_path)
                worksheet = workbook.add_worksheet()
                # Записываем заголовки
                worksheet.write(0, 0, 'Имя существа')  # Заголовок для Имени существа
                worksheet.write(0, 1, 'Название')  # Заголовок для Названия
                worksheet.write(0, 2, 'Описание')  # Заголовок для Описания
                worksheet.write(0, 3, 'Способность')  # Заголовок для Способности
                worksheet.write(0, 4, 'Время первого упоминания')  # Заголовок для Времени первого упоминания
                worksheet.write(0, 5, 'Источник первого упоминания')  # Заголовок для Источника первого упоминания
                #  записываем информацию в файл
                for row, (name, alias, description, ability,
                          time_of_stay, book) in enumerate(rows, start=1):  # Начинаем с 1, чтобы пропустить заголовки
                    worksheet.write(row, 0, name)
                    worksheet.write(row, 1, alias)
                    worksheet.write(row, 2, description)
                    worksheet.write(row, 3, ability)
                    worksheet.write(row, 4, time_of_stay)
                    worksheet.write(row, 5, book)
                workbook.close()
                QMessageBox.information(self, "Информация", f"Файл успешно сохранен.")
            else:
                QMessageBox.warning(self, 'Ошибка', 'Выберите формат файла!')
        elif one_god:
            conn = sqlite3.connect('project_2.db')
            cursor = conn.cursor()

            cursor.execute('''SELECT name FROM gods''')

            rows = cursor.fetchall()

            god, ok_pressed = QInputDialog.getItem(
                self, "Выбор существа", "Выберите существо", ([i[0] for i in rows]), 1, False)  # выбираем существо
            if ok_pressed:
                cursor.execute(f'''SELECT gods.name AS god_name, gods.alias, gods.description,
                                             gods.ability, gods.time_of_stay, books.name AS book_name FROM gods_books
                                            JOIN gods ON gods_books.god_id = gods.ID
                                            JOIN books ON gods_books.book_id = books.ID
                                            WHERE gods.name = "{god}"''')
                row = cursor.fetchall()
                # выбераем папку
                target_folder = QFileDialog.getExistingDirectory(self, "Выберите папку")
                # получаем имя файла
                csv_file_name, ok_pressed = QInputDialog.getText(self, "Введите имя", "Введите название файла:")
                if ok_pressed and csv_file_name:
                    # Получение названий столбцов
                    column_names = [description[0] for description in cursor.description]
                if save_csv:
                    # Запись в CSV файл
                    csv_file_path = os.path.join(target_folder, f"{csv_file_name}.csv")  # формируем путь файла
                    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
                        writer = csv.writer(file)
                        writer.writerow(column_names)  # Записываем заголовки
                        writer.writerows(row)  # записываем информацию
                    QMessageBox.information(self, "Информация", f"Файл успешно сохранен.")
                elif save_exe:
                    # Запись в exe файл
                    csv_file_path = os.path.join(target_folder, f"{csv_file_name}.xlsx")  # формируем путь файла
                    workbook = xlsxwriter.Workbook(csv_file_path)
                    worksheet = workbook.add_worksheet()
                    # Записываем заголовки
                    worksheet.write(0, 0, 'Имя существа')  # Заголовок для Имени существа
                    worksheet.write(0, 1, 'Название')  # Заголовок для Названия
                    worksheet.write(0, 2, 'Описание')  # Заголовок для Описания
                    worksheet.write(0, 3, 'Способность')  # Заголовок для Способности
                    worksheet.write(0, 4, 'Время первого упоминания')  # Заголовок для Времени первого упоминания
                    worksheet.write(0, 5, 'Источник первого упоминания')  # Заголовок для Источника первого упоминания

                    for row, (name, alias, description, ability,
                              time_of_stay, book) in enumerate(row, start=1):
                        # Начинаем с 1, чтобы пропустить заголовки
                        worksheet.write(row, 0, name)
                        worksheet.write(row, 1, alias)
                        worksheet.write(row, 2, description)
                        worksheet.write(row, 3, ability)
                        worksheet.write(row, 4, time_of_stay)
                        worksheet.write(row, 5, book)
                    workbook.close()
                    QMessageBox.information(self, "Информация", f"Файл успешно сохранен.")
                else:
                    QMessageBox.warning(self, 'Ошибка', 'Выберите формат файла!')
        else:
            QMessageBox.warning(self, 'Ошибка', 'Выберите что и в какой формате вы хотите сохранить!')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = FirstForm()
    ex.show()
    sys.exit(app.exec())

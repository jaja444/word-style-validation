import win32com.client as win32
from dataclasses import dataclass
import re

@dataclass
class HeadFormat:
    font_size: float
    bold: bool
    italics: bool
    left_indent: float  
    alignment: int
    space_before: float
    space_after: float

class Checker:
    def __init__(self, file_path): #выполняет запуск ms word для взаимодействия
        self.file_path = file_path
        self.word_app = win32.Dispatch("Word.Application")
        self.word_app.Visible = False
        self.doc = None
        
        self.format_standards = { #стандарты офоррмления
            1: HeadFormat(14, True, False, 1.5, 0, 6, 12),
            2: HeadFormat(12, True, False, 1.5, 0, 6, 6),
            3: HeadFormat(11, True, False, 1.5, 0, 6, 3),
            4: HeadFormat(11, True, True, 1.5, 0, 6, 3),
        }
        
        self.ignore_styles = [ #cтили, которые нужно игнорировать при проверке заголовков
            'таблица', 
            'table',
            'название таблицы',
            'table title',
            'таблица-название',
            'table-title'
        ]
        
        self.headings_results = None
        self.tables_results = None
    
    def detect_heading(self, paragraph): #определяет ур заголовка
        text = paragraph.Range.Text.strip()
        if not text:
            return None
        
        style_name = paragraph.Style.NameLocal.lower()
        
        if any(ignore_style in style_name for ignore_style in self.ignore_styles):
            return None
        
        if re.match(r'^Таблица\s+\d+\s*--\s*.+$', text, re.IGNORECASE):
            return None
        
        if any(ind in style_name for ind in ['heading', 'заголовок', 'h1', 'h2', 'h3', 'h4']):
            match = re.search(r'\d+', style_name)
            return min(int(match.group()), 4) if match else 1
        
        try:
            font_size = float(paragraph.Range.Font.Size)
            is_bold = bool(paragraph.Range.Font.Bold)
            is_italic = bool(paragraph.Range.Font.Italic)
            
            if font_size >= 13.5 and is_bold:
                return 1
            elif font_size >= 12 and is_bold:
                return 2
            elif font_size >= 11 and is_bold and is_italic:
                return 4
            elif font_size >= 10.5 and is_bold:
                return 3
        except:
            pass
                
        return None

    def check_format(self, paragraph, level): #проверяет форматирование
        errors = []
        std = self.format_standards[level]
        
        try:
            font_size = float(paragraph.Range.Font.Size)
            if abs(font_size - std.font_size) > 0.5:  
                errors.append(f"Размер шрифта: {font_size} вместо {std.font_size}")
            
            if bool(paragraph.Range.Font.Bold) != std.bold:
                errors.append(f"Жирность: {'ДА' if std.bold else 'НЕТ'}")
            
            if bool(paragraph.Range.Font.Italic) != std.italics:
                errors.append(f"Курсив: {'ДА' if std.italics else 'НЕТ'}")
            
            if hasattr(paragraph.Format, 'SpaceBefore'):
                space = float(paragraph.Format.SpaceBefore)
                if abs(space - std.space_before) > 1:  
                    errors.append(f"Отступ перед: {space} вместо {std.space_before}")
            
            if hasattr(paragraph.Format, 'SpaceAfter'):
                space = float(paragraph.Format.SpaceAfter)
                if abs(space - std.space_after) > 1:
                    errors.append(f"Отступ после: {space} вместо {std.space_after}")
                    
        except Exception as e: #если происходит ошибка создается переменная e
            errors.append(f"Ошибка: {str(e)}")
        
        return errors
    
    def check_headings(self): #основная проверка заголовков
        if not self.doc:
            self.doc = self.word_app.Documents.Open(self.file_path)
            
        results = {'headings': [], 'errors': [], 'correct': 0, 'ignored': []}
        
        for i, paragraph in enumerate(self.doc.Paragraphs, 1):
            try:
                if not paragraph.Range.Text.strip():
                    continue
                
                text = paragraph.Range.Text.strip()
                style_name = paragraph.Style.NameLocal.lower()
                
                if any(ignore_style in style_name for ignore_style in self.ignore_styles):
                    results['ignored'].append({
                        'num': i,
                        'text': text[:50],
                        'style': paragraph.Style.NameLocal,
                        'reason': 'стиль названия таблицы'
                    })
                    continue
                
                if re.match(r'^Таблица\s+\d+\s*--\s*.+$', text, re.IGNORECASE):
                    results['ignored'].append({
                        'num': i,
                        'text': text[:50],
                        'style': paragraph.Style.NameLocal,
                        'reason': 'название таблицы по тексту'
                    })
                    continue
                
                level = self.detect_heading(paragraph)
                if level is None:
                    continue
                
                font_size = float(paragraph.Range.Font.Size)
                is_bold = bool(paragraph.Range.Font.Bold)
                
                info = {
                    'num': i,
                    'level': level,
                    'text': text[:100],
                    'font_size': font_size,
                    'bold': is_bold,
                    'style': paragraph.Style.NameLocal,
                    'paragraph_index': i  
                }
                
                format_errors = self.check_format(paragraph, level)
                
                if format_errors:
                    info['errors'] = format_errors
                    results['errors'].append(info)
                else:
                    results['correct'] += 1
                
                results['headings'].append(info)
                    
            except Exception as e:
                print(f"Ошибка в параграфе {i}: {e}")
        
        self.headings_results = results
        return results
    
    def get_cell_text(self, cell): #получает текст из ячейки и передает ее другим методам
        try:
            text = cell.Range.Text
            text = text.replace('\x07', '').replace('\r', '').strip()
            return text
        except:
            return ""

    def check_empty_cells(self, table): #проверка на наличие пустые ячейки таблицы
        empty_cells = []
        for row_idx in range(1, table.Rows.Count + 1):
            for col_idx in range(1, table.Columns.Count + 1):
                try:
                    cell = table.Cell(row_idx, col_idx)
                    text = self.get_cell_text(cell)

                    if not text or text.lower() in ["", " ", "\t"]:
                        empty_cells.append((row_idx, col_idx))
                except:
                    continue
        return empty_cells

    def check_punctuation(self, table): #проверка на наличие точек в ячейках таблицы
        cells_with_dot = []
        for row_idx in range(1, table.Rows.Count + 1):
            for col_idx in range(1, table.Columns.Count + 1):
                try:
                    cell = table.Cell(row_idx, col_idx)
                    text = self.get_cell_text(cell)

                    if text and text.endswith('.'):
                        cells_with_dot.append((row_idx, col_idx))
                except:
                    continue
        return cells_with_dot

    def check_table_title(self, table_index, table): #проверка названия таблицы на корректность
        try:
            table_range = table.Range
            start_pos = table_range.Start - 1
            
            if start_pos < 0:
                return False, "Название таблицы отсутствует"

            title_paragraph = self.doc.Range(start_pos, start_pos + 1).Paragraphs(1)
            title_text = title_paragraph.Range.Text.strip()

            title_text = title_text.replace('\r', '').strip()

            pattern = r'^Таблица\s+\d+\s*--\s*.+$'
            if not re.match(pattern, title_text, re.IGNORECASE):
                return False, f"Неправильный формат названия: '{title_text}'"

            if not title_text[0].isupper():
                return False, "Название должно начинаться с заглавной буквы"

            style_name = title_paragraph.Style.NameLocal
            if not any(table_style in style_name.lower() for table_style in ['таблица', 'table']):
                return False, f"Неверный стиль названия: '{style_name}' (должен содержать 'таблица' или 'table')"
            
            return True, title_text
        except Exception as e:
            return False, f"Ошибка при проверке названия: {str(e)}"

    def check_header_row(self, table): #проверка заголовка таблицы на соотв. стиль и двойную линию
        try:
            if table.Rows.Count < 1:
                return False, "Таблица не содержит строк"
            
            errors = []
            
            first_row = table.Rows(1)
            first_row_style = first_row.Range.ParagraphFormat.Style.NameLocal
            
            if "таблица" not in first_row_style.lower() or "заголовок" not in first_row_style.lower():
                errors.append(f"Неверный стиль заголовка: '{first_row_style}'")
            
            try:
                border = first_row.Borders(3)  #3 = нижняя граница
                
                if border.LineStyle != 7:  #7 = двойная линия
                    errors.append(f"Отсутствует двойная линия под заголовком (стиль линии: {border.LineStyle})")
            except Exception as border_error:
                errors.append(f"Не удалось проверить границу заголовка: {str(border_error)}")
            
            if errors:
                return False, errors
            else:
                return True, []
                
        except Exception as e:
            return False, [f"Ошибка при проверке заголовка: {str(e)}"]

    def check_caps(self, table): #проверка заглавных букв в ячейках
        cells_without_caps = []
        
        for row_idx in range(2, table.Rows.Count + 1):  
            for col_idx in range(1, table.Columns.Count + 1):
                try:
                    cell = table.Cell(row_idx, col_idx)
                    text = self.get_cell_text(cell)
                    
                    if not text:
                        continue
                    
                    first_char = text[0] if text else ""
                    if first_char and first_char.isalpha() and not first_char.isupper():
                        cells_without_caps.append((row_idx, col_idx))
                except:
                    continue
        return cells_without_caps

    def check_tables(self, table_index, table): #главный метод, который собирает результаты всех отдельных проверок
        table_info = {
            'index': table_index,
            'title': "",
            'has_title': False,
            'title_correct': False,
            'header_correct': False,
            'errors': [],
            'empty_cells': [],
            'cells_with_dot': [],
            'cells_without_caps': []
        }
        
        has_title, title_result = self.check_table_title(table_index, table)
        table_info['has_title'] = has_title
        
        if has_title:
            table_info['title_correct'], table_info['title'] = True, title_result
        else:
            table_info['errors'].append(title_result)
        
        header_correct, header_error = self.check_header_row(table)
        table_info['header_correct'] = header_correct
        if not header_correct:
            table_info['errors'].append(header_error)

        empty_cells = self.check_empty_cells(table)
        table_info['empty_cells'] = empty_cells
        if empty_cells:
            for row, col in empty_cells:
                table_info['errors'].append(f"Пустая ячейка в строке {row}, столбце {col}")

        cells_with_dot = self.check_punctuation(table)
        table_info['cells_with_dot'] = cells_with_dot
        if cells_with_dot:
            for row, col in cells_with_dot:
                table_info['errors'].append(f"Точка в конце текста в строке {row}, столбце {col}")

        cells_without_caps = self.check_caps(table)
        table_info['cells_without_caps'] = cells_without_caps
        if cells_without_caps:
            for row, col in cells_without_caps:
                table_info['errors'].append(f"Отсутствует заглавная буква в строке {row}, столбце {col}")
        
        return table_info

    def check_all_tables(self): #проверка всех таблиц
        if not self.doc:
            self.doc = self.word_app.Documents.Open(self.file_path)
            
        tables_info = []
        
        try:
            tables_count = self.doc.Tables.Count
            
            for table_idx in range(1, tables_count + 1):
                table = self.doc.Tables(table_idx)
                table_info = self.check_tables(table_idx, table)
                tables_info.append(table_info)
        
        except Exception as e:
            print(f"Ошибка при проверке таблиц: {e}")
        
        self.tables_results = tables_info
        return tables_info
    
    def print_results(self):
        try:
            self.word_app = win32.Dispatch("Word.Application")
            self.word_app.Visible = False
            
            self.doc = self.word_app.Documents.Open(self.file_path)
            headings_data = self.check_headings()
            tables_data = self.check_all_tables()
            # добавили валидацию рисунков
            
            if 'error' in headings_data:
                print(f"\n ОШИБКА ПРОВЕРКИ ЗАГОЛОВКОВ: {headings_data['error']}")
            else:
                total_headings = len(headings_data['headings'])
                correct_headings = headings_data['correct']
                
                print("Результаты проверки заголовков")
                print(f"Всего заголовков: {total_headings}")
                print(f"Правильно оформленных: {correct_headings}")
                print(f"С ошибками: {len(headings_data['errors'])}")
                
                if total_headings > 0:
                    for heading in headings_data['headings'][:10]:  
                        print(f"Ур.{heading['level']} | "
                              f"Шрифт: {heading['font_size']:4.1f} | "
                              f"Стиль: {heading['style'][:20]:20} | "
                              f"Текст: {heading['text']}")
                
                if headings_data['errors']:
                    print(f"\nОШИБКИ В ЗАГОЛОВКАХ:")
                    for error in headings_data['errors'][:5]:  #показываем первые 5 ошибок
                        print(f"\nСтрока {error['paragraph_index']} (Уровень {error['level']}):")
                        print(f"Текст: '{error['text'][:80]}...'" if len(error['text']) > 80 else f"Текст: '{error['text']}'")
                        for err in error.get('errors', []):
                            print(f"  - {err}")
                    
                elif total_headings > 0:
                    print(f"\n Все заголовки оформлены правильно!")
                else:
                    print(f"\n Заголовки не найдены!")
            
            tables_count = len(tables_data)
            print("Результаты проверки таблиц")
            print(f"Обнаружено таблиц: {tables_count}\n")
            
            correct_tables = 0
            incorrect_tables = 0
            
            for table_info in tables_data:
                table_num = table_info['index']
                
                if table_info['errors']:
                    incorrect_tables += 1
                    print(f"Таблица {table_num}: Найдены ошибки")
                    for error in table_info['errors'][:3]:  #показываем первые 3 ошибки
                        print(f"  - {error}")
                    
                    if len(table_info['errors']) > 3:
                        print(f"  ... и еще {len(table_info['errors']) - 3} ошибок")
                    print()
                else:
                    correct_tables += 1
                    print(f"Таблица {table_num}: Правильно оформлена")
                    if table_info['title']:
                        print(f"  Название: {table_info['title']}")
                    print()
            
            print(f"Правильно оформленных таблиц: {correct_tables}")
            print(f"Таблиц с ошибками: {incorrect_tables}")
            
            if incorrect_tables == 0 and tables_count > 0:
                print(f"\n Все таблицы оформлены правильно!")
            
        except Exception as e:
            print(f"\n ОШИБКА ПРИ ПРОВЕРКЕ ДОКУМЕНТА: {e}")
            return None
            
        finally:
            self._close()

    def _close(self):
        try:
            if self.doc:
                self.doc.Close(False)
            if self.word_app:
                self.word_app.Quit()
        except:
            pass
if __name__ == "__main__":
    file_path = r"C:\Users\glebk\Desktop\тест.docx"
    checker = Checker(file_path)
    results = checker.print_results()

    
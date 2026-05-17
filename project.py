# -*- coding: utf-8 -*-
import calendar
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import DataBarRule

def create_habit_tracker():
    wb = Workbook()
    
    # Настройка листа "Данные"
    ws_data = wb.active
    ws_data.title = "Данные"
    
    habits_data = [
        ("🧘 Медитация", 10),
        ("💧 Пить воду", 31),
        ("🚶 10 000 шагов", 31),
        ("📖 Читать книгу", 10),
        ("🥗 Полезное питание", 31),
        ("🏋️ Тренировка", 10),
        ("😴 Сон до 23:00", 31),
        ("📓 Ведение дневника", 4),
        ("🧴 Уход за собой", 15),
        ("📵 Час без телефона", 31),
        ("📥 Проверить входящие заявки", 31),
        ("📱 Проверить соцсети", 31),
        ("📸 Выложить пост", 31),
        ("🎬 Снять/смонтировать сторис", 15),
        ("💬 Ответить клиентам", 31),
        ("📂 Разобрать рабочие чаты", 31),
        ("✅ Закрыть задачи в CRM", 31),
        ("📊 Проверить рекламу / статистику", 2),
        ("📄 Выставить счета / КП", 31),
        ("🗂 Навести порядок в проектах", 1),
    ]
    
    ws_data['A1'] = "Год"
    ws_data['B1'] = "Ваши привычки"
    ws_data['C1'] = "Сочетания клавиш для эмодзи:"
    ws_data['A2'] = 2026
    ws_data['C2'] = "Windows: Win + ; или Win + ."
    ws_data['C3'] = "Mac: Ctrl + Cmd + Пробел"
    
    for idx, (habit, goal) in enumerate(habits_data, start=2):
        ws_data[f'B{idx}'] = habit
    
    months_ru = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    ]
    weekdays_short = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    
    for month_idx in range(12):
        month_name = months_ru[month_idx]
        ws = wb.create_sheet(month_name)
        
        ws['A2'] = "ТРЕКЕР ПРИВЫЧЕК"
        ws['A2'].font = Font(bold=True, size=14)
        
        month_year_text = f"{month_name.upper()} 2026"
        ws['A3'] = month_year_text
        ws['A3'].font = Font(italic=True, size=12)
        
        # Календарь справа
        cal_start_col = 61 
        cal = calendar.monthcalendar(2026, month_idx + 1)
        month_days = calendar.monthrange(2026, month_idx + 1)[1]
        
        ws.merge_cells(start_row=2, start_column=cal_start_col, end_row=2, end_column=cal_start_col+6)
        cal_header_cell = ws.cell(row=2, column=cal_start_col)
        cal_header_cell.value = month_year_text
        cal_header_cell.alignment = Alignment(horizontal='center')
        cal_header_cell.font = Font(bold=True)
        
        for col_offset, day_name in enumerate(weekdays_short):
            cell = ws.cell(row=3, column=cal_start_col + col_offset)
            cell.value = day_name
            cell.font = Font(bold=True, size=8)
            cell.alignment = Alignment(horizontal='center')
        
        row_offset = 4
        for week in cal:
            for col_offset, day in enumerate(week):
                if day != 0:
                    cell = ws.cell(row=row_offset, column=cal_start_col + col_offset)
                    cell.value = day
                    cell.alignment = Alignment(horizontal='center')
            row_offset += 1
        
        # Сетка трекера
        week_headers = ["1 неделя", "2 неделя", "3 неделя", "4 неделя", "5 неделя"]
        # ИСПРАВЛЕНО: убран символ #, оставлены чистые HEX-коды
        week_colors = ["FFE4D6", "D6E4FF", "E0F0E0", "E8E0F0", "FFE0E8"]
        
        for i, (header, color) in enumerate(zip(week_headers, week_colors)):
            start_col = 3 + i * 7
            end_col = start_col + 6
            ws.merge_cells(start_row=11, start_column=start_col, end_row=11, end_column=end_col)            cell = ws.cell(row=11, column=start_col)
            cell.value = header
            cell.alignment = Alignment(horizontal='center')
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
        
        for day in range(1, 32):
            col = 2 + day
            cell_num = ws.cell(row=12, column=col)
            cell_num.value = day
            cell_num.alignment = Alignment(horizontal='center')
            cell_num.font = Font(size=8)
            
            if day <= month_days:
                date_obj = datetime(2026, month_idx + 1, day)
                weekday_name = weekdays_short[date_obj.weekday()]
                cell_wd = ws.cell(row=13, column=col)
                cell_wd.value = weekday_name
                cell_wd.alignment = Alignment(horizontal='center')
                cell_wd.font = Font(size=7)
        
        ws['B12'] = "ЦЕЛЬ"
        ws['B12'].font = Font(bold=True)
        
        headers_right = ["ИТОГО", "СЕРИЯ", "ВЫПОЛНЕНО", "ПРОГРЕСС"]
        right_cols = [42, 43, 44, 45]
        for col, header in zip(right_cols, headers_right):
            cell = ws.cell(row=12, column=col)
            cell.value = header
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        
        visible_habits = habits_data[:10]
        
        for row_idx, (habit, goal) in enumerate(visible_habits, start=14):
            ws.cell(row=row_idx, column=1, value=habit)
            ws.cell(row=row_idx, column=2, value=goal)
            
            for day in range(1, 32):
                if day <= month_days:
                    ws.cell(row=row_idx, column=2 + day, value="")
            
            total_col = 42
            ws.cell(row=row_idx, column=total_col, value=f'=COUNTIF(C{row_idx}:AG{row_idx},"x")+COUNTIF(C{row_idx}:AG{row_idx},"х")')
            ws.cell(row=row_idx, column=43, value=0)
            ws.cell(row=row_idx, column=44, value=f'=AJ{row_idx}&"/"&{goal}')
            ws.cell(row=row_idx, column=45, value=f'=AJ{row_idx}/{goal}')
            ws.cell(row=row_idx, column=45).number_format = '0%'
        
        for row in range(11, 25):
            for col in range(1, 46):                cell = ws.cell(row=row, column=col)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if row in (11, 12, 13) or col in (1, 2) or col >= 42:
                    cell.border = Border(bottom=Side(style='thin'), top=Side(style='thin'))
        
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 8
        for day in range(1, 32):
            ws.column_dimensions[get_column_letter(day + 2)].width = 3.5
        
        ws.column_dimensions['AJ'].width = 8
        ws.column_dimensions['AK'].width = 8
        ws.column_dimensions['AL'].width = 12
        ws.column_dimensions['AM'].width = 12
        
        for row_idx in range(14, 24):
            ws.conditional_formatting.add(
                f'AM{row_idx}',
                DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color="00B050", showValue=True)
            )
        
        ws['A4'] = '=TEXT(AVERAGE(AM14:AM23),"0%")'
        ws['A4'].font = Font(bold=True, size=12)
        ws['A4'].number_format = '0%'

    filename = 'Трекер_привычек_2026.xlsx'
    wb.save(filename)
    print(f"✅ Трекер успешно создан: {filename}")

if __name__ == "__main__":
    create_habit_tracker()
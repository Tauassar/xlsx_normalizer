import logging

from settings import max_search_index, fact_sheet_name, search_strings

logger = logging.getLogger(__name__)


def get_export_filename(file_name):
    export_filename = file_name.replace('.xlsx', '')
    permitted_symbols = '()<>'
    for symbol in permitted_symbols:
        export_filename = export_filename.replace(symbol, '')
    return f'out/{export_filename}_для_загрузки.xlsx'


def search_value_in_col_idx(ws, search_string):
    try:
        global fact_search_row
        for col_idx in range(0, max_search_index):
            for row in range(1, max_search_index):
                if ws[row][col_idx].value:
                    if search_string in str(ws[row][col_idx].value):
                        fact_search_row = row
                        return col_idx, row - 1
        return col_idx, None
    except IndexError:
        logger.debug(f'Столбец {search_string} не найден')
        return None


def search_fact_values(ws):
    try:
        search_string = 'Факт'
        fact_data = {}
        for col_idx in range(0, max_search_index):
            if ws[fact_search_row][col_idx].value:
                if search_string in str(ws[fact_search_row][col_idx].value):
                    fact_data[ws[fact_search_row - 1][col_idx].value] = [col_idx, fact_search_row - 1]
                    # continue
        return fact_data

    except IndexError:
        logger.debug(f'Парсинг данных завершен {fact_data}')
        return fact_data


def get_indexes_for_search_values(ws):
    ws_indices = {}
    for value in search_strings:
        ws_indices[value] = search_value_in_col_idx(ws, search_string=value)
        logger.debug(f'{value}: {ws_indices[value]}')
        if ws_indices[value] is None:
            raise ValueError(f'В исходном файле отсутствует данные для столбца {value}. Пожалуйста дополните данные в исходном документе')

    return ws_indices | search_fact_values(ws)


def get_fact_sheet(wb):
    """
    Searching for sheet containing Факт word in input file
    :param wb:
    :return: wb.sheet
    """
    fact_found = False
    logger.debug("Поиск листа с фактическими данными инициирован")
    if len(wb.sheetnames) > 1:
        for i, s in enumerate(wb.sheetnames):
            if fact_sheet_name.lower() in s.lower():
                fact_found = True
                logger.debug(f'Лист с фактическими данными найден, наименование листа: {s}')
                break
        if fact_found:
            wb.active = i
            logger.info(f'Лист {wb.active} установлен активным листом, из-за наличия слова "{fact_sheet_name}" в названии')
        else:
            logger.info(f'Лист с наименованием {fact_sheet_name} не найден, оставлен лист по умолчанию')
    return wb.active


def check_if_sheet_content_positions_satisfies_requirements(ws):
    values_to_check = {
        'Инвестиционный проект/': 'A7',
        'Метод консолидации': 'C7',
        'Статус реализации': 'E7',
        'В рамках Бизнес-плана': 'G7',
        'Жизненный этап проекта': 'I7',
        'Статья инвестиционного проекта': 'K7',
        'КалендГод/Квартал': 'L6',
    }
    for value in values_to_check:
        if ws[values_to_check[value]].value != value:
            logger.error('Документ не соответствует требованиям')
            return False
    logger.info('Контент соответствует требованиям')
    return True


def check_if_sheet_name_satisfies_requirements(wb):
    logger.debug("Проверка названий листов")
    if len(wb.sheetnames) > 1 or wb.active.title != fact_sheet_name:
        logger.warning('Наименование листа не соответствует требованию')
        logger.info(wb.active)
        return False
    logger.info('Наименование листа соответствует требованию')
    return True

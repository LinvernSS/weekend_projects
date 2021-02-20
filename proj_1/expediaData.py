import openpyxl
import logging
import datetime


def load_wb(fname):
    filepath = r'C:/Users/lucas/OneDrive/Desktop/Stuff/Job Stuff/Smoothstack git/weekend_proj/proj_1/Weekend_work/'
    try:
        workbook = openpyxl.load_workbook(filepath + fname)
    except IOError:
        logging.error('file not found.')
        exit(-1)
    else:
        logging.info('file loaded.')
        return workbook


def get_date(fname):
    fname = fname.split('_')
    month = fname[-2]
    month = month[0].upper() + month[1:]
    month = datetime.datetime.strptime(month, "%B").month
    year = int(fname[-1][:-5])
    return datetime.datetime(year, month, 1)


def trunc_date(date):
    return date.replace(day=1, hour=0, minute=0, second=0, microsecond=0)


def get_data(date, book):
    worksheets = []
    for sheet in book:
        worksheets.append(sheet)

    row_num = 2
    for cell in worksheets[0]['A2':'A13']:
        val = trunc_date(cell[0].value)
        if date == val:
            break
        row_num += 1
    row_num = str(row_num)

    for cell in worksheets[0]['B'+row_num:'F'+row_num][0]:
        val = cell.value
        col = str(cell.coordinate[0])
        col_name = worksheets[0][col + '1']
        if type(val) != int:
            val *= 100
            val = str(val) + '%'
        else:
            val = f'{val:,}'
        logging.info('{} : {}'.format(col_name.value.strip(), val))


if __name__ == "__main__":
    logging.basicConfig(filename='expediaData.log', level=logging.DEBUG)
    filename = 'expedia_report_monthly_january_2018.xlsx'
    wb = load_wb(filename)
    mon_yr = get_date(filename)
    get_data(mon_yr, wb)
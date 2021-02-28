import os
import re
import imghdr
import smtplib
import logging
import numpy as np
import pandas as pd
import phonenumbers as pn
import matplotlib.pyplot as plt
from email.message import EmailMessage


# Given a list of filenames, returns file with most recent date
# Assumed all files in list are same format
def find_recent_file(fns):
    if type(fns) == np.ndarray:
        if len(fns[0]) > 4:
            remove_ext = lambda f: f.split('_')[-1][:-4]
            ind = np.array(list(map(remove_ext, fns)))
            rf = fns[ind.argmax()]
            return rf
    logging.error('One or more files is in the wrong format.')
    return False


# Creates/reads from NYL.lst, checks if given filename is logged there.
# Stops process if logged, logs and continues otherwise
def log_process(f):
    try:
        file_list = open('output/NYL.lst', 'x')
        file_list.write(f + '\n')
        file_list.close()
        logging.info('Created NYL.lst and logged {} as processed.'.format(f))
        return True
    except FileExistsError:
        file_list = open('output/NYL.lst', 'r')
        if f in file_list.read().split('\n'):
            file_list.close()
            logging.warning('File already processed.')
            return False
        else:
            file_list.close()
            file_list = open('output/NYL.lst', 'a')
            file_list.write(f + '\n')
            file_list.close()
            logging.info('Logged {} as processed.'.format(f))
            return True


# Given a csv filename, tries to return file contents
def load_data(f):
    try:
        d = pd.read_csv('data/' + f)
    except FileNotFoundError:
        logging.error('File {} Not Found.'.format(f))
        return False
    else:
        logging.info('Successfully loaded {}.'.format(f))
        return d


# Given two dataframes, halt process if length difference > 500, else continue
def validate_file_len(d, prevd):
    if abs(len(d.index) - len(prevd.index)) > 500:
        logging.error('{} File variance too great.')
        return False
    else:
        logging.info('File lengths validated.')
        return True


# Replaces secondary headers with standardized headers if present
def replace_headers(d):
    c1 = 'Agent Writing Contract Start Date (Carrier appointment start date)'
    c3 = "Agent Writing Contract Status (actually active and cancelled's should come in two different files)"
    if c1 in d.columns or c3 in d.columns:
        d = d.rename(columns={c1: 'Agent Writing Contract Start Date'})
        d = d.rename(columns={c3: 'Agent Writing Contract Status'})
    return d


# Removes unnecessary whitespace from given dataframe
def format_data(d):
    try:
        for i, col in d.items():
            if col.dtype == object:
                d[i] = col.str.strip().replace({' +': ' '}, regex=True)
    except Exception as ex:
        logging.error('Failed formatting data. {}'.format(ex))
    finally:
        return d


# Returns true if given US phone number is valid, else false
def is_invalid_pn(num):
    num = '+1' + num
    try:
        num = pn.parse(num)
    except pn.NumberParseException:
        return True
    else:
        if pn.is_valid_number(num) and pn.is_possible_number(num):
            return False
        else:
            return True


# Validates all phone numbers in given dataframe, logs invalid phone numbers
def find_valid_pn(d):
    invalid_nums = pd.Series(map(is_invalid_pn, d['Agency Phone Number']))
    if invalid_nums.sum() > 0:
        for agent_id in d[invalid_nums]['Agent Id']:
            logging.warning('Agent {} has an invalid phone number.'.format(agent_id))
        return False
    else:
        logging.info('All phone numbers validated.')
        return True


# Validates all agency states in given dataframe, logs invalid states
def find_valid_state(d):
    states_abbr = ['AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA',
                   'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD',
                   'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
                   'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC',
                   'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY']
    is_invalid_st = lambda s: False if s in states_abbr else True
    invalid_agency_states = pd.Series(map(is_invalid_st, d['Agency State']))
    invalid_agent_states = pd.Series(map(is_invalid_st, d['Agent State']))
    invalid_states = invalid_agent_states | invalid_agency_states
    if invalid_states.sum() > 0:
        for agent_id in d[invalid_states]['Agent Id']:
            logging.warning('Agent {} has an invalid state.'.format(agent_id))
        return False
    else:
        logging.info('All states validated.')
        return True


# Validates all agent's email addresses in given dataframe, logs invalid emails
def find_valid_email(d):
    is_invalid_em = lambda em: True if not re.match(r'[\w\.-]+@[\w\.-]+(\.[\w]+)+', em) else False
    invalid_emails = pd.Series(map(is_invalid_em, d['Agent Email Address']))
    if invalid_emails.sum() > 0:
        for agent_id in d[invalid_emails]['Agent Id']:
            logging.warning('Agent {} has an invalid email.'.format(agent_id))
        return False
    else:
        logging.info('All emails validated.')
        return True


# Combines and formats first, middle and last names from given dataframe
def format_agent_names(d):
    format_middle_names = lambda n: n if (n == '') else (n + ' ')
    first_names = d['Agent First Name'].str.strip().str.title() + ' '
    middle_names = pd.Series(map(format_middle_names, d['Agent Middle Name'].str.strip().str.title()))
    last_names = d['Agent Last Name'].str.strip().str.title()
    return first_names + middle_names + last_names


# Given dataframe, plots first column as x axis against two given shared y columns. saves in n.png
def plot_agent_info(ai, c1, c2, n):
    fig, ax1 = plt.subplots()
    color = 'tab:red'
    ax1.set_ylabel('A2O Date', color=color)
    ax1.plot(ai[0], ai[c2], '.r')
    ax1.tick_params(axis='y', labelcolor=color)

    ax2 = ax1.twinx()
    color = 'tab:blue'
    ax2.set_ylabel('Start Date', color=color)
    ax2.plot(ai[0], ai[c1], '.b')
    ax2.tick_params(axis='y', labelcolor=color)

    plt.xticks([])
    fig.tight_layout()
    plt.savefig('output/graphs/' + n + '.png')


# Summarizes, logs, and plots agent info data from given dataframe
def agent_info_data(d):
    names = format_agent_names(d)
    col1 = 'Agent Writing Contract Start Date'
    col2 = 'Date when an agent became A2O'
    agent_info = pd.concat([names, d[[col1, col2]]], axis=1)
    logging.info('Agent Start and A2O Date Info:\n\n{}\n'.format(agent_info))

    agent_info[col1] = pd.to_datetime(agent_info[col1], 'coerce')
    agent_info[col2] = pd.to_datetime(agent_info[col2], 'coerce')

    agent_info = agent_info.sort_values(col2)
    img_name = 'Agent_info_by_A2O'
    plot_agent_info(agent_info, col1, col2, img_name)

    agent_info = agent_info.sort_values(col1)
    img_name = 'Agent_info_by_start'
    plot_agent_info(agent_info, col1, col2, img_name)


def data_summary(d):
    logging.info('Starting data summary.')

    trans_d = d.T
    logging.info('Transposed Data:\n\n{}\n'.format(trans_d))

    state_group = d.sort_values('Agency State')
    logging.info('Data grouped by state:\n\n{}\n'.format(state_group))

    state_counts = state_group.groupby('Agency State')['Agent Id'].count().sort_values()
    plt.plot(state_counts.index, state_counts, '.r')
    plt.xticks(np.arange(0, 50, 3))
    plt.xlabel('State')
    plt.ylabel('Count')
    plt.title('Occurence of Agency States')
    plt.savefig('output/graphs/State_counts.png')

    agent_info_data(d)

    logging.info('Data summary complete.')


# sends email of log file, attaches images if cond is true, exits program when complete
def send_email(cond):
    if cond:
        logging.info('Program executed successfully.')
    else:
        logging.info('Program failed execution.')

    username = 'lucas.invernizzi@gmail.com'
    with open('output/conf.txt') as fp:
        password = fp.read()
    smtp_server = "smtp.gmail.com:587"
    msg = EmailMessage()
    msg['Subject'] = 'Results of NYL data analysis'
    msg['From'] = username
    msg['To'] = username

    with open('output/NYLData.log') as fp:
        msg.set_content(fp.read())

    if cond:
        imgs = ['Agent_info_by_A2O.png', 'Agent_info_by_start.png', 'State_counts.png']
        for img_name in imgs:
            with open('output/graphs/' + img_name, 'rb') as fp:
                img = fp.read()
            msg.add_attachment(img, maintype='image', subtype=imghdr.what(None, img), filename=img_name)

    s = smtplib.SMTP(smtp_server)
    s.starttls()
    s.login(username, password)
    # s.send_message(msg)
    s.quit()

    if cond:
        exit(0)
    else:
        exit(-1)


if __name__ == "__main__":
    escape = lambda b: (send_email(False) if not b else b) if type(b) == bool else b
    logging.basicConfig(format='%(asctime)s %(levelname)-4s %(message)s',
                        filename='output/NYLData.log',
                        filemode='w',
                        level=logging.DEBUG,
                        datefmt='%Y-%m-%d %H:%M:%S')
    mpl_logger = logging.getLogger('matplotlib')
    mpl_logger.setLevel(logging.WARNING)
    file_names = np.array(os.listdir('data/'))

    recent_file = escape(find_recent_file(file_names))
    # escape(log_process(recent_file))

    data = escape(load_data(recent_file))

    prev_file_names = np.delete(file_names, np.where(file_names == recent_file))
    prev_data = escape(load_data(find_recent_file(prev_file_names)))

    escape(validate_file_len(data, prev_data))
    data = replace_headers(data)
    data = format_data(data)

    find_valid_pn(data)
    find_valid_state(data)
    find_valid_email(data)

    data_summary(data)
    send_email(True)


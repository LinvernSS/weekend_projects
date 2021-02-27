import numpy as np
import pandas as pd
import os
import phonenumbers as pn
import re
import smtplib
import matplotlib.pyplot as plt
import imghdr
from email.message import EmailMessage
import logging


def find_recent_file(fns):
    try:
        remove_ext = lambda f: f.split('_')[-1][:-4]
        ind = np.array(list(map(remove_ext, fns)))
        rf = fns[ind.argmax()]
    except:
        logging.error('One or more files in wrong format')
        send_email(False)
    else:
        return rf


def log_process(f):
    try:
        file_list = open('NYL.lst', 'x')
        file_list.write(f + '\n')
        file_list.close()
        logging.info('Created NYL.lst and logged {} as processed.'.format(f))
    except FileExistsError:
        file_list = open('NYL.lst', 'r')
        if f in file_list.read().split('\n'):
            file_list.close()
            logging.warning('File already processed.')
            send_email(False)
        else:
            file_list.close()
            file_list = open('NYL.lst', 'a')
            file_list.write(f + '\n')
            file_list.close()
            logging.info('Logged {} as processed.'.format(f))


def load_data(f, p):
    try:
        d = pd.read_csv(p + f)
    except FileNotFoundError:
        logging.error('File {} Not Found.'.format(f))
        return None
    else:
        return d


def validate_file_len(rf, pf, fp):
    d = load_data(rf, fp)
    prev_d = load_data(pf, fp)
    if d is None or prev_d is None:
        logging.error('Unable to validate file lengths.')
        send_email(False)
    elif len(d.index) - len(prev_d.index) > 500:
        logging.error('File variance too great.')
        send_email(False)
    else:
        logging.info('Successfully loaded {}.'.format(rf))
        return d


def replace_headers(d):
    c1 = 'Agent Writing Contract Start Date (Carrier appointment start date)'
    c3 = 'Agent Writing Contract Status (actually active and cancelled\'s should come in two different files)'
    d.rename(columns={c1: 'Agent Writing Contract Start Date'}, inplace=True)
    d.rename(columns={c3: 'Agent Writing Contract Status'}, inplace=True)
    return d


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


def find_valid_pn(d):
    invalid_nums = pd.Series(map(is_invalid_pn, d['Agency Phone Number']))
    if invalid_nums.sum() > 0:
        for name in d[invalid_nums]['Agent Id']:
            logging.warning('Agent {} has an invalid phone number.'.format(name))
    else:
        logging.info('All phone numbers validated.')


def find_valid_state(d):
    states_abbr = ['AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA',
                   'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD',
                   'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
                   'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC',
                   'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY']
    is_invalid_st = lambda s, a: True if s in a else False
    invalid_state = pd.Series(map(is_invalid_st, d['Agency State'], states_abbr))
    if invalid_state.sum() > 0:
        for name in d[invalid_state]['Agent Id']:
            logging.warning('Agent {} has an invalid state.'.format(name))
    else:
        logging.info('All states validated.')


def find_valid_email(d):
    is_invalid_em = lambda e: True if not re.match(r'[\w\.-]+@[\w\.-]+(\.[\w]+)+', e) else False
    invalid_emails = pd.Series(map(is_invalid_em, d['Agent Email Address']))
    if invalid_emails.sum() > 0:
        for name in d[invalid_emails]['Agent Id']:
            logging.warning('Agent {} has an invalid email.'.format(name))
    else:
        logging.info('All emails validated.\n')


def data_summary(d):
    trans_d = d.T
    logging.info('Transposed Data:\n')
    logging.info(trans_d)
    logging.info('\n')

    state_group = d.sort_values('Agency State')
    logging.info('Data grouped by state:\n')
    logging.info(state_group)
    logging.info('\n')

    state_counts = state_group.groupby('Agency State')['Agent Id'].count().sort_values()
    plt.plot(state_counts.index, state_counts, '.r')
    plt.xticks(np.arange(0, 50, 3))
    plt.xlabel('State')
    plt.ylabel('Count')
    plt.title('Occurence of Agency States')
    plt.savefig('graphs/State_counts.png')

    format_middle_names = lambda n: n if (n == '') else (n + ' ')
    first_names = d['Agent First Name'].str.strip().str.title() + ' '
    middle_names = pd.Series(map(format_middle_names, d['Agent Middle Name'].str.strip().str.title()))
    last_names = d['Agent Last Name'].str.strip().str.title()
    names = first_names + middle_names + last_names

    c1 = 'Agent Writing Contract Start Date'
    c2 = 'Date when an agent became A2O'
    agent_info = pd.concat([names, d[[c1, c2]]], axis=1)
    logging.info('Agent Start and A2O Date Info:\n')
    logging.info(agent_info)
    logging.info('\n')
    agent_info[c1] = pd.to_datetime(agent_info[c1], 'coerce')
    agent_info[c2] = pd.to_datetime(agent_info[c2], 'coerce')

    agent_info = agent_info.sort_values(c2)
    fig, ax1 = plt.subplots()
    color = 'tab:red'
    ax1.set_ylabel('A2O Date', color=color)
    ax1.plot(agent_info[0], agent_info[c2], '.r')
    ax1.tick_params(axis='y', labelcolor=color)
    ax2 = ax1.twinx()
    color = 'tab:blue'
    ax2.set_ylabel('Start Date', color=color)
    ax2.plot(agent_info[0], agent_info[c1], '.b')
    ax2.tick_params(axis='y', labelcolor=color)
    plt.xticks([])
    fig.tight_layout()
    plt.savefig('graphs/Agent_info_by_A2O.png')

    agent_info = agent_info.sort_values(c1)
    fig, ax1 = plt.subplots()
    color = 'tab:red'
    ax1.set_ylabel('A2O Date', color=color)
    ax1.plot(agent_info[0], agent_info[c2], '.r')
    ax1.tick_params(axis='y', labelcolor=color)
    ax2 = ax1.twinx()
    color = 'tab:blue'
    ax2.set_ylabel('Start Date', color=color)
    ax2.plot(agent_info[0], agent_info[c1], '.b')
    ax2.tick_params(axis='y', labelcolor=color)
    plt.xticks([])
    fig.tight_layout()
    plt.savefig('graphs/Agent_info_by_start.png')


def send_email(cond):
    if cond:
        logging.info('Program executed successfully')
    else:
        logging.info('Program failed execution')

    with open('NYLData.log') as fp:
        msg = EmailMessage()
        msg.set_content(fp.read())

    username = 'lucas.invernizzi@gmail.com'
    with open('conf.txt') as fp:
        password = fp.read()
    smtp_server = "smtp.gmail.com:587"

    msg['Subject'] = 'Results of NYL data analysis'
    msg['From'] = username
    msg['To'] = username
    if cond:
        path = 'graphs/'
        imgs = ['Agent_info_by_A2O.png', 'Agent_info_by_start.png', 'State_counts.png']
        for img_name in imgs:
            with open(path+img_name, 'rb') as fp:
                img = fp.read()
            msg.add_attachment(img, maintype='image', subtype=imghdr.what(None, img))

    s = smtplib.SMTP(smtp_server)
    s.starttls()
    s.login(username, password)
    s.send_message(msg)
    s.quit()
    if cond:
        exit(0)
    else:
        exit(-1)


if __name__ == "__main__":
    logging.basicConfig(filename='NYLData.log', filemode='w', level=logging.DEBUG)
    mpl_logger = logging.getLogger('matplotlib')
    mpl_logger.setLevel(logging.WARNING)
    file_path = 'data/'

    file_names = np.array(os.listdir(file_path))
    recent_file = find_recent_file(file_names)

    # log_process(recent_file)

    prev_file_names = np.delete(file_names, np.where(file_names == recent_file))
    prev_file = find_recent_file(prev_file_names)

    data = validate_file_len(recent_file, prev_file, file_path)
    data = replace_headers(data)

    try:
        find_valid_pn(data)
        find_valid_state(data)
        find_valid_email(data)
    except:
        logging.error('Data validation failed.')
        send_email(False)

    try:
        data_summary(data)
    except:
        logging.error('Data summarization failed.')
        send_email(False)
    else:
        send_email(True)

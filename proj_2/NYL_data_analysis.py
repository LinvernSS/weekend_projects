import os  # obtain list of files in directory
import re  # email validation
import imghdr  # image attachment
import smtplib  # sending emails
import logging
import numpy as np
import pandas as pd
import phonenumbers as pn  # phone number validation
import matplotlib.pyplot as plt
from email.message import EmailMessage

states_abbr = ['AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA',
               'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD',
               'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
               'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC',
               'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY']


# Given a list of filenames, returns file with most recent date
# Assumed all files in list are same format
def find_recent_file(fns):
    # extracts last part of file name minus extension
    remove_ext = np.vectorize(lambda f: f.split('_')[-1][:-4])

    if type(fns) == np.ndarray:  # must be a list of files
        if len(fns[0]) > 4:  # cannot be just an extension
            dates = remove_ext(fns)
            rf_index = dates.argmax()  # index of recent file
            rf = fns[rf_index]
            return rf
    logging.error('One or more files is in the wrong format.')
    return False


# Creates/reads from NYL.lst, checks if given filename is logged there.
# Stops process if logged, logs and continues otherwise
def log_process(f):
    try:
        file_list = open('output/NYL.lst', 'x')  # tries to create lst file, throws error if it already exists
        file_list.write(f + '\n')
        file_list.close()
        logging.info('Created NYL.lst and logged {} as processed.'.format(f))
        return True
    except FileExistsError:
        file_list = open('output/NYL.lst', 'r')  # opens as read only
        if f in file_list.read().split('\n'):  # checks if given file name is in lst file
            file_list.close()
            logging.warning('File already processed.')
            return False
        else:
            file_list.close()
            file_list = open('output/NYL.lst', 'a')  # opens as append to add current file
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
    if c1 in d.columns:
        d = d.rename(columns={c1: 'Agent Writing Contract Start Date'})
    if c3 in d.columns:
        d = d.rename(columns={c3: 'Agent Writing Contract Status'})
    return d


# Removes unnecessary whitespace from given dataframe
def format_data(d):
    try:
        # removes leading and trailing whitespace, as well as multiple spaces in middle of object columns
        d = d.apply(lambda col: col.str.strip().replace({' +': ' '}, regex=True) if col.dtype == object else col)
    except Exception as ex:
        logging.error('Failed formatting data. {}'.format(ex))
    finally:
        return d


# Returns true if given US phone number is valid, else false
def is_invalid_pn(num):
    # ignore blank phone numbers
    if num.strip() == '':
        return False
    num = '+1' + num  # assumes all phone numbers are from the USA
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
    invalid_agency_nums = d['Agency Phone Number'].apply(is_invalid_pn)
    invalid_agent_nums = d['Agent Phone Number'].apply(is_invalid_pn)
    # boolean mask which is true wherever there is an invalid phone number
    invalid_nums = invalid_agency_nums | invalid_agent_nums

    if invalid_nums.sum() > 0:  # if there are any invalid phone numbers, log that agent's id
        for agent_id in d[invalid_nums]['Agent Id']:
            logging.warning('Agent {} has an invalid phone number.'.format(agent_id))
        return False
    else:
        logging.info('All phone numbers validated.')
        return True


# Finds if there are invalid states in list
def invalid_list_states(s):
    for st in s[:-1].split(','):
        if st not in states_abbr:
            return True
    return False


# Validates all agency states in given dataframe, logs invalid states
def find_valid_state(d):
    # boolean masks which are true wherever there is an invalid state
    invalid_agency_states = d['Agency State'].apply(lambda s: False if s in states_abbr else True)
    invalid_agent_states = d['Agent State'].apply(lambda s: False if s in states_abbr else True)
    invalid_license_states = d['Agent License State (active)'].fillna('').apply(invalid_list_states)

    # if any are invalid, there is an invalid state
    invalid_states = invalid_agent_states | invalid_agency_states | invalid_license_states

    if invalid_states.sum() > 0:  # if there are any invalid states, log that agent's id
        for agent_id in d[invalid_states]['Agent Id']:
            logging.warning('Agent {} has an invalid state.'.format(agent_id))
        return False
    else:
        logging.info('All states validated.')
        return True


# Validates all agent's email addresses in given dataframe, logs invalid emails
def find_valid_email(d):
    # regex check for invalid email. returns true if invalid
    is_invalid_em = lambda em: True if not re.match(r'[\w\.-]+@[\w\.-]+(\.[\w]+)+', em) else False
    # boolean mask which is true for every invalid email
    invalid_emails = d['Agent Email Address'].apply(is_invalid_em)

    if invalid_emails.sum() > 0:  # if there are any invalid emails, log that agent's id
        for agent_id in d[invalid_emails]['Agent Id']:
            logging.warning('Agent {} has an invalid email.'.format(agent_id))
        return False
    else:
        logging.info('All emails validated.')
        return True


# Plots column data
def column_data(d):
    n_unique_cols = d.apply(lambda col: col.nunique())
    logging.info('Number of Unique Values by Column:\n\n{}\n'.format(n_unique_cols))
    # plot unique columns
    fig, ax = plt.subplots()
    ax.bar(n_unique_cols.index, n_unique_cols)
    plt.setp(ax.get_xticklabels(), rotation=90, horizontalalignment='right')
    plt.xlabel('Column')
    plt.ylabel('Count')
    plt.title('Number of Unique Values by Column')
    fig.set_figheight(8)
    fig.set_figwidth(8)
    plt.tight_layout()
    plt.savefig('output/graphs/Unique_cols.png')


# Plots state data
def state_data(d):
    state_group = d.sort_values('Agency State')[['Agent Id', 'Agency State']]
    logging.info('Data grouped by state:\n\n{}\n'.format(state_group))
    # plot state counts
    state_counts = state_group.groupby('Agency State')['Agent Id'].count().sort_values()
    fig, ax = plt.subplots()
    ax.bar(state_counts.index, state_counts)
    plt.xlabel('State')
    plt.ylabel('Count')
    plt.title('Occurence of Agency States')
    fig.set_figheight(5)
    fig.set_figwidth(14)
    plt.tight_layout()
    plt.savefig('output/graphs/State_counts.png')


# Combines and formats first, middle and last names from given dataframe
def format_agent_names(d):
    first_names = d['Agent First Name'].str.title() + ' '
    middle_names = d['Agent Middle Name'].str.title().apply(lambda n: n if (n == '') else (n + ' '))
    last_names = d['Agent Last Name'].str.title()

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

    plt.xticks([])  # there are 2000+ names -- don't plot them
    fig.tight_layout()
    plt.savefig('output/graphs/' + n + '.png')


# Summarizes, logs, and plots agent info data from given dataframe
def agent_info_data(d):
    names = format_agent_names(d)
    col1 = 'Agent Writing Contract Start Date'
    col2 = 'Date when an agent became A2O'

    # aggregate necessary data
    agent_info = pd.concat([names, d[[col1, col2]]], axis=1)
    logging.info('Agent Start and A2O Date Info:\n\n{}\n'.format(agent_info))

    # convert to datetime so they can be sorted
    agent_info[col1] = pd.to_datetime(agent_info[col1], 'coerce')
    agent_info[col2] = pd.to_datetime(agent_info[col2], 'coerce')

    # plot with sorted by A20 date
    agent_info = agent_info.sort_values(col2)
    img_name = 'Agent_info_by_A2O'
    plot_agent_info(agent_info, col1, col2, img_name)

    # plot with sorted by start date
    agent_info = agent_info.sort_values(col1)
    img_name = 'Agent_info_by_start'
    plot_agent_info(agent_info, col1, col2, img_name)


def data_summary(d):
    logging.info('Starting data summary.')

    logging.info('Transposed Data:\n\n{}\n'.format(d.T))

    column_data(d)
    state_data(d)
    agent_info_data(d)

    logging.info('Data summary complete.')


# sends email of log file, attaches images if cond is true, exits program when complete
def send_email(success):
    if success:
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

    if success:
        imgs = os.listdir('output/graphs/')
        for img_name in imgs:
            with open('output/graphs/' + img_name, 'rb') as fp:
                img = fp.read()
            msg.add_attachment(img, maintype='image', subtype=imghdr.what(None, img), filename=img_name)

    s = smtplib.SMTP(smtp_server)
    s.starttls()
    s.login(username, password)
    # s.send_message(msg)
    s.quit()

    exit(0) if success else exit(-1)


# Function to send email if process failed
def escape(b):
    if type(b) == bool:
        if not b:
            send_email(False)
    return b


if __name__ == "__main__":
    pd.set_option('display.expand_frame_repr', False)
    pd.set_option('display.max_columns', 3)
    # Logging setup
    logging.basicConfig(format='%(asctime)s %(levelname)-4s %(message)s',
                        filename='output/NYLData.log',
                        filemode='w',
                        level=logging.DEBUG,
                        datefmt='%Y-%m-%d %H:%M:%S')
    # Remove matplotlib logging
    mpl_logger = logging.getLogger('matplotlib')
    mpl_logger.setLevel(logging.WARNING)

    # Get list of files in data folder
    file_names = np.array(os.listdir('data/'))

    recent_file = escape(find_recent_file(file_names))
    # escape(log_process(recent_file))

    data = escape(load_data(recent_file))

    # Gets list of files without most recent file
    prev_file_names = np.delete(file_names, np.where(file_names == recent_file))
    prev_file = escape(find_recent_file(prev_file_names))
    prev_data = escape(load_data(prev_file))

    escape(validate_file_len(data, prev_data))
    data = format_data(replace_headers(data))

    find_valid_pn(data)
    find_valid_state(data)
    find_valid_email(data)

    data_summary(data)
    send_email(True)

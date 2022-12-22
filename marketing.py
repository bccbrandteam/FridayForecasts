# Import dependencies and libraries
import os
import glob
import requests
import pandas
from time import sleep
import pymysql
import sys
import xlsxwriter
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

from custom_functions import Scraper, mysqlalchemy, mysqlcursor
from custom_functions import utc_converter
# ===========================================================
from selenium.webdriver.support.ui import WebDriverWait
import selenium.webdriver.support.expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import datetime as DT

HOME_DIR = '/home/pi/Desktop/CodeFiles/FridayForecasts'
TEMP = '/home/pi/Desktop/CodeFiles/FridayForecasts/temp'
FINAL_FILENAME = 'posting_download*.csv'
FINAL_FILEPATH = os.path.join(TEMP, FINAL_FILENAME)

#Clear temp folders
for f in os.listdir(TEMP):
    os.remove(os.path.join(TEMP, f))

if (os.path.exists('/home/pi/Desktop/CodeFiles/FridayForecasts/temp/marketing_forecast.xlsx') == True):
    os.chmod('/home/pi/Desktop/CodeFiles/FridayForecasts/temp/marketing_forecast.xlsx', 0o777)
    os.remove('/home/pi/Desktop/CodeFiles/FridayForecasts/temp/marketing_forecast.xlsx')
print('Temp folder cleared')

# Selenium web scraper
try:
    # Initiate an instance of the Scraper class
    scraper = Scraper().driver

    # Login to Handshake
    scraper.get('https://app.joinhandshake.com/login')
    (scraper.find_element('xpath', '//*[@id="email-address-identifier"]')).send_keys('recruitersupport@byu.edu')  # USERNAME
    (scraper.find_element('xpath', '//*[@id="ui-id-1"]/div[3]/form/div/button')).click()
    (scraper.find_element('xpath', '//*[@id="password"]')).send_keys('BCCrecruiter1!')  # PASSWORD
    (scraper.find_element('xpath', '/html/body/div[1]/div[2]/div/form/div/div/button')).click()
    # If necessary, redirect out of a Employer or Student Account
    try:
        scraper.find_element('xpath', '//*[@id="ui-id-1"]/div[1]/div[1]/a')
        scraper.get('https://app.joinhandshake.com/user_switcher/options')
        (scraper.find_element('xpath', '//*[@id="main"]/div/div/div/a')).click()
    except NoSuchElementException:
        pass
    try:
        scraper.find_element('xpath', '//*[@id="main"]/div[2]/div/div[1]/div/div/div[1]/div[1]/div/div[1]/div/div[2]/h1')
        scraper.get('https://app.joinhandshake.com/user_switcher/options')
        scraper.find_element('xpath', '//*[@id="main"]/div/div/div/a').click()
    except NoSuchElementException:
        pass

    # Get the jobs report from Handshake
    while True:
        print('Logged in to Handshake.byu.edu')
        wait = WebDriverWait(scraper, 100)
        today = DT.date.today()
        week_ago = today - DT.timedelta(days=7)
        print(week_ago)

        # ALL COMPANIES VV
        scraperurl = 'https://app.joinhandshake.com/edu/postings?page=1&per_page=100&sort_direction=desc&sort_column=expires&status%5B%5D=approved&job.job_role_groups%5B%5D=470&job.job_role_groups%5B%5D=441&job.job_role_groups%5B%5D=440&apply_start_begin={}T06%3A00%3A00.000Z'.format(week_ago)
        # ONLY INCLUDING COMPANIES THAT RECRUIT ON CAMPUS VV
        # scraperurl = 'https://app.joinhandshake.com/edu/postings?page=1&per_page=100&sort_direction=desc&sort_column=expires&status%5B%5D=approved&employers%5B%5D=8440&employers%5B%5D=8480&employers%5B%5D=8520&employers%5B%5D=8530&employers%5B%5D=8544&employers%5B%5D=8559&employers%5B%5D=8594&employers%5B%5D=8598&employers%5B%5D=8750&employers%5B%5D=8757&employers%5B%5D=8759&employers%5B%5D=8932&employers%5B%5D=8967&employers%5B%5D=9073&employers%5B%5D=9111&employers%5B%5D=9160&employers%5B%5D=9196&employers%5B%5D=9206&employers%5B%5D=9226&employers%5B%5D=9294&employers%5B%5D=9563&employers%5B%5D=9643&employers%5B%5D=9863&employers%5B%5D=9977&employers%5B%5D=10145&employers%5B%5D=10505&employers%5B%5D=10614&employers%5B%5D=10706&employers%5B%5D=10796&employers%5B%5D=10870&employers%5B%5D=11054&employers%5B%5D=11332&employers%5B%5D=11357&employers%5B%5D=11363&employers%5B%5D=11455&employers%5B%5D=11614&employers%5B%5D=11617&employers%5B%5D=11693&employers%5B%5D=11916&employers%5B%5D=12273&employers%5B%5D=12493&employers%5B%5D=12679&employers%5B%5D=12910&employers%5B%5D=12914&employers%5B%5D=12960&employers%5B%5D=12999&employers%5B%5D=13107&employers%5B%5D=13202&employers%5B%5D=13337&employers%5B%5D=13377&employers%5B%5D=13515&employers%5B%5D=13541&employers%5B%5D=13661&employers%5B%5D=13719&employers%5B%5D=13758&employers%5B%5D=13861&employers%5B%5D=14034&employers%5B%5D=14072&employers%5B%5D=14106&employers%5B%5D=14185&employers%5B%5D=14186&employers%5B%5D=14368&employers%5B%5D=14388&employers%5B%5D=14392&employers%5B%5D=14462&employers%5B%5D=14561&employers%5B%5D=14620&employers%5B%5D=14624&employers%5B%5D=14634&employers%5B%5D=14826&employers%5B%5D=14929&employers%5B%5D=14974&employers%5B%5D=15059&employers%5B%5D=15107&employers%5B%5D=15216&employers%5B%5D=15427&employers%5B%5D=15554&employers%5B%5D=15623&employers%5B%5D=15769&employers%5B%5D=15783&employers%5B%5D=15815&employers%5B%5D=15978&employers%5B%5D=16317&employers%5B%5D=16337&employers%5B%5D=16363&employers%5B%5D=16447&employers%5B%5D=16455&employers%5B%5D=16612&employers%5B%5D=16795&employers%5B%5D=16917&employers%5B%5D=17051&employers%5B%5D=17150&employers%5B%5D=17161&employers%5B%5D=17427&employers%5B%5D=17674&employers%5B%5D=17717&employers%5B%5D=17767&employers%5B%5D=17770&employers%5B%5D=18288&employers%5B%5D=18383&employers%5B%5D=18668&employers%5B%5D=18699&employers%5B%5D=19044&employers%5B%5D=19164&employers%5B%5D=19207&employers%5B%5D=19217&employers%5B%5D=19389&employers%5B%5D=19455&employers%5B%5D=19647&employers%5B%5D=19688&employers%5B%5D=19885&employers%5B%5D=19933&employers%5B%5D=20095&employers%5B%5D=20232&employers%5B%5D=20238&employers%5B%5D=20535&employers%5B%5D=20675&employers%5B%5D=20733&employers%5B%5D=21056&employers%5B%5D=21399&employers%5B%5D=21703&employers%5B%5D=22003&employers%5B%5D=22160&employers%5B%5D=22186&employers%5B%5D=22447&employers%5B%5D=22706&employers%5B%5D=22842&employers%5B%5D=23187&employers%5B%5D=23397&employers%5B%5D=23444&employers%5B%5D=23923&employers%5B%5D=24119&employers%5B%5D=24370&employers%5B%5D=24481&employers%5B%5D=24541&employers%5B%5D=25264&employers%5B%5D=25724&employers%5B%5D=26567&employers%5B%5D=28104&employers%5B%5D=28461&employers%5B%5D=29088&employers%5B%5D=30326&employers%5B%5D=30881&employers%5B%5D=31675&employers%5B%5D=32781&employers%5B%5D=32868&employers%5B%5D=33011&employers%5B%5D=33037&employers%5B%5D=33968&employers%5B%5D=33984&employers%5B%5D=34148&employers%5B%5D=34424&employers%5B%5D=34768&employers%5B%5D=34790&employers%5B%5D=34874&employers%5B%5D=35226&employers%5B%5D=35284&employers%5B%5D=36518&employers%5B%5D=36945&employers%5B%5D=36986&employers%5B%5D=37052&employers%5B%5D=37201&employers%5B%5D=38203&employers%5B%5D=38531&employers%5B%5D=38601&employers%5B%5D=38858&employers%5B%5D=39013&employers%5B%5D=39835&employers%5B%5D=42311&employers%5B%5D=42334&employers%5B%5D=43232&employers%5B%5D=44145&employers%5B%5D=45209&employers%5B%5D=46352&employers%5B%5D=47038&employers%5B%5D=48616&employers%5B%5D=51452&employers%5B%5D=52387&employers%5B%5D=52602&employers%5B%5D=55071&employers%5B%5D=56012&employers%5B%5D=57414&employers%5B%5D=60457&employers%5B%5D=60664&employers%5B%5D=60983&employers%5B%5D=61876&employers%5B%5D=63123&employers%5B%5D=63655&employers%5B%5D=64503&employers%5B%5D=68172&employers%5B%5D=71978&employers%5B%5D=72463&employers%5B%5D=80179&employers%5B%5D=81702&employers%5B%5D=82683&employers%5B%5D=85786&employers%5B%5D=88972&employers%5B%5D=91483&employers%5B%5D=94828&employers%5B%5D=97822&employers%5B%5D=98123&employers%5B%5D=106922&employers%5B%5D=109871&employers%5B%5D=111512&employers%5B%5D=113906&employers%5B%5D=127020&employers%5B%5D=130683&employers%5B%5D=133187&employers%5B%5D=140505&employers%5B%5D=142009&employers%5B%5D=148677&employers%5B%5D=148824&employers%5B%5D=151816&employers%5B%5D=158444&employers%5B%5D=160307&employers%5B%5D=161553&employers%5B%5D=162791&employers%5B%5D=165015&employers%5B%5D=170012&employers%5B%5D=170129&employers%5B%5D=170418&employers%5B%5D=171104&employers%5B%5D=171332&employers%5B%5D=171474&employers%5B%5D=173288&employers%5B%5D=174588&employers%5B%5D=175123&employers%5B%5D=176306&employers%5B%5D=180159&employers%5B%5D=186008&employers%5B%5D=187653&employers%5B%5D=188979&employers%5B%5D=189411&employers%5B%5D=194158&employers%5B%5D=203073&employers%5B%5D=204386&employers%5B%5D=205305&employers%5B%5D=209881&employers%5B%5D=215534&employers%5B%5D=220261&employers%5B%5D=222297&employers%5B%5D=225613&employers%5B%5D=226511&employers%5B%5D=227469&employers%5B%5D=228883&employers%5B%5D=233372&employers%5B%5D=239295&employers%5B%5D=239995&employers%5B%5D=242804&employers%5B%5D=243692&employers%5B%5D=245865&employers%5B%5D=248160&employers%5B%5D=253628&employers%5B%5D=258120&employers%5B%5D=269328&employers%5B%5D=271794&employers%5B%5D=275065&employers%5B%5D=286505&employers%5B%5D=287504&employers%5B%5D=287762&employers%5B%5D=295966&employers%5B%5D=304778&employers%5B%5D=308772&employers%5B%5D=313769&employers%5B%5D=313819&employers%5B%5D=316095&employers%5B%5D=331798&employers%5B%5D=332011&employers%5B%5D=343930&employers%5B%5D=356797&employers%5B%5D=359860&employers%5B%5D=369418&employers%5B%5D=370063&employers%5B%5D=375172&employers%5B%5D=377206&employers%5B%5D=378672&employers%5B%5D=384772&employers%5B%5D=395724&employers%5B%5D=398676&employers%5B%5D=409978&employers%5B%5D=415689&employers%5B%5D=425812&employers%5B%5D=447595&employers%5B%5D=461043&employers%5B%5D=468723&employers%5B%5D=469892&employers%5B%5D=489649&employers%5B%5D=490583&employers%5B%5D=492639&employers%5B%5D=515017&employers%5B%5D=521770&employers%5B%5D=556463&employers%5B%5D=574164&employers%5B%5D=598209&employers%5B%5D=601160&employers%5B%5D=628469&employers%5B%5D=628750&employers%5B%5D=639614&employers%5B%5D=641955&employers%5B%5D=642126&employers%5B%5D=659114&employers%5B%5D=664587&employers%5B%5D=674389&employers%5B%5D=693936&employers%5B%5D=713977&employers%5B%5D=716737&employers%5B%5D=737560&employers%5B%5D=745812&employers%5B%5D=750218&employers%5B%5D=793163&employers%5B%5D=794003&employers%5B%5D=799712&employers%5B%5D=814880&employers%5B%5D=826490&employers%5B%5D=833145&employers%5B%5D=15127&employers%5B%5D=22992&employers%5B%5D=243323&employers%5B%5D=817366&employers%5B%5D=519499&employers%5B%5D=834129&employers%5B%5D=504486&employers%5B%5D=283039&employers%5B%5D=573672&employers%5B%5D=343653&employers%5B%5D=45688&employers%5B%5D=18613&employers%5B%5D=19453&employers%5B%5D=97326&employers%5B%5D=235826&employers%5B%5D=297042&employers%5B%5D=152544&employers%5B%5D=699313&employers%5B%5D=586634&employers%5B%5D=61112&employers%5B%5D=615904&employers%5B%5D=12711&employers%5B%5D=17930&employers%5B%5D=762812&employers%5B%5D=403218&employers%5B%5D=815780&employers%5B%5D=802364&employers%5B%5D=540550&job.job_role_groups%5B%5D=440&job.job_role_groups%5B%5D=470&job.job_role_groups%5B%5D=441&apply_start_begin={}T06%3A00%3A00.000Z'.format(week_ago)
        
        scraper.get(scraperurl)
        print('Navigated to the jobs page on Handshake')                                                                    
        sleep(20)
        
        wait.until(ec.element_to_be_clickable((By.XPATH, '//*[@id="main"]/div/div/div[1]/div/form/div[1]/div/div[1]/div[2]/div/button'))).click()
        print('Began downloading the jobs report')
        count = 0
        while not glob.glob(os.path.join(TEMP, 'posting_download*.csv')) and count < 6000:
            count += 1
            sleep(1)
        if count == 6000:
            raise FileNotFoundError('The report could not be downloaded')
        break

    scraper.close()

except Exception as e:
    scraper.close()
    raise e

# API call
HEADERS = {
        'Authorization': 'Token token="d3d7a166e273c5c3d3ebd0f89a45ecc7"',
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache--url',
    }
page_number = 0
full_response = list()

print('Starting Handshake API requests for jobs data')
while True:
    url = 'https://app.joinhandshake.com/api/v1/jobs?page={}&per_page=50&sort_direction=desc&sort_column=start_date'.format(page_number)
    r = requests.get(url, headers=HEADERS)
    if r.status_code != 200:
        continue
    r = r.json()
    jobs = r['jobs']
    if len(jobs) == 0:
        break
    for job in jobs:
        single_job = list()
        single_job.append(job['id'])
        single_job.append(job['updated_at'])
        single_job.append(job['created_at'])
        single_job.append(job['description'])
        # ADDITIONAL WANTED FIELDS FROM API GO ABOVE HERE
        full_response.append(single_job)
    page_number += 1
    print('page {} recorded'.format(page_number))

# Merge API data and handshake download data
apiData = os.path.join(TEMP, ((glob.glob(os.path.join(TEMP, 'posting_download*.csv')))[0]))
print(apiData)
jobs_download = pandas.read_csv(apiData)
jobs_api = pandas.DataFrame(data=full_response, columns=['API_ID', 'Updated At', 'Created At', 'Description'])  # ADDITIONAL WANTED FIELDS FROM API SHOULD BE DENOTED HERE
merged = pandas.merge(jobs_download, jobs_api, how='left', left_on='Job Id', right_on='API_ID')
merged = merged.drop('API_ID', axis=1)
merged = merged.applymap(lambda x: x if isinstance(x, str) else x) #.decode('utf8')
merged = merged.dropna(axis=1, how='all')
merged = utc_converter(merged, 'US/Mountain')


# Prep data for SQL injection
job_report = merged[['Job Id', 'Title', 'Job Type', 'Employment Type', 'Employer', 'Date Posted', 'Apply Start Date', 'Expires', 'Job Location', 'Description']]
job_report.columns=['job_id', 'title', 'type', 'position_type', 'employer', 'date_posted', 'apply_start_date', 'expires', 'location', 'job_description']
job_report.reset_index()

alchemyengine = mysqlalchemy('utf8')
job_report = job_report.applymap(lambda x: x if isinstance(x, str) else x)
job_report.to_sql(name='marketing_forecast', con=alchemyengine, if_exists="replace")
print('Pushed jobs to MySQL database')

# Clear temp folder
for root, directory, files in os.walk(TEMP):
    for file in files:
        filepath = os.path.join(root, file)
        os.chmod(filepath, 0o777)
        os.remove(filepath)

# connect to mySQL server
db_opts = mysqlcursor()

db = pymysql.connect(**db_opts)
cur = db.cursor()
xlsx_file_path = '/home/pi/Desktop/CodeFiles/FridayForecasts/temp/marketing_forecast.xlsx'

try:
    # creates job_url column
    cur.execute('''ALTER TABLE marketing_forecast
    ADD url VARCHAR(60) AS (concat("https://byu.joinhandshake.com/stu/jobs/", job_id));''')

    # selects desired rows from table, reformats dates, orders by when jobs expire
    cur.execute('''select title, type, position_type, employer, DATE_FORMAT(expires, '%m/%d/%y'), DATE_FORMAT(date_posted, '%m/%d/%y'), DATE_FORMAT(apply_start_date, '%m/%d/%y'),
    location, url from marketing_forecast
    where job_description like '%marketing%'
    order by expires ASC
    ;''')

    rows = cur.fetchall()
finally:
    db.close()

# Continue only if there are rows returned.
if rows:
    # holds job info
    result = list()

    # grab headers
    column_names = list()
    for i in cur.description:
        column_names.append(i[0])

    # grab job info
    for row in rows:
        result.append(row)

    # write xlsx file
    wb = xlsxwriter.Workbook('/home/pi/Desktop/CodeFiles/FridayForecasts/temp/marketing_forecast.xlsx')
    ws = wb.add_worksheet('data')
    col = 0
    color = True
    
    with open(xlsx_file_path, 'w', newline='') as f:
        # set headers
        format = wb.add_format({'bg_color': '#006DC0', 'font_color': 'white', 'bold': True})
        for column in column_names:
            ws.write('A1', 'Title', format)
            ws.write('B1', 'Type', format)
            ws.write('C1', 'Position Type', format)
            ws.write('D1', 'Employer', format)
            ws.write('E1', 'Expires', format)
            ws.write('F1', 'Date Posted', format)
            ws.write('G1', 'Apply Start', format)
            ws.write('H1', 'Location', format)
            ws.write('I1', 'Handshake URL', format)
            col += 1

        # fill in other information
        for r, row in enumerate(result):
            for c in range(col):
                # add formatting, like alternating colors
                if color is True:
                    format = wb.add_format({'bg_color': '#96C2E4', 'text_wrap': True, 'align': 'top'})
                    # color is dark blue
                else:
                    format = wb.add_format({'bg_color': '#C3D4E1', 'text_wrap': True, 'align': 'top'})
                    # color is light blue
                ws.write((r+1), c, row[c], format)
                ws.set_row(r+1, 30)

            # set colors to alternate
            if color is True:
                color = False
            else:
                color = True

        # set column sizes
        ws.set_column(0, 0, 31)
        ws.set_column(1, 1, 18)
        ws.set_column(2, 2, 12)
        ws.set_column(3, 3, 31)
        ws.set_column(4, 6, 12)
        ws.set_column(7, 7, 21)
        ws.set_column(8, 8, 42)

    wb.close()
    
else:
    sys.exit("No rows found for query: {}".format('SELECT * from iscareer_operations.marketing_forecast'))

print("Processing excel document...")
while(os.path.isfile('/home/pi/Desktop/CodeFiles/FridayForecasts/temp/marketing_forecast.xlsx') == False):
    sleep(10)
print("Excel document created.")

print("Emailing excel document out...")

# Send email with xlsx file
os.walk('/home/pi/Desktop/CodeFiles/FridayForecasts/temp/marketing_forecast.xlsx')
def send_mail(send_from,send_to,subject,text,server,port,username='bcc.notification.noreply@gmail.com',password='pkkqrzkfwtvlsgqg',isTls=True):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open("/home/pi/Desktop/CodeFiles/FridayForecasts/temp/marketing_forecast.xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="marketing_forecast.xlsx"')
    msg.attach(part)

    #SSL connection only working on Python 3+
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login(username, password)
        server.login(username,password)
        server.sendmail(send_from, send_to, msg.as_string())
        server.quit()

today = DT.date.today()
next_forecast = today + DT.timedelta(days=3)
subjectLine = '{} MARKETING FORECAST'.format(next_forecast)

# Add additional emails here
send_mail('bccdataanalytics@gmail.com', 'bccdataanalytics@gmail.com', subjectLine,
            'Here are the marketing jobs for the week!', 'smtp.gmail.com', '465')
send_mail('bccdataanalytics@gmail.com', 'olivia.davis2234@gmail.com', subjectLine,
            'Here are the marketing jobs for the week!', 'smtp.gmail.com', '465')

print("Excel document emailed.")
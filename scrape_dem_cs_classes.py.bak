# user input automation and scraping
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

# for sleep() and date/time file naming
import time
from datetime import datetime

# for creating and modifying xlsx files with data scraped
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# for file directory
import os

# for auto-mailing
import smtplib

term = ''
term_options = []
show_open = ''

# data buffers to store information scraped
cs_title_list = []
cs_section_list = []
cs_DnT_list = []
cs_availability_list = []
cs_instructors_list = []
cs_mode_list = []

def autoPage1():
    print("Page #1:")
    time.sleep(3)

    # Page 1
    # Institution
    driver.find_element(By.CSS_SELECTOR, '#CTY01').click()
    print("clicked City College of New York")
    time.sleep(1)
    # Term: choose SPRING[4], SUMMER[3], FALL[2]
    if term.lower() == 'spring':
        term_options.append(4)
        term_options.append(21)
    elif term.lower() == 'summer':
        term_options.append(3)
        term_options.append(14)
    elif term.lower() == 'fall':
        term_options.append(2)
        term_options.append(19)
    driver.find_element(By.XPATH, '//*[@id="t_pd"]/option[{}]'
                        .format(term_options[0])).click()
    print("clicked", term.upper(), "semester")
    time.sleep(1)
    # Next button
    driver.find_element(By.CLASS_NAME, 'SSSBUTTON_CONFIRMLINK').click()
    print("clicked 'NEXT' button" + '\n')
def autoPage2():
    print("Page #2:")
    time.sleep(3)

    # Page 2
    # Subject: SPRING[21], SUMMER[14], FALL[19]
    driver.find_element(By.XPATH, '//*[@id="subject_ld"]/option[{}]'
                        .format(term_options[1])).click()
    print("clicked Computer Science subject")
    time.sleep(1)
    # Course Career
    driver.find_element(By.XPATH, '//*[@id="courseCareerId"]/option[4]').click()
    print("clicked Undergraduate")
    time.sleep(1)
    # Shown Open Classes Only
    if show_open.lower() == 'no':
        driver.find_element(By.XPATH, '//*[@id="open_classId"]').click()
        print("show ALL Classes")
        time.sleep(1)
    else:
        print("show OPEN Classes only")
    # Search
    driver.find_element(By.ID, 'btnGetAjax').click()
    print("clicked 'SEARCH' button" + '\n')
def autoPage3():
    # Page 3
    # Show courses
    print("Page #3:")
    time.sleep(3)
    driver.find_element(By.ID, 'imageDivLink_inst0').click()
    print("showing ALL Courses..." + '\n')
    time.sleep(1)
    scrapeSearch()
    now = datetime.now()
    current_date_time = now.strftime("%H:%M:%S")
    print("Time courses were scraped:", current_date_time, '\n')

def scrapeSearch():
    # Show all classes
    i = 0 # courses
    j = 2 # classes in a course. 2 is starting index of row
    # Course/Class counters
    cs_courses = 1 
    cs_classes = 0
    cs_classes_sum = 0
    while True:
        try:
            # Clicks course drop down
            driver.find_element(By.ID, 'imageDivLink{}'.format(i)).click()
            while True:
                try: 
                    # Find class' title
                    cs_title = driver.find_element(By.XPATH, '//*[@id="contentDivImg_inst0"]/table[{}]/tbody/tr/td/b/span[@class="cunylite_LABEL"]'
                                                .format(cs_courses)).text
                    # Find class' section
                    cs_section = driver.find_element(By.XPATH, '//*[@id="contentDivImg{}"]/table/tbody/tr[{}]/td[3]'
                                                    .format(i,j)).text
                    # Find class' Days & Times
                    cs_DnT = driver.find_element(By.XPATH, '//*[@id="contentDivImg{}"]/table/tbody/tr[{}]/td[4]'
                                                    .format(i,j)).text
                    # Find class' availability
                    cs_availability = driver.find_element(By.XPATH, '//*[@id="contentDivImg{}"]/table/tbody/tr[{}]/td[9]/img'
                                                    .format(i,j))
                    availability = cs_availability.get_attribute('title')
                    # Find class' instructors
                    cs_instructors = driver.find_element(By.XPATH, '//*[@id="contentDivImg{}"]/table/tbody/tr[{}]/td[6]'
                                                    .format(i,j)).text
                    # Find class' instruction mode
                    cs_mode = driver.find_element(By.XPATH, '//*[@id="contentDivImg{}"]/table/tbody/tr[{}]/td[7]'
                                                    .format(i,j)).text
                    # if availability == 'Open':
                    cs_title_list.append(cs_title)
                    cs_section_list.append(cs_section)
                    cs_DnT_list.append(cs_DnT)
                    cs_availability_list.append(availability)
                    cs_instructors_list.append(cs_instructors)
                    cs_mode_list.append(cs_mode)
                    # Next class
                    cs_classes += 1
                    cs_classes_sum += 1
                    j += 1
                except NoSuchElementException:
                    # print(cs_classes, 'classes founded.')
                    break
            print(cs_classes, 'class(es) found in:' + cs_title)
            i += 1  # Next course
            cs_courses += 1 # next course title
            j = 2   # reset starting index of classes row
            cs_classes = 0
        except NoSuchElementException:
            print('\n' + "Finished collecting all CS courses data.")
            print(i, 'CS Courses found.')
            print(cs_classes_sum, 'CS Class Sections in total.')
            break

def importXLSX():
    # Create a new workbook
    workbook = openpyxl.Workbook()
    # Select the active worksheet
    worksheet = workbook.active

    # Set the header row
    header = ["Course", "Section", "Days & Time", "Availability", "Instructors", "Mode"]
    worksheet.append(header)
    # Add the data to rows after header
    for i in range(len(cs_title_list)):
        row = [cs_title_list[i], cs_section_list[i], cs_DnT_list[i], cs_availability_list[i], cs_instructors_list[i], cs_mode_list[i]]
        worksheet.append(row)

    # Get current time and date
    now = datetime.now()
    current_date_time = now.strftime("%Y%m%d-%H%M")

    # set the file's name
    xlsx_f = str(current_date_time) + "_" + term.upper() + "_CCNY_CS_Classes.xlsx"
    # specify directory/file path
    dir = 'C:/Users/Zed/Documents/Code/repos/cunyf_enrollme/class-status-logs/'
    # set date-specific folder name
    current_date = now.strftime("%Y%m%d")
    date_dir = os.path.join(dir, current_date)
    # make date specific directory
    if os.path.exists(date_dir):
        print("Folder /" + current_date_time + "/ already exists" + '\n')
    else:
        print("Created folder /" + current_date_time + '/\n')
        os.makedirs(date_dir)
    fp = os.path.join(date_dir, xlsx_f)

    # Save the workbook
    workbook.save(fp)
    workbook.close()
    print("Created XLSX file named:" + '\n\t' + xlsx_f + '\n')

    wb = load_workbook(fp)
    ws = wb['Sheet']

    # update xlsx column and row sizes and color fills
    updateXLSX(ws)
    wb.save(fp)
    wb.close()
    
    # rename xlsx file if class im eyeing on is open
    found_csc = False
    found_open = False
    # write the class you want in class_i_want
    # class_i_want = "  CSC 34200 - Computer Organization"
    for row in ws.iter_rows():
        if row[0].value in class_i_want:
            found_csc = True
            print("Found" + class_i_want)
            for cell in row:
                if cell.value == "Open":
                    found_open = True
                    print(class_i_want, "is open" + '\n')
                
    if found_csc and found_open:
        new_xlsx = '~' + xlsx_f
        new_fp = os.path.join(date_dir, new_xlsx)
        os.rename(fp, new_fp)
        print("XLSX file renamed to:" + '\n\t' + new_xlsx + '\n')
    else:
        print(class_i_want, "is NOT open" + '\n')

def updateXLSX_cols_rows(ws):
    # Update column width
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 11
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 10

    # Update row height
    # enable text wrapping for the cells
    for cell in ws['C:C']:
        cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
    for cell in ws['E:E']:
        cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
    print("Columns width resized to fit")
    # set the row height to auto adjust for all rows
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
        row[0].parent.auto_size = True
    print("Rows height resized to fit")
def colorXLSX_data(ws):
    # WANTED courses      
    desired_courses = ['  CSC 30100 - Scientific Prgrmng',
                       '  CSC 30400 - Intro to Theoretical Comp Sci',
                       '  CSC 31800 - Internet Programming',
                       '  CSC 33500 - Programming Language Paradigms',
                       '  CSC 33600 - Database Systems',
                       '  CSC 34200 - Computer Organization',
                       '  CSC 34300 - Computer Systems Design Lab',
                       '  CSC 41200 - Computer Networks',
                       '  CSC 44700 - Introduction to Machine Learni',
                       '  CSC 44800 - Artificial Intelligence',
                       '  CSC 45600 - Topics in Software Engineering',
                       '  CSC 47300 - Web Site Design',
                    #    '',
                    #    '',
                    #    '',
                       ]
    # Aqua = classes i want to take
    fillAqua = PatternFill(start_color='00FFFF', end_color='00FFFF', fill_type='solid')
    for row in ws.iter_rows():
        if row[0].value in desired_courses:
            for cell in row:
                cell.fill = fillAqua

    # CURRENTLY taking courses
    current_courses = ['  CSC 34200 - Computer Organization',
                       '  CSC 34300 - Computer Systems Design Lab',
                       '  CSC 38000 - Computer Security',
                       '  CSC 32200 - Software Engineering',
                    #    '',
                       ]
    # orange = currently taking
    fillOrange = PatternFill(start_color='FF9900', end_color='FF9900', fill_type='solid')
    for row in ws.iter_rows():
        if row[0].value in current_courses:
            for cell in row:
                cell.fill = fillOrange

    # NEEDED courses      
    needed_courses = ['  CSC 30100 - Scientific Prgrmng',
                       '  CSC 30400 - Intro to Theoretical Comp Sci',
                    #    '  CSC 31800 - Internet Programming',
                       '  CSC 33500 - Programming Language Paradigms',
                       '  CSC 33600 - Database Systems',
                    #    '  CSC 34200 - Computer Organization',
                    #    '  CSC 34300 - Computer Systems Design Lab',
                       '  CSC 59866 - Senior Project I',
                       '  CSC 59867 - Senior Project II',
                    #    '',
                       ]
    # purple = classes i need to take
    fillPurple = PatternFill(start_color='800080', end_color='800080', fill_type='solid')
    for row in ws.iter_rows():
        if row[0].value in needed_courses:
            for cell in row:
                cell.fill = fillPurple

    # NEXT semester planned classes
    next_courses = [#'  CSC 30100 - Scientific Prgrmng',
                       '  CSC 30400 - Intro to Theoretical Comp Sci',
                       '  CSC 31800 - Internet Programming',
                       '  CSC 41200 - Computer Networks',
                       '  CSC 33500 - Programming Language Paradigms',
                       '  CSC 33600 - Database Systems',
                    #    '  CSC 34200 - Computer Organization',
                    #    '  CSC 34300 - Computer Systems Design Lab',
                       '  CSC 59866 - Senior Project I',
                    #    '  CSC 59867 - Senior Project II',
                    #    '',
                       ]
    # purple = classes i need to take
    fillYellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    for row in ws.iter_rows():
        if row[0].value in next_courses:
            for cell in row:
                cell.fill = fillYellow

    # TAKEN courses
    taken_courses = ['  CSC 10300 - Intrmd Cmptr Progrm',
                     '  CSC 10400 - Discrete Structrs 1',
                     '  CSC 11300 - Programming Lang',
                     '  CSC 21100 - Fundamentals Computer Systems',
                     '  CSC 21200 - Data Structures',
                     '  CSC 21700 - Prob & Stat For Csc',
                     '  CSC 22000 - Algorithms',
                     '  CSC 22100 - Software Design Laboratory',
                     '  CSC 33200 - Operating Systems',
                     ]
    # gray = <courses already taken>
    fillGray = PatternFill(start_color='B2BEB5', end_color='B2BEB5', fill_type='solid')
    # paints whole row Red if its a course already taken
    for row in ws.iter_rows():
        if row[0].value in taken_courses:
            for cell in row:
                cell.fill = fillGray

    # green = "Open", blue = "Closed"
    fillGreen = PatternFill(start_color='70bf22', end_color='70bf22', fill_type='solid')
    fillBlue = PatternFill(start_color='96bfec', end_color='96bfec', fill_type='solid')
    for row in ws.iter_rows():
        for cell in row:
            if 'Open' in cell.value:
                cell.fill = fillGreen
            elif 'Closed' in cell.value:
                cell.fill = fillBlue
    print("Cells color coded" + '\n')
def updateXLSX(ws):
    # update columns and rows
    updateXLSX_cols_rows(ws)
    # color code cells
    colorXLSX_data(ws)

    
# user input semester
while True:
    try:
        term = input("Enter SEMESTER (fall, winter, spring, summer): ")
        if term.lower() not in ["fall", "winter", "spring", "summer"]:
            raise ValueError("Invalid semester entered.")
        break
    except ValueError as e:
        print(e)

# user decides whether to show all courses, or only available courses
while True:
    try:
        show_open = input("Show OPEN classes ONLY (yes/no)?: ")
        if show_open.lower() not in ['yes', 'no']:
            raise ValueError("Please enter 'yes' or 'no'")
        break
    except ValueError as e:
        print(e)

class_i_want =  input("What class are you eyeing for:")

while True:
    url = 'https://globalsearch.cuny.edu/CFGlobalSearchTool/search.jsp'

    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    # prevent browser from opening. running in the background.
    options.add_argument('--headless')

    driver = webdriver.Chrome(options=options)
    driver.get(url)

    autoPage1()
    autoPage2()
    autoPage3()
    importXLSX()

    # reset data that was scraped for next iteration
    cs_title_list = []
    cs_section_list = []
    cs_DnT_list = []
    cs_availability_list = []
    cs_instructors_list = []
    cs_mode_list = []

    print("Data scraped successfully!")
    driver.quit()
    print("Page closed down." + '\n')
    time.sleep(3)

    print("Scraping again in 30 seconds.")
    print("========================================")
    time.sleep(32)
    

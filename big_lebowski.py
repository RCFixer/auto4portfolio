from tkinter import *
from tkinter import filedialog
import requests
from bs4 import BeautifulSoup
import csv
import threading
import re
from openpyxl import Workbook
import subprocess


global row_number
row_number = 1

root=Tk()
root.title("Автоматизированная проверка 4portfolio.ru")
root.geometry('800x600')

#программа-----------------

def auth(url):
        stages['text']+='1)Авторизация на сайте\n'
        global session
        session = requests.Session()
        url = url+'?login'
        params = {
                'login_username':u'fobas505@gmail.com',
                'login_password':u'132456789',
                'submit':u'Вход',
                'login_submitted':1,
                'login_submitted':1,
                'sesskey':'',
                'pieform_login':''
                }
        r = session.post(url, params)

def write_csv(data):
        with open('students.csv', 'a') as file:
                order = ['name', 'url']
                writer = csv.DictWriter(file, fieldnames=order)
                writer.writerow(data)

def write_xlsx(output):
        wb = Workbook()
        ws = wb.active
        for row in output:
                ws.append(row)
        wb.save("result.xlsx")

def get_html(url):
        r = session.get(url)
        return r.text

def refined(s):
        r = s.split('(')[0]
        r = r.split(' ')
        finall_text = ''
        for i in r:
                if len(i)>0:
                        finall_text += i+' '
        return finall_text.strip()

def students():
        FILENAME = filename
        students = []
        with open(FILENAME, "r", newline="") as file:
                reader = csv.reader(file)
                for row in reader:
                        students.append(row)
        return(students)

def check3(soup,url):
        table = soup.find('div', id = 'bottom-pane').text.strip().split(' ')
        for i in range(table.count('')):
                table.remove('')
        text = soup.find('div', id = 'view-description').text.strip().split(' ')
        for i in range(text.count('')):
                text.remove('')
        summ = len(table)+len(text)
        result = ''
        result = '=HYPERLINK("'+url+'","'+str(summ)+'")'
        return(result)

def check2(url):
        html = get_html(url)
        soup = BeautifulSoup(html, 'lxml')
        page_title = soup.find('div', class_ = 'collection-nav').find('h2').text
        page_title = page_title.split('.')[1].strip()
        check3(soup)
        return(page_title)

def personal_(soup):
        page = soup.find('a', text=re.compile('Личное портфолио'))
        if page !=None:
                url = page.get('href')
                html = get_html(url)
                soup = BeautifulSoup(html, 'lxml')
                return(check3(soup,url))
        else:
                return('-')

def achivments_(soup):
        page = soup.find('a', text=re.compile('Портфолио достижений'))
        if page !=None:
                url = page.get('href')
                html = get_html(url)
                soup = BeautifulSoup(html, 'lxml')
                return(check3(soup,url))
        else:
                return('-')
def documents_(soup):
        page = soup.find('a', text=re.compile('Портфолио документов'))
        if page !=None:
                url = page.get('href')
                html = get_html(url)
                soup = BeautifulSoup(html, 'lxml')
                return(check3(soup,url))
        else:
                return('-')

def reviews_(soup):
        page = soup.find('a', text=re.compile('Портфолио отзывов'))
        if page !=None:
                url = page.get('href')
                html = get_html(url)
                soup = BeautifulSoup(html, 'lxml')
                return(check3(soup,url))
        else:
                return('-')
def check(name, url):
        global row_number
        html=get_html(url)
        soup = BeautifulSoup(html, 'lxml')
        personal = personal_(soup)
        achivments = achivments_(soup)
        documents = documents_(soup)
        reviews = reviews_(soup)
        check_result = [name,personal,achivments,documents,reviews]
        return(check_result)

def processing(students):
        stages['text']+='3)Проверка портфолио нужных нам студентов\n'
        output = [['ФИО','Личное портфолио','Портфолио достижений','Портфолио документов','Портфолио отзывов']]
        with open('students.csv', "r", newline="") as file:
                reader = csv.reader(file)
                for row in reader:
                        for i in students:
                                if row[0].strip()==i[0]:
                                        output.append(check(row[0],row[1]))
                                        students.remove(i)
        stages['text']+='4)Сохранение результатов\n'
        write_xlsx(output)
        for student in students:
                nostudents.insert(1.0, (student[0] + "\n"))
        nostudents.insert(1.0, 'Портфолио не создали:\n')
        stages['text']+='5)Конец'
        subprocess.Popen(['see',"result.xlsx"])
        
                                        

def get_page_data():
        stages['text']+='2)Сбор списка студентов на сайте\n'
        page=0
        f = open('students.csv', 'w')
        f.write("")
        f.close()
        url = 'https://4portfolio.ru/group/members.php?id=750'
        html = get_html(url)
        soup = BeautifulSoup(html,'lxml')
        members = soup.find('div', class_='lead text-small results pull-right').text.split(' ')[0]
        members = int(members)
        while page<members:
                url = 'https://4portfolio.ru/group/members.php?id=750&sortoption=adminfirst&offset='+str(page)+'&setlimit=1&limit=500'
                html = get_html(url)
                soup = BeautifulSoup(html, 'lxml')
                students = soup.find_all('div', class_='list-group-item ')

                for student in students:
                        name = student.find('span', class_='sr-only').text
                        name = refined(name)
                        url = student.find('a', class_='outer-link').get('href')
                        data = {'name': name,
                                'url': url}
                        write_csv(data)
                page+=500
#---------------------------
#оболочка-------------------
def secondSlide(event):
    button1.pack_forget()
    label2.pack_forget()
    text1.pack_forget()
    
    button3.pack(side = BOTTOM, fill=X)
    
    button_open.bind('<Button-1>',openfile)
    button3.bind('<Button-1>',thirdSlide)
    text2.pack(side = TOP, fill=X)
    button_open.pack(expand=1)
    label.pack(expand=1)
    

def thirdSlide(event):
    button3.pack_forget()
    button_open.pack_forget()
    text2.pack_forget()
    label.pack_forget()
    stages.pack(side = TOP)
    stages['text'] = "Этапы проверки(5 пунктов)\n"
    stages['text']+= "---------------------\n"
    nostudents.pack(side = LEFT)
    scroll.pack(side = LEFT, fill=Y)

    
    t = threading.Thread(target=justdoit)
    t.start()

def justdoit():
    url = 'https://4portfolio.ru/'
    auth(url)
    get_page_data()
    processing(students())
    
    
def openfile(event = None):
    global filename
    filename = filedialog.askopenfilename()
    
button3=Button(text='Начать проверку',width=12, height=2)
button_open = Button(text='Обзор', width=12, height=2)

text2=Label(text='При помощи кнопки "Обзор" укажите путь к файлу списка студентов.\n После чего нажмите кнопку "Начать проверку"', bg='#fff')
button1=Button(text='Далее',width=12, height=2)
text1=Label(text='Подготовьте список студентов в виде электронной таблицы с расширением .csv \n(для этого можно использовать офисный пакет LibreOffice). \nФормат имён студентов должен быть "Имя Отчество Фамилия".', bg='#fff')
stages = Label(width=50, height=8)
nostudents = Text(width=98, height=50)
scroll = Scrollbar(command=nostudents.yview)
nostudents.config(yscrollcommand=scroll.set)


button1.bind('<Button-1>',secondSlide)

button1.pack(side = BOTTOM, fill=X)

photo = PhotoImage(file='2slide.png')
label = Label(image=photo)
photo2 = PhotoImage(file='1slide2.png')
label2 = Label(image=photo2)
text1.pack(expand=1)
label2.pack(expand=1)
#--------------------------
root.mainloop()

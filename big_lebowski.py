# -*- coding: utf8 -*-
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup
import csv
import threading
import re
from openpyxl import Workbook
import subprocess

urls_achivments = []
row_number = 1
filename = ''
filename_acc = ''

root=Tk()
root.title("Автоматизированная проверка 4portfolio.ru")
root.geometry('800x600')

#программа-----------------

def write_csv(data):
        with open('students.csv', 'a', newline='', encoding="utf-8") as file:
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
        table = soup.find('div', id = 'bottom-pane').text.strip().replace('\n', '').split(' ')
        for i in range(table.count('')):
                table.remove('')
        text = soup.find('div', id = 'view-description').text.strip().replace('\n', '').split(' ')
        for i in range(text.count('')):
                text.remove('')
        summ = len(table)+len(text)
        result = ''
        result = '=HYPERLINK("'+url+'","'+str(summ)+'")'
        return(result)

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
        global urls_achivments
        page = soup.find('a', text=re.compile('Портфолио достижений'))
        if page !=None:
                url = page.get('href')
                html = get_html(url)
                soup = BeautifulSoup(html, 'lxml')
                achurls = soup.find('nav', class_ = 'custom-dropdown dropdown').find_all('li')
                if len(achurls) == 5:
                        for i in achurls:
                                if i.find('a')==None:
                                        continue
                                urls_achivments.append(i.find('a').get('href'))
                                
                return(check3(soup,url))
        else:
                return('-')

def achivmetns_dlc(url):
        html=get_html(url)
        soup = BeautifulSoup(html, 'lxml')
        return(check3(soup,url))

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
        global urls_achivments
        html=get_html(url)
        soup = BeautifulSoup(html, 'lxml')
        personal = personal_(soup)
        achivments = achivments_(soup)
        documents = documents_(soup)
        reviews = reviews_(soup)
        if len(urls_achivments) == 4:
                achivments_2 = achivmetns_dlc(urls_achivments[0])
                achivments_3 = achivmetns_dlc(urls_achivments[1])
                achivments_4 = achivmetns_dlc(urls_achivments[2])
                achivments_5 = achivmetns_dlc(urls_achivments[3])
        else:
                achivments_2 = '-'
                achivments_3 = '-'
                achivments_4 = '-'
                achivments_5 = '-'
        check_result = [name,personal,achivments,achivments_2,achivments_3,achivments_4,achivments_5,documents,reviews]
        urls_achivments = []
        return(check_result)

def processing(students):
        stages['text']+='2)Проверка портфолио нужных нам студентов\n'
        output = [['ФИО','Личное портфолио','Портфолио достижений 1 страница',
                   '2 страница','3 страница','4 страница','5 страница',
                   'Портфолио документов','Портфолио отзывов']]
        with open('students.csv', "r", newline="", encoding="utf-8") as file:
                reader = csv.reader(file)
                for row in reader:
                        for i in students:
                                if row[0].strip()==i[0]:
                                        output.append(check(row[0],row[1]))
                                        students.remove(i)
        stages['text']+='3)Сохранение результатов\n'
        write_xlsx(output)
        nostudents.insert(1.0, '(Вы можете скопировать этот список при помощи кнопки "Скопировать")\n')
        nostudents.insert(1.0, '--------------------------------------\n')
        for student in students:
                nostudents.insert(1.0, (student[0] + "\n"))
        nostudents.insert(1.0, 'Портфолио не создали:\n')
        stages['text']+='4)Конец'
        button4.pack(expand=1)
        button4.bind('<Button-1>',close)
        button5.pack(expand=1)
        button5.bind('<Button-1>',copy)
        subprocess.Popen(['start',"result.xlsx"], shell=True)
        
                                        

def get_page_data():
        stages['text']+='1)Сбор списка студентов на сайте\n'
        page=0
        f = open('students.csv', 'w', encoding="utf-8")
        f.write("")
        f.close()
        url = 'https://4portfolio.ru/group/members.php?id=750'
        html = get_html(url)
        soup = BeautifulSoup(html,'lxml')
        members = soup.find('div', class_='lead text-small results pull-right').text.split(' ')[0]
        members = int(members)
        while page<members:
                url = '''https://4portfolio.ru/group/members.php?id=750&
                sortoption=adminfirst&offset='''+str(page)+'&setlimit=1&limit=500'
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
        text0.pack_forget()
        text3.pack_forget()
        button1.pack_forget()
        label2.pack_forget()
        text1.pack_forget()
        button3.pack(side = BOTTOM, fill=X)
        button_open.bind('<Button-1>',openfile)
        button3.bind('<Button-1>',thirdSlide)
        text2.pack(expand=1)
        button_open.pack(expand=1)
        label.pack(expand=1)
        attention.pack(expand=1)

def thirdSlide(event):
        global filename
        if len(filename) == 0:
                attention['text'] = "Вы не указали путь к файлу списка студентов!"
                attention['bg'] = "red"
        else:
                attention.pack_forget()
                button3.pack_forget()
                button_open.pack_forget()
                text2.pack_forget()
                label.pack_forget()
                text4.pack(expand=1)
                auth_all.pack(expand=1)
                text5.pack(expand=1)
                pole_login.pack(expand=1)
                text6.pack(expand=1)
                pole_password.pack(expand=1)
                attention_auth.pack(expand=1)
                button6.pack(side = BOTTOM, fill=X)
                button6.bind('<Button-1>',auth)

def fourthSlide():
        text4.pack_forget()
        auth_all.pack_forget()
        text5.pack_forget()
        pole_login.pack_forget()
        text6.pack_forget()
        pole_password.pack_forget()
        attention_auth.pack_forget()
        button6.pack_forget()
        stages.pack(side = TOP)
        stages['text'] = "Этапы проверки(4 пункта)\n"
        stages['text']+= "---------------------\n"
        nostudents.pack(side = LEFT, padx = 3)
        scroll.pack(side = LEFT, fill=Y)
        t = threading.Thread(target=justdoit)
        t.start()

def justdoit():
        get_page_data()
        processing(students())
    
    
def openfile(event = None):
        global filename
        filename = filedialog.askopenfilename()

def openfile_acc(event = None):
        global filename_acc
        filename_acc = filedialog.askopenfilename()

def close(event = None):
        if messagebox.askyesno("Выход", "Вы действительно хотите выйти?"):
                root.destroy()

def auth(event = None):
        login = pole_login.get().strip()
        password = pole_password.get().strip()
        url = 'https://4portfolio.ru/'
        global session
        session = requests.Session()
        url = url+'?login'
        params = {
                'login_username':login,
                'login_password':password,
                'submit':u'Вход',
                'login_submitted':1,
                'login_submitted':1,
                'sesskey':'',
                'pieform_login':''
                }
        r = session.post(url, params)
        html=get_html('https://4portfolio.ru/?login')
        soup = BeautifulSoup(html, 'lxml')
        page = soup.find('span', text=re.compile('Войти на 4portfolio'))
        if page != None:
                attention_auth['text'] = "Вы неверно указали логин или пароль!"
                attention_auth['bg'] = "red"
        else:
                fourthSlide()
        
def copy(event = None):
        nostudents.event_generate("<<SelectAll>>") 
        nostudents.event_generate("<<Copy>>")

button3=Button(text='Далее',width=12, height=2)
button_open = Button(text='Обзор', width=12, height=2)
button4=Button(text='Выход',width=12, height=2)
button5=Button(text='Скопировать',width=12, height=2)
text2=Label(text='При помощи кнопки "Обзор" укажите путь к файлу списка студентов.\n После чего нажмите кнопку "Начать проверку"', bg='#fff')
button1=Button(text='Далее',width=12, height=2)
text1=Label(text='Шаг 1. Подготовьте список студентов в виде электронной таблицы с расширением .csv (например, spisok.csv)\n(для этого можно использовать офисный пакет LibreOffice или MicrosoftOffice). \nФормат имён студентов должен быть "Имя Отчество Фамилия".', bg='#fff')
text3=Label(text='Шаг 2. Сохраните список и нажмите кнопку "Далее"', bg='#fff')
text0=Label(text='Добро пожаловать в мастер\n проверки портфолио',font=("Verdana", 24, 'bold'))
stages = Label(width=50, height=8)
nostudents = Text(width=70, height=50)
scroll = Scrollbar(command=nostudents.yview)
nostudents.config(yscrollcommand=scroll.set)
attention = Label(width = 50, height = 2)
auth_all = Frame()
pole_login = Entry(auth_all, width=30)
pole_password = Entry(auth_all, show="*", width=30)
button6=Button(text='Начать проверку',width=12, height=2)
text4=Label(text='Введите свой логин (e-mail) и пароль от сервиса 4portfolio.ru\n Аккаунт обязан состоять в группе МАГУ' , bg='#fff',font=("Verdana", 15, 'bold'))
text5=Label(auth_all,text='Логин (e-mail)')
text6=Label(auth_all,text='Пароль')
attention_auth = Label(width = 50, height = 2)


button1.bind('<Button-1>',secondSlide)

button1.pack(side = BOTTOM, fill=X)

photo = PhotoImage(file='2slide.png')
label = Label(image=photo)
photo2 = PhotoImage(file='1slide2.png')
label2 = Label(image=photo2)
text0.pack(expand=1)
text1.pack(expand=1)
label2.pack(expand=1)
text3.pack(expand=1)
#--------------------------
root.mainloop()

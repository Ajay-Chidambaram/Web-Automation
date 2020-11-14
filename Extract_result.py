from selenium import webdriver
import openpyxl as op
import pyautogui

wb = op.load_workbook('AC_06.xlsx')
sheet = wb["LARGE SECTION"]
url = webdriver.Chrome()

##User Interface 

##User Input For Subject Data
sub_start = int(pyautogui.prompt(text='Subject Starting Row Number in Excel', title='MTS-A', default=''))
sub_end = int(pyautogui.prompt(text='Subject Ending Row Number in Excel', title='MTS-A', default=''))

subjects = [None] * (sub_end - sub_start + 1)

##Getting Subject Name from Excel File
temp_variable = 0

for sub in range(sub_start , sub_end+1) :
    
    subjects[temp_variable] = sheet['B'+str(sub)].value
    temp_variable = temp_variable + 1


print(subjects)
##Url Provided by User
web_link = pyautogui.prompt(text='Enter the Result URL', title='MTS-A', default='url')
url.get(web_link)

##User Input For Student Data
student_start = int(pyautogui.prompt(text='Student Staritng Row Number in Excel', title='MTS-A', default=''))
student_end = int(pyautogui.prompt(text='Student Ending Row Number in Excel', title='MTS-A', default=''))


url = webdriver.Chrome()
url.get(web_link)


for stu in range(student_start, student_end) :

    registration_field = url.find_element_by_name("regno")
    registration_field.clear()

    student_rollnum = sheet['B'+str(stu)].value
    
    try :

        ##Login using the Registration Number
        registration_field.send_keys(student_rollnum)
        url.find_element_by_xpath("/html/body/div/font[3]/form/input[2]").click()
        
        ##Getting Roll Number from Result Page
        rollno = url.find_element_by_xpath("(//font[@color='#8B008B' and @size='2'])[2]")
        print(rollno.text)
        sheet['B'+str(stu)] = rollno.text

        ##Getting Name from Result Page
        name = url.find_element_by_xpath("//font[@color='#8B008B' and @size='2']")
        print(name.text)
        sheet['C'+str(stu)] = name.text

        ##Grading Column in Excel File (May be User Input)
        entry = ['D' ,'E' ,'F' ,'G' ,'H','I' ,'J' ,'K' ,'L' ,'M']


        ##Getting the Grades of Subjects From Result Page
        tempVariable = 0
        
        for sub in subjects :

            tempVariable = tempVariable + 1

            try :
            
                elec = url.find_element_by_xpath("//tbody/tr[contains(.,'"+sub+"')]/th[2]")
                print(elec.text)
                sheet[entry[tempVariable-1]+str(stu)] = elec.text

            except :

                print("There is no Subject Like " + sub )

    
        ##Number of Subjects Absent
        absent = url.find_elements_by_xpath("//tbody/tr[contains(.,'5') and contains(.,'AB')]")
        print("Present Absent = " +str(len(absent)))
        sheet['S'+str(stu)] = int(len(absent))

        ##Number of Subjects Arrear
        prar = url.find_elements_by_xpath("//tbody/tr[contains(.,'5') and contains(.,'RA')]")
        print("Present Arrears = " +str(len(prar)))
        sheet['T'+str(stu)] = int(len(prar))

        ##Number of Subjects Withdrawn
        withdraw = url.find_elements_by_xpath("//tbody/tr[contains(.,'5') and contains(.,'W')]")
        print("Present Withdraw = " +str(len(withdraw)))
        sheet['U'+str(stu)] = int(len(withdraw))

        ##Number of Subjects Withheld
        withheld = url.find_elements_by_xpath("//tbody/tr[contains(.,'5') and contains(.,'WH')]")
        print("Present Withheld = " +str(len(withheld)))
        sheet['V'+str(stu)] = int(len(withheld))

        #if len(absent)== 0 and len(prar) == 0 and len(withdraw) == 0 and len(withheld) == 0:

        ##Getting GPA from Result Page
        gpa = url.find_element_by_xpath("//table[4]//tr[contains(.,'GPA')]/th[3]")
        print(gpa.text)

        ##Incase of Arrear we can't fill it the Excel GPA column
        try:
            sheet['R'+str(stu)] = float(gpa.text[0:4])
        except:
            sheet['R'+str(stu)] = ""

        url.back()



    except :

        url.back()
            


wb.save('save.xlsx')



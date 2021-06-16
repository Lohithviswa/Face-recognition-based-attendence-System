import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
from tkinter import *
# Function to set focus (cursor)
def focus1(event):
    # set focus on the course_field box
    name_field.focus_set()
def focus2(event):
    # set focus on the course_field box
    branch_field.focus_set()
def focus3(event):
    # set focus on the sem_field box
    year.focus_set()
# Function to set focus
def focus4(event):
    # set focus on the sem_field box
    sem_field.focus_set()
# Function to set focus
def focus5(event):
    # set focus on the form_no_field box
    contact_no_field.focus_set()

def focus6(event):
    # set focus on the email_id_field box
    email_id_field.focus_set()


# Function for clearing the
# contents of text entry boxes
def clear():
    # clear the content of text entry box
    name_field.delete(0, END)
    branch_field.delete(0, END)
    year_field.delete(0, END)
    sem_field.delete(0, END)
    contact_no_field.delete(0, END)
    email_id_field.delete(0, END)


# Function to take data from GUI
# window and write to an excel file
def insert():
    global toaddr
    global year
    global branch
    global section
    # if user not fill any entry
    # then print "empty input"
    if (name_field.get() == "" and
            branch_field.get() == "" and
            year_field.get() == "" and
            sem_field.get() == "" and
            contact_no_field.get() == "" and
            email_id_field.get() == ""):

        print("empty input")

    else:
        name_field.focus_set()
        #print(name_field.get())
        branch=branch_field.get()
        year=year_field.get()
        section=contact_no_field.get()
        toaddr=email_id_field.get()
        # call the clear() function
        clear()

    # Driver code
root = Tk()

    # set the background colour of GUI window
root.configure(background='light green')

    # set the title of GUI window
root.title("registration form")

    # set the configuration of GUI window
root.geometry("500x300")

    # create a Form label
heading = Label(root, text="Form", bg="light green")

    # create a Name label
name = Label(root, text="Name", bg="light green")

    # create a Course label
branch = Label(root, text="Course", bg="light green")

    # create a Semester label
sem = Label(root, text="Semester", bg="light green")

    # create a Form No. lable
year = Label(root, text="Year", bg="light green")

    # create a Contact No. label
contact_no = Label(root, text="Section", bg="light green")

    # create a Email id label
email_id = Label(root, text="Email id", bg="light green")
    # grid method is used for placing
    # the widgets at respective positions
    # in table like structure .
heading.grid(row=0, column=1)
name.grid(row=1, column=0)
branch.grid(row=2, column=0)
sem.grid(row=3, column=0)
year.grid(row=4, column=0)
contact_no.grid(row=5, column=0)
email_id.grid(row=6, column=0)

    # create a text entry box
    # for typing the information
name_field = Entry(root)
branch_field = Entry(root)
sem_field = Entry(root)
year_field = Entry(root)
contact_no_field = Entry(root)
email_id_field = Entry(root)

    # bind method of widget is used for
    # the binding the function with the events

    # whenever the enter key is pressed
    # then call the focus1 function
name_field.bind("<Return>", focus1)

    # whenever the enter key is pressed
    # then call the focus2 function
branch_field.bind("<Return>", focus2)

    # whenever the enter key is pressed
    # then call the focus3 function
sem_field.bind("<Return>", focus3)

    # whenever the enter key is pressed
    # then call the focus4 function
year_field.bind("<Return>", focus4)

    # whenever the enter key is pressed
    # then call the focus5 function
contact_no_field.bind("<Return>", focus5)

    # whenever the enter key is pressed
    # then call the focus6 function
email_id_field.bind("<Return>", focus6)

    # grid method is used for placing
    # the widgets at respective positions
    # in table like structure .
name_field.grid(row=1, column=1, ipadx="100")
branch_field.grid(row=2, column=1, ipadx="100")
sem_field.grid(row=3, column=1, ipadx="100")
year_field.grid(row=4, column=1, ipadx="100")
contact_no_field.grid(row=5, column=1, ipadx="100")
email_id_field.grid(row=6, column=1, ipadx="100")
    # call excel function

    # create a Submit Button and place into the root window
submit = Button(root, text="Submit", fg="Black",bg="Red", command=insert)
submit.grid(row=7, column=1)
root.mainloop()

fromaddr = "lohithviswa@gmail.com"



def sendmail():  
    # instance of MIMEMultipart 
    msg = MIMEMultipart() 
      
    
    msg['From'] = fromaddr 
      
    
    msg['To'] = toaddr 
      
    # storing the subject  
    msg['Subject'] = "Today's attendance"
      
    # string to store the body of the mail 
    body = "Attendence of "+year+"-"+branch+"-"+section+"Please refer below attachment"
      
    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 
      
    # open the file to be sent  
    filename = "Attendance.csv"
    attachment = open("Attendance.csv", "rb") 
      
    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 
      
    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 
      
    # encode into base64 
    encoders.encode_base64(p) 
       
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
      
    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 
    
     filename = "Absentees.csv"
    attachment = open("Absentees.csv", "rb") 
      
    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 
      
    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 
      
    # encode into base64 
    encoders.encode_base64(p) 
       
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
      
    # attach the instance 'p' to instance 'msg' 
    msg.attach(p)
      
    # creates SMTP session 
    s = smtplib.SMTP('smtp.gmail.com', 587) 
      
    # start TLS for security 
    s.starttls() 
      
    # Authentication 
    s.login(fromaddr,"8500636635") 
      
    # Converts the Multipart msg into a string 
    text = msg.as_string() 
      
    # sending the mail 
    s.sendmail(fromaddr, toaddr, text) 
      
    # terminating the session 
    s.quit() 
    
    print("email sent to "+toaddr)




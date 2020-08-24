from tkinter import filedialog
from tkinter import *
from tkinter.filedialog import askopenfilename,asksaveasfile
from tkinter.ttk import Progressbar
from PIL import Image, ImageTk
import docx2txt
import pdfplumber
import os
import time
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import random

if __name__ == "__main__":

    root = Tk()
    
    root.geometry("550x670")
    root.config(bg="white")
    root.title("Applicant Tracking System | Amit")
    root.resizable(0,0)
    headImage = Image.open(f"{os.path.split(os.path.realpath(__file__))[0]}/images/icon.png")
    headPhoto = ImageTk.PhotoImage(headImage)
    root.iconphoto(False, headPhoto)

    tips = ['Optimize your resume with keywords','Avoid images, charts, and other graphics','Avoid Special Characters','Keep Contact Information at the Top','Use a common resume font']
    choice = random.choice(tips)

    def browse_description():
        file = askopenfilename()
        if file == "":  
            file = None
        else:
            descriptionEntry.delete(0,END)
            descriptionEntry.insert(0,file)
        
        descriptionFilePath = descriptionEntry.get()

        def browse_resume():
            file = askopenfilename()
            if file == "":  
                file = None
            else:
                resumeEntry.delete(0,END)
                resumeEntry.insert(0,file)
            
            resumeFilePath = resumeEntry.get()

            def analyse():
                try:
                    description_ext = os.path.splitext(descriptionFilePath)[1].lower()
                    resume_ext = os.path.splitext(resumeFilePath)[1].lower()

                    # Job Description Text Extraction
                    if description_ext=='.docx' or description_ext=='.doc':
                        job_desc = docx2txt.process(descriptionFilePath)
                        
                    elif description_ext=='.pdf':
                        with pdfplumber.open(descriptionFilePath) as pdf:
                            first_page = pdf.pages[0]
                            job_desc = first_page.extract_text()
                            
                    
                    # Resume Text Extraction
                    if resume_ext=='.docx' or resume_ext=='.doc':
                        resume = docx2txt.process(resumeFilePath)
                        
                    elif resume_ext=='.pdf':
                        with pdfplumber.open(resumeFilePath) as pdf:
                            first_page = pdf.pages[0]
                            resume = first_page.extract_text()
 
                    text = [job_desc, resume]

                    cv = CountVectorizer()
                    count_matrix = cv.fit_transform(text)

                    match = cosine_similarity(count_matrix)[0][1]
                    match = match*10
                    result = round(match,1)
                    if result == 10.0:
                        result = 10

                    analyseTextLabel = Label(root,justify=LEFT, compound = RIGHT, padx = 10,text="Analysing",image=iconPhoto,font=('arial',18,'bold'), bg="white", fg="#483D8B")
                    analyseTextLabel.place(x=165,y=420)

                    analysisLabel = Label(root,font=('arial',12,'bold'), bg="white")
                    analysisLabel.place(x=250,y=470)

                    progress = Progressbar(root,orient="horizontal",length=300, mode="determinate")
                    progress.place(x=115, y=500)

                    for i in range(1,100,1):
                        progress['value'] = i
                        root.update_idletasks()
                        analysisLabel.config(text = str(i)+"%")
                        time.sleep(0.03)
                    progress['value'] = 100
                    analyseTextLabel.destroy()
                    analysisLabel.destroy()
                    progress.destroy()
                    analyseButton.destroy()


                    headingLabel = Label(root,text="Your ATS Score",font=('arial',18,'bold'), bg="white", fg="teal")
                    headingLabel.place(x=170,y=430)
                    

                    if result > 7.0:
                        resultLabel = Label(root,text=f"{result} / 10",font=('arial',25,'bold'), bg="white", fg="green")
                        resultLabel.place(x=205,y=480)
                        textLabel = Label(root,text="Your resume is perfectly aligned with the job description and",font=('arial',12,'bold'), bg="white", fg="green")
                        textLabel.place(x=40,y=540)
                        textLabel1 = Label(root,text="is ready for this job.",font=('arial',12,'bold'), bg="white", fg="green")
                        textLabel1.place(x=40,y=570)

                    elif result > 4.0 and result < 6.9:
                        resultLabel = Label(root,text=f"{result} / 10",font=('arial',25,'bold'), bg="white", fg="#FF8C00")
                        resultLabel.place(x=205,y=480)
                        textLabel = Label(root,text="Your resume needs a little bit of improvement before",font=('arial',12,'bold'), bg="white", fg="#FF8C00")
                        textLabel.place(x=60,y=540)
                        textLabel1 = Label(root,text="applying for this job.",font=('arial',12,'bold'), bg="white", fg="#FF8C00")
                        textLabel1.place(x=60,y=570)
                        tipLabel = Label(root,text=f"Tip : {choice}",font=('arial',10,'bold'), bg="white")
                        tipLabel.place(x=60,y=620)
                        
                    else:
                        resultLabel = Label(root,text=f"{result} / 10",font=('arial',25,'bold'), bg="white", fg="red")
                        resultLabel.place(x=205,y=480)
                        textLabel = Label(root,text="Your resume does not align with job description and is not",font=('arial',12,'bold'), bg="white", fg="red")
                        textLabel.place(x=40,y=540)
                        textLabel1 = Label(root,text="suitable for this job.",font=('arial',12,'bold'), bg="white", fg="red")
                        textLabel1.place(x=40,y=570)
                        tipLabel = Label(root,text="Tips : Optimize your resume with keywords.",font=('arial',10,'bold'), bg="white")
                        tipLabel.place(x=40,y=610)
                        tipLabel1 = Label(root,text=": Avoid images, charts, and other graphics.",font=('arial',10,'bold'), bg="white")
                        tipLabel1.place(x=70,y=630)
                except:
                    errorLabel = Label(root,text="Error!!",font=('arial',20,'bold'), bg="white", fg="red")
                    errorLabel.place(x=225, y=430)
                    if resume_ext!='.docx' or resume_ext!='.doc' or resume_ext!='.pdf' or description_ext!='.docx' or description_ext!='.doc' or description_ext!='.pdf' or not resume or not job_desc:
                        errorLabel1 = Label(root,text="Please select the correct files.",font=('arial',16,'bold'), bg="white", fg="red")
                        errorLabel1.place(x=115, y=480)
                    

                def reset():
                    descriptionEntry.delete(0,'end')
                    resumeEntry.delete(0,'end')
                    textLabel.destroy()
                    headingLabel.destroy()
                    textLabel1.destroy()
                    resultLabel.destroy()
                    tipLabel.destroy()
                    tipLabel1.destroy()
                    errorLabel.destroy()
                    errorLabel1.destroy()
                    
                analysisButton = Button(root,text=" Analyse ",font=('arial',12,'bold'),width=20, bg="#4169E1",fg='#FFFFFF', command=analyse)
                analysisButton.place(x=160,y=350)
                resetButton = Button(root,text=" Reset ",font=('arial',12,'bold'),width=20, bg="#9370DB",fg='#FFFFFF', command=reset)
                resetButton.place(x=160,y=350)
            analyseButton = Button(root,text=" Analyse ",font=('arial',12,'bold'),width=20, bg="#4169E1",fg='#FFFFFF', command=analyse)
            analyseButton.place(x=160,y=350)
            

        # Resume Upload
        resumeLabel = Label(root,text="Upload your Resume",font=('sans-serif',12,'bold'), bg="white")
        resumeLabel.place(x=80,y=235)   
        resumeEntry = Entry(root,font=('calibri',12),width=30)
        resumeEntry.place(x=80, y=270)

        openResumeButton = Button(root,text=" Browse ",font=('arial',12,'bold'),width=10, bg="#008080",fg='white',command=browse_resume)
        openResumeButton.place(x=350,y=270)          
        
       


    # ////////////////////////// Front End Part /////////////////////////////////
    
    analysisImage = Image.open(f"{os.path.split(os.path.realpath(__file__))[0]}/images/analysis.png")
    analysisPhoto = ImageTk.PhotoImage(analysisImage)

    image = Image.open(f"{os.path.split(os.path.realpath(__file__))[0]}/images/icon.png")
    iconPhoto = ImageTk.PhotoImage(image)
    appName = Label(root,justify=LEFT, compound = LEFT, padx = 10,text="Applicant Tracking System",image=iconPhoto,font=('cooper black',18,'bold'),
                    bg="white",fg='maroon')
    appName.place(x=35,y=15)

    labelFile = Label(root,text="Check whether your resume is perfect for job or not.",font=('sans-serif',12,'bold'), bg="white")
    labelFile.place(x=70,y=70)


    # Job Description Upload
    descriptionLabel = Label(root,text="Upload Job Description",font=('sans-serif',12,'bold'), bg="white")
    descriptionLabel.place(x=80,y=135)   
    descriptionEntry = Entry(root,font=('calibri',12),width=30)
    descriptionEntry.place(x=80, y=170)
    openDescButton = Button(root,text=" Browse ",font=('arial',12,'bold'),width=10, bg="#008080",fg='white',command=browse_description)
    openDescButton.place(x=350,y=170)

    
 

    root.mainloop()

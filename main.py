import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import PhotoImage
import aspose.slides as slides
import aspose.pydrawing as drawing
from gesture import detection
import cv2
import numpy as np
from PIL import Image, ImageTk
from air_canvas import airscape


def open():
    new =Toplevel()
    new.title('INTELLISLIDES')
    new.geometry('900x600')
    new.configure(background='#FDFEFE')

    browse_button = Button(new, text="Upload PPT", command=browse_file,fg='#FDFEFE', bg='#000000' , padx=30,pady=20)
    browse_button.pack(pady=60)

    convert_button = Button(new, text="Convert to Images", command=convert_ppt_to_img,fg='#FDFEFE', bg='#000000' , padx=30,pady=20)
    convert_button.pack(pady=90)

    def exit_window():
        new.destroy()  # Close the current window
        root.deiconify()  # Bring back the root window

    B2 = Button(new,text="Exit",command=exit_window)
    B2.pack()

def open_manual():
    # Open and display the 'read.png' image
     os.system("start usermanual.pdf")
def air_canvas():
    root.destroy()
    airscape()


def browse_file():
    global file_path
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename()
    #new.deiconify()  # Bring back the new window
    popup()
    
def popup():
    popup_window = messagebox.showinfo("Success", "PPT Uploaded Successfully" )

def convert_ppt_to_img(): 
    ppt_path = file_path
    
    ppt_name = os.path.basename(ppt_path)
    img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'PPT_IMG')
    os.makedirs(img_dir, exist_ok=True)
    
    # Open PowerPoint application
    pres = slides.Presentation(ppt_path)
    thumbnail_size = drawing.Size(1024, 768)
    for index in range(pres.slides.length):
        # Get reference of slide
        slide = pres.slides[index]

        # Save as JPG
        slide.get_thumbnail(thumbnail_size).save(os.path.join(img_dir, f'slide_{index}.jpg'), drawing.imaging.ImageFormat.jpeg)


    print('Images saved successfully in', img_dir)
    # Open navigation window
    quit_window()
    detection('PPT_IMG')
    # detection("PPT")


# Create GUI for selecting PPT file
root = Tk()
root.title('INTELLISLIDES')
root.geometry('900x600')
root.configure(background='#FDFEFE')   # set background color

my_label= Label(root, text='WELCOME TO INTELLISLIDES',fg='#000000',font=("Arial", 25,'bold'))
my_label.pack(pady=20)


B1 = Button(root, text="PPT Controller", command=open,fg='#FDFEFE', bg='#000000' ,font=("Arial", 12,'bold'), padx=30,pady=20)
B1.pack(pady=60)

B2 = Button(root, text="Air Canvas", command=air_canvas,fg='#FDFEFE', bg='#000000' , font=("Arial", 12,'bold'), padx=40,pady=20)
B2.pack(pady=60)



manual_icon = PhotoImage(file="book3.png")   
manual_icon = manual_icon.subsample(4, 4)   # reduce size of image
manual_button = Button(root, text="User Manual", command=open_manual, image=manual_icon, fg='#070607', padx=10, pady=5)
manual_button.place(relx=1.0, rely=1.0, anchor='se')

def quit_window():
    root.destroy()


root.mainloop()

from tkinter import *
from tkinter import filedialog, messagebox
import shutil
import os
from PIL import Image, ImageTk
import subprocess
import cv2
from facenet_pytorch import MTCNN, InceptionResnetV1
import torch
from torch.utils.data import DataLoader
from torchvision import datasets
import pandas as pd
import numpy as np
from collections import Counter
import exifread
import openpyxl
import win32com.client

class GUI:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("Attendance System")
        self.root.iconbitmap("attendance.ico")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        self.root.geometry("%dx%d+%d+%d" % (1400, 800, (screen_width - 1400) / 2, (screen_height - 800) / 2))
        self.root.resizable(False, False)

        self.class_photo_path = ""
        self.already_face_detection = False
        self.already_face_recognition = False

        self.button_face_registration = Button(self.root, text="Face Registration", font=("Arial", 16), command=self.face_registration)
        self.button_face_registration.place(x=0, y=0, width=250, height=100)

        self.button_choose_class_photo = Button(self.root, text="Choose Class Photo", font=("Arial", 16), command=self.choose_class_photo)
        self.button_choose_class_photo.place(x=0, y=100, width=250, height=100)

        self.button_face_detection = Button(self.root, text="Face Detection", command=self.face_detection, font=("Arial", 16))
        self.button_face_detection.place(x=0, y=200, width=250, height=100)

        self.button_face_recognition = Button(self.root, text="Face Recognition", command=self.face_recognition, font=("Arial", 16))
        self.button_face_recognition.place(x=0, y=300, width=250, height=100)

        self.button_clear_list = Button(self.root, text="Clear List", command = self.clear_list, font=("Arial", 16))
        self.button_clear_list.place(x=0, y=400, width=250, height=100)

        self.button_check_attendance = Button(self.root, text="Check Attendance", command=self.check_attendance, font=("Arial", 16))
        self.button_check_attendance.place(x=0, y=500, width=250, height=100)

        self.button_generate_excel = Button(self.root, text="Generate Excel", command=self.generate_excel, font=("Arial", 16))
        self.button_generate_excel.place(x=0, y=600, width=250, height=100)

        self.button_other = Button(self.root, state=DISABLED)
        self.button_other.place(x=0, y=700, width=250, height=100)

        self.container = Frame(self.root, bg="lightgrey")
        self.container.place(x=250, y=0, width=900, height=800)

        self.label_class_photo = Label(self.container)
        self.label_class_photo.config(compound=CENTER, anchor=CENTER)
        self.label_class_photo.place(width=900, height=800)

        self.label_time_info = Label(self.root, text="Time: ", font=("Arial", 16), bg="white", anchor="w")
        self.label_time_info.place(x=1150, y=0, width=250, height=50)

        self.label_pos_info = Label(self.root, text="Position: ", font=("Arial", 16), bg="white", anchor="w")
        self.label_pos_info.place(x=1150, y=50, width=250, height=50)

        self.listbox_attend = Listbox(self.root)
        self.listbox_attend.place(x=1150, y=100, width=250, height=350)

        self.listbox_not_attend = Listbox(self.root)
        self.listbox_not_attend.place(x=1150, y=450, width=250, height=350)

        self.root.mainloop()

    def face_registration(self):
        self.file_paths = ()

        new_window = Toplevel(self.root)
        new_window.title("Face Registration")
        new_window.resizable(False, False)

        label_name = Label(new_window, text="Name:").grid(row=0, column=0, padx=10, pady=10)
        entry_name = Entry(new_window, width=20)
        entry_name.grid(row=0, column=1, padx=10, pady=10)

        label_file = Label(new_window, text="Photos:").grid(row=1, column=0, padx=10, pady=10)
        entry_file = Entry(new_window, state="readonly").grid(row=1, column=1, padx=10, pady=10)
        button_file = Button(new_window, text="Choose file", width=15, command=self.choose_student_photos).grid(row=1, column=2, padx=10, pady=10)

        button_check_dataset = Button(new_window, text="Open Dataset", width=15, command=self.open_student_dataset).grid(row=2, column=0, padx=10, pady=10)
        button_confirm = Button(new_window, text="Confirm", width=15, command=lambda:self.check_student_photos(entry_name.get())).grid(row=2, column=1, padx=10, pady=10)
        button_cancel = Button(new_window, text="Cancel", width=15, command=new_window.destroy).grid(row=2, column=2, columnspan=2, padx=10, pady=10)

    def open_student_dataset(self):
        folder_path = filedialog.askopenfilenames(initialdir="dataset/Student", filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])

    def choose_student_photos(self):
        file_paths = filedialog.askopenfilenames(title="Choose Student Photos", filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
        if file_paths:
            self.file_paths = file_paths

    def check_student_photos(self, name):
        if name == "":
            messagebox.showerror("Error", "Please enter the name of the student!")
        else:
            flag = False
            new_folder_path = "dataset/Student/" + name
            if not os.path.exists(new_folder_path):
                os.makedirs(new_folder_path, exist_ok=True)
                if len(self.file_paths) < 3:
                    messagebox.showerror("Error", "Please choose at least 3 student photos!")
                else:
                    flag = True
            else:
                    flag = True
            if flag:
                count = 0
                
                for file_name in os.listdir("dataset/Student/" + name):
                    file_name = file_name[:-4]
                    count = max(count, int(file_name))
                        
                for path in self.file_paths:
                    count += 1
                    image_name = os.path.basename(path)
                    image_name = str(count) + image_name[-4:]
                    destination_path = new_folder_path + "/" + image_name

                    shutil.copyfile(path, destination_path)

                messagebox.showinfo("Information", "Successfully uploaded student photos!")
                self.file_paths = ()

    def load_class_photo(self, filepath):
        image = Image.open(filepath)

        # 图片太大就缩小
        if image.size[0] > 900 and image.size[1] > 800:
            if image.size[0] >= image.size[1]:
                image = image.resize((900, round(image.size[1] * 900 / image.size[0])), Image.LANCZOS)
            else:
                image = image.resize((round(image.size[0] * 800 / image.size[1]), 800), Image.LANCZOS)
        elif image.size[0] > 900:
            image = image.resize((900, round(image.size[1] * 900 / image.size[0])), Image.LANCZOS)
        elif image.size[1] > 800:
            image = image.resize((round(image.size[0] * 800 / image.size[1]), 800), Image.LANCZOS)

        # 图片太小就放大
        if image.size[0] < 900 and image.size[1] < 800:
            if image.size[0] >= image.size[1]:
                image = image.resize((900, round(image.size[1] * 900 / image.size[0])), Image.LANCZOS)
            else:
                image = image.resize((round(image.size[0] * 800 / image.size[1]), 800), Image.LANCZOS)
        
        image = ImageTk.PhotoImage(image)
        self.label_class_photo.config(image=image)
        self.label_class_photo.image = image
        
    def choose_class_photo(self):
        class_photo_path = filedialog.askopenfilename(title="Choose Student Photos", filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
        if class_photo_path:
            self.class_photo_path = class_photo_path
            self.already_face_detection = False
            self.already_face_recognition = False
            self.load_class_photo(self.class_photo_path)

            for file_name in os.listdir("dataset/Class"):
                file_path = "dataset/Class/" + file_name
                if os.path.isfile(file_path):
                    os.remove(file_path)
        
            shutil.copyfile(self.class_photo_path, "dataset/Class/" + os.path.basename(self.class_photo_path))

            with open(self.class_photo_path, 'rb') as f:
                img = Image.open(f)
                exif_data = img._getexif()

                tags = exifread.process_file(f)

                location_info = None
                time_info = None

                if exif_data:
                    gps_info = exif_data.get(34853, None)  # 34853 对应 GPS 数据
                    if gps_info:
                        latitude = gps_info[2][0] + gps_info[2][1] / 60.0 + gps_info[2][2] / 3600.0
                        longitude = gps_info[4][0] + gps_info[4][1] / 60.0 + gps_info[4][2] / 3600.0
                        location_info = {'latitude': latitude, 'longitude': longitude}
                        temp = "latitude: {}\nlongitude: {}".format(round(location_info['latitude']), round(location_info['longitude']))
                        self.label_pos_info.config(text=temp)

                if 'EXIF DateTimeOriginal' in tags:
                    time_info = tags['EXIF DateTimeOriginal']
                    self.label_time_info.config(text=time_info)
                

    def face_detection(self):
        if self.already_face_detection:
            messagebox.showwarning("Warning", "You already have done face detection for this image!")
        else:
            if self.class_photo_path:
                shutil.rmtree("yolov5/runs/detect")
                
                messagebox.showinfo("Information", "Start face detection! Please wait a few seconds")
                command = "python yolov5/detect.py --weights yolov5/runs/train/exp/weights/best.pt --conf 0.3 --source dataset/Class --save-txt"

                try:
                    os.system(command)
                    
                    coordinates = []
                    with open('yolov5/runs/detect/exp/labels/' + os.path.splitext(os.path.basename(self.class_photo_path))[0] + ".txt", 'r') as file:
                        lines = file.readlines()
                        for line in lines:
                            data = line[:-1].split(" ")[1:]
                            data = [float(t) for t in data]
                            if data[2] >= 0.05 and data[3] >= 0.05:
                                x_min = data[0] - data[2] / 2
                                x_max = data[0] + data[2] / 2
                                y_min = data[1] - data[3] / 2
                                y_max = data[1] + data[3] / 2
                                coordinates.append([x_min, y_min, x_max, y_max])

                    img = cv2.imread(self.class_photo_path)
                    image_height, image_width = img.shape[:2]

                    for file_name in os.listdir("dataset/Student/Crops"):
                        file_path = "dataset/Student/Crops/" + file_name

                        if os.path.isfile(file_path):
                            os.remove(file_path)

                    count = 0
                    for coord in coordinates:
                        x_min = round(coord[0] * image_width)
                        y_min = round(coord[1] * image_height)
                        x_max = round(coord[2] * image_width)
                        y_max = round(coord[3] * image_height)
                        cropped_img = img[y_min:y_max, x_min:x_max]
                        count += 1
                        cv2.imwrite("dataset/Student/Crops/" + str(count) + ".jpg", cropped_img)

                    for coord in coordinates:
                        x_min = round(coord[0] * image_width)
                        y_min = round(coord[1] * image_height)
                        x_max = round(coord[2] * image_width)
                        y_max = round(coord[3] * image_height)
                        cv2.rectangle(img, (x_min, y_min), (x_max, y_max), (0, 255, 0), 2)

                    cv2.imwrite("dataset/detected.jpg", img)
                    self.load_class_photo("dataset/detected.jpg")

                    self.already_face_detection = True
                    messagebox.showinfo("Information", "Face detection succeeded!")
                except subprocess.CalledProcessError:
                    messagebox.showerror("Error", "Face detection failed!")
            else:
                messagebox.showerror("Error", "Please choose one class photo first!")

    def face_recognition(self):
        flag = True
        for root, _, files in os.walk("dataset/Student"):
            if root == "dataset/Student" or root == "dataset/Student\Crops":
                continue
            if len(files) < 3:
                messagebox.showerror("Error", "Please make sure each student has at least 3 photos in the dataset!")
                flag = False
                break

        if (flag):
            if self.already_face_detection:
                if not self.already_face_recognition:
                    messagebox.showinfo("Information", "Start face recognition! Please wait a few seconds")
                    
                    device = torch.device('cuda:0' if torch.cuda.is_available() else 'cpu')
                    mtcnn = MTCNN(image_size=160, margin=0, min_face_size=20, thresholds=[0.3, 0.3, 0.3], factor=0.709, post_process=True,
                                device=device)
                    resnet = InceptionResnetV1(pretrained='vggface2').eval().to(device)
                    def collate_fn(x):
                        return x[0]
                    
                    dataset = datasets.ImageFolder('dataset/Student')
                    dataset.idx_to_class = {i:c for c, i in dataset.class_to_idx.items()}
                    loader = DataLoader(dataset, collate_fn=collate_fn, num_workers=0)

                    aligned = []
                    names = []
                    for x, y in loader:
                        x_aligned = mtcnn(x, return_prob=False)
                        if x_aligned is not None:
                            aligned.append(x_aligned)
                            names.append(dataset.idx_to_class[y])
                    
                    aligned = torch.stack(aligned).to(device)
                    embeddings = resnet(aligned).detach().cpu()
                    
                    dists = [[(e1 - e2).norm().item() for e2 in embeddings] for e1 in embeddings]
                    
                    result = pd.DataFrame(dists, columns=names, index=names).drop("Crops", axis=1).loc["Crops"]
                    
                    def sort_row(row):
                        return sorted(row)
                    
                    df_sorted = result.apply(sort_row, axis=1)

                    labels = []
                    for i in range(df_sorted.shape[0]):
                        
                        if df_sorted.iloc[i][0] >= 0.9:
                            labels.append("others")
                        else:
                            index = []
                            index.append(result.iloc[i].index[result.iloc[i].eq(df_sorted.iloc[i][0])].to_list()[0])
                            index.append(result.iloc[i].index[result.iloc[i].eq(df_sorted.iloc[i][1])].to_list()[0])
                            index.append(result.iloc[i].index[result.iloc[i].eq(df_sorted.iloc[i][2])].to_list()[0])
                            counter = Counter(index)
                            duplicates = [item for item, count in counter.items() if count > 1]
                            if len(duplicates) == 0:
                                labels.append("others")
                            else:
                                labels.append(sorted(duplicates, reverse=True)[0])


                    img = cv2.imread("dataset/detected.jpg")
                    image_height, image_width = img.shape[:2]

                    coordinates = []
                    with open('yolov5/runs/detect/exp/labels/' + os.path.splitext(os.path.basename(self.class_photo_path))[0] + ".txt", 'r') as file:
                        lines = file.readlines()
                        for line in lines:
                            data = line[:-1].split(" ")[1:]
                            data = [float(t) for t in data]
                            if data[2] >= 0.05 and data[3] >= 0.05:
                                x_min = data[0] - data[2] / 2
                                x_max = data[0] + data[2] / 2
                                y_min = data[1] - data[3] / 2
                                y_max = data[1] + data[3] / 2
                                coordinates.append([x_min, y_min, x_max, y_max])

                    for i, coord in enumerate(coordinates):
                        x_min = round(coord[0] * image_width)
                        y_min = round(coord[1] * image_height)
                        x_max = round(coord[2] * image_width)
                        y_max = round(coord[3] * image_height)

                        cv2.putText(img, labels[i], (x_max - 80, y_min + 10), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)

                    cv2.imwrite("dataset/labeled.jpg", img)
                    self.load_class_photo("dataset/labeled.jpg")
                    
                    student_already_attend = self.listbox_attend.get(0, END)
                    for label in labels:
                        if label != "others": 
                            if label not in student_already_attend:
                                self.listbox_attend.insert(END, label)



                    self.already_face_recognition = True
                    messagebox.showinfo("Information", "Face recognition succeeded!")
                else:
                    messagebox.showwarning("Warning", "You already have done face recognition for this image!")       
            else:
                messagebox.showerror("Error", "Please choose one class photo and do face detection first!")

    def clear_list(self):
        self.listbox_attend.delete(0, END)
        self.listbox_not_attend.delete(0, END)

    def check_attendance(self):
        self.listbox_not_attend.delete(0, END)
        student_name_list = [dirname for dirname in os.listdir("dataset/Student") if os.path.isdir("dataset/Student/" + dirname) and dirname != "Crops"]
        student_already_attend = self.listbox_attend.get(0, END)
        for i, name in enumerate(student_name_list):
            if not name in student_already_attend:
                self.listbox_not_attend.insert(END, name)
                self.listbox_not_attend.itemconfig(END, {'fg': "red"})

    def generate_excel(self):
        student_attend = self.listbox_attend.get(0, END)
        student_not_attend = self.listbox_not_attend.get(0, END)
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        sheet["A1"] = self.label_pos_info.cget("text")
        sheet["A2"] = self.label_time_info.cget("text")
        
        sheet["A4"] = "Attend: "
        sheet.append(student_attend)
        sheet["A7"] = "Not Attend: "
        sheet.append(student_not_attend)
        
        workbook.save('output.xlsx')
        messagebox.showinfo("Information", "Output excel succeeded!")

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True
            workbook = excel.Workbooks.Open(os.path.abspath('output.xlsx'))
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    GUI()
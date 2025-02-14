import customtkinter
from PIL import Image
import os


class ToplevelWindow0(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("600x400")
        self.title('第一周校历示例')

        self.image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "Images")
        self.label = customtkinter.CTkLabel(self, text="第一周校历示例")
        self.label.pack(padx=20, pady=20)
        self.page1_image = customtkinter.CTkImage(Image.open(os.path.join(self.image_path, "frist week.jpg")), size=(400, 230))
        self.page1_image_lable = customtkinter.CTkLabel(self, text="", image=self.page1_image)
        self.page1_image_lable.pack()
class ToplevelWindow1(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("700x600")
        self.title('课程考勤示例')

        self.image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "Images")
        self.label = customtkinter.CTkLabel(self, text="单科考勤示例")
        self.label.pack(padx=20, pady=20)
        self.page1_image = customtkinter.CTkImage(Image.open(os.path.join(self.image_path, "course attendance.jpg")), size=(650, 540))
        self.page1_image_lable = customtkinter.CTkLabel(self, text="", image=self.page1_image)
        self.page1_image_lable.pack()

class ToplevelWindow2(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("700x400")
        self.title('成绩导入模板示例')

        self.image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "Images")
        self.label = customtkinter.CTkLabel(self, text="成绩导入模板示例")
        self.label.pack(padx=20, pady=20)
        self.page1_image = customtkinter.CTkImage(Image.open(os.path.join(self.image_path, "score modle.jpg")), size=(600, 320))
        self.page1_image_lable = customtkinter.CTkLabel(self, text="", image=self.page1_image)
        self.page1_image_lable.pack()
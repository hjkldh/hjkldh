import customtkinter
from CTkDatePicker import *
from pack import Message,page_process,topnotice
import os
from PIL import Image
#import pywinstyles

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("期末考核工具")
        #self.geometry(f"{1000}x{500}")

        self.width = int(self.winfo_screenwidth()/2.5)
        self.height = int(self.winfo_screenheight()/2)
        self.geometry(f"{self.width}x{self.height}")
        self.minsize(1200,700)
        # configure grid layout (4x4)
        #self.grid_columnconfigure(1, weight=1)
        #self.grid_columnconfigure((2, 3), weight=0)
        #self.grid_rowconfigure((0, 1, 2), weight=1)
        self.image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "pack/Images")
        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=10, height=20, corner_radius=0)
        self.sidebar_frame.pack(expand=False, fill="x", padx=10, pady=10)
        #self.sidebar_frame.grid(row=0, column=0, rowspan=1, columnspan=1, sticky="nsew")
        #self.sidebar_frame.grid_rowconfigure(2, weight=1)
#page1
        self.page1_button = customtkinter.CTkButton(self.sidebar_frame, corner_radius=0, height=40, border_spacing=10, text="课程考勤",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), anchor="center", command=self.page1_button_event)
        self.page1_button.grid(row=1, column=0, sticky="ew")
        self.frame_page1 = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        #self.frame_page1.pack(expand=True, fill="both", padx=10, pady=10)
        self.label_1 = customtkinter.CTkLabel(self.frame_page1, text='1.请选择第一周开始日期：')
        self.label_1.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.date_picker = CTkDatePicker(self.frame_page1)
        self.date_picker.grid(row=0, column=1, padx=20, pady=20, sticky="w")
        self.date_picker.set_allow_manual_input(False)
        self.button_notice0 = customtkinter.CTkButton(self.frame_page1, width=10, height=20, corner_radius=3, text='示例',fg_color='red',hover_color="green",command=self.page1_top0)
        self.button_notice0.grid(row=0, column=1, padx=230, pady=10, sticky="w")
        self.page1_toplevel0 = None

        self.label_2 = customtkinter.CTkLabel(self.frame_page1, text='2.请选择系统导出考勤excel文件：')
        self.label_2.grid(row=1, column=0, padx=10, pady=10)

        self.page1_entry1 = customtkinter.CTkEntry(master=self.frame_page1,placeholder_text="请选择文件或者粘贴文件路径",state="normal", width=500, height=20, corner_radius=3)
        self.page1_entry1.grid(row=1, column=1, padx=(20, 10), pady=(20, 20), sticky="w")

        self.button_1 = customtkinter.CTkButton(self.frame_page1, width=10, height=20, corner_radius=3, text='选择文件',
                                                command=self.page1_file1)
        self.button_1.grid(row=1, column=2, padx=10, pady=10)
        self.button_notice1 = customtkinter.CTkButton(self.frame_page1, width=10, height=20, corner_radius=3, text='示例',fg_color='red',hover_color="green",command=self.page1_top1)
        self.button_notice1.grid(row=1, column=3, padx=10, pady=10)
        self.page1_toplevel1 = None

        self.label_3 = customtkinter.CTkLabel(self.frame_page1, text='3.选择班级名册：')
        self.label_3.grid(row=2, column=0, padx=10, pady=10, sticky="w")

        self.page1_entry2 = customtkinter.CTkEntry(master=self.frame_page1,placeholder_text="请选择文件或者粘贴文件路径",state="normal", width=500, height=20, corner_radius=3)
        self.page1_entry2.grid(row=2, column=1, padx=(20, 10), pady=(20, 20), sticky="w")
        self.button_2 = customtkinter.CTkButton(self.frame_page1, width=10, height=20, corner_radius=3,
                                                text='选择文件', command=self.page1_file2)
        self.button_2.grid(row=2, column=2, padx=10, pady=10)
        self.label_4 = customtkinter.CTkLabel(self.frame_page1, text='4.选择输出目录：')
        self.label_4.grid(row=3, column=0, padx=10, pady=10, sticky="w")

        self.page1_entry3 = customtkinter.CTkEntry(master=self.frame_page1,placeholder_text="请选择或粘贴输出目录，不选择则输出到系统名册的同级目录",state="normal", width=500, height=20, corner_radius=3)
        self.page1_entry3.grid(row=3, column=1, padx=(20, 10), pady=(20, 20), sticky="w")
        self.button_3 = customtkinter.CTkButton(self.frame_page1, width=10, height=20, corner_radius=3,
                                                text='选择目录', command=self.page1_dir1)
        self.button_3.grid(row=3, column=2, padx=10, pady=10)
        self.label_5 = customtkinter.CTkLabel(self.frame_page1, text='5.处理并输出：')
        self.label_5.grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.button_4 = customtkinter.CTkButton(self.frame_page1, width=10, height=20, corner_radius=3,
                                                text='处理并输出', command=self.page1_process)
        self.button_4.grid(row=4, column=1, padx=(20, 10), pady=(20, 20), sticky="w")
#page2
        self.page2_button = customtkinter.CTkButton(self.sidebar_frame, corner_radius=0, height=40, border_spacing=10, text="考核记录单",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), anchor="center", command=self.page2_button_event)
        self.page2_button.grid(row=1, column=1, sticky="ew")
        self.frame_page2 = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        #self.frame_page2.pack(expand=True, fill="both", padx=10, pady=10)
        self.page2_label1 = customtkinter.CTkLabel(self.frame_page2, text='考勤占比(%)：')
        self.page2_label1.grid(row=0, column=0, sticky="e")
        self.page2_entry1 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认20",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry1.grid(row=0, column=1, sticky="w")
        self.page2_entry1.insert(0,'20')
        self.page2_label2 = customtkinter.CTkLabel(self.frame_page2, text='课堂参与占比(%)：')
        self.page2_label2.grid(row=0, column=2, sticky="e")
        self.page2_entry2 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认20",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry2.grid(row=0, column=3, sticky="w")
        self.page2_entry2.insert(0,'20')
        self.page2_label3 = customtkinter.CTkLabel(self.frame_page2, text='作业占比(%)：')
        self.page2_label3.grid(row=0, column=4, sticky="e")
        self.page2_entry3 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认0",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry3.grid(row=0, column=5, sticky="w")
        self.page2_entry3.insert(0,'0')
        self.page2_label4 = customtkinter.CTkLabel(self.frame_page2, text='期末测试占比(%)：')
        self.page2_label4.grid(row=0, column=6, sticky="e")
        self.page2_entry4 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认60",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry4.grid(row=0, column=7, sticky="w")
        self.page2_entry4.insert(0,'60')
        self.page2_label5 = customtkinter.CTkLabel(self.frame_page2, text='公假一次扣：')
        self.page2_label5.grid(row=1, column=0, sticky="e")
        self.page2_entry5 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认 0",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry5.grid(row=1, column=1, sticky="w")
        self.page2_entry5.insert(0,'0')
        self.page2_label6 = customtkinter.CTkLabel(self.frame_page2, text='病假一次扣：')
        self.page2_label6.grid(row=1, column=2, sticky="e")
        self.page2_entry6 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认 0",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry6.grid(row=1, column=3, sticky="w")
        self.page2_entry6.insert(0,'0')
        self.page2_label7 = customtkinter.CTkLabel(self.frame_page2, text='事假一次扣：')
        self.page2_label7.grid(row=1, column=4, sticky="e")
        self.page2_entry7 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认 0",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry7.grid(row=1, column=5, sticky="w")
        self.page2_entry7.insert(0,'1')
        self.page2_label8 = customtkinter.CTkLabel(self.frame_page2, text='迟到一次扣：')
        self.page2_label8.grid(row=1, column=6, sticky="e")
        self.page2_entry8 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认 0",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry8.grid(row=1, column=7, sticky="w")
        self.page2_entry8.insert(0,'0.5')
        self.page2_label9 = customtkinter.CTkLabel(self.frame_page2, text='早退一次扣：')
        self.page2_label9.grid(row=1, column=8, sticky="e")
        self.page2_entry9 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认 0",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry9.grid(row=1, column=9, sticky="w")
        self.page2_entry9.insert(0,'1')
        self.page2_label10 = customtkinter.CTkLabel(self.frame_page2, text='旷课一次扣：')
        self.page2_label10.grid(row=1, column=10, sticky="e")
        self.page2_entry10 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="默认 0",state="normal", width=50, height=20, corner_radius=3)
        self.page2_entry10.grid(row=1, column=11, sticky="w")
        self.page2_entry10.insert(0,'4')
        self.page2_label11 = customtkinter.CTkLabel(self.frame_page2, text='1、请选择课程考勤名单：')
        self.page2_label11.grid(row=2, column=0, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_entry11 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="选择18周考勤文件，即“课程考勤”生成为文件",state="normal", width=500, height=20, corner_radius=3)
        self.page2_entry11.grid(row=2, column=1, columnspan=6, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_button1 = customtkinter.CTkButton(self.frame_page2, width=10, height=20, corner_radius=3, text='选择文件', command=self.page2_file1)
        self.page2_button1.grid(row=2, column=7, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_label12 = customtkinter.CTkLabel(self.frame_page2, text='2、请选择班级名册：')
        self.page2_label12.grid(row=3, column=0, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_entry12 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="选择班级名册文件",state="normal", width=500, height=20, corner_radius=3)
        self.page2_entry12.grid(row=3, column=1, columnspan=6, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_button2 = customtkinter.CTkButton(self.frame_page2, width=10, height=20, corner_radius=3, text='选择文件', command=self.page2_file2)
        self.page2_button2.grid(row=3, column=7, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_label13 = customtkinter.CTkLabel(self.frame_page2, text='3、请选择课程作业成绩：')
        self.page2_label13.grid(row=4, column=0, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_entry13 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="不选择则以设置的分数打分",state="normal", width=500, height=20, corner_radius=3)
        self.page2_entry13.grid(row=4, column=1, columnspan=6, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_button3 = customtkinter.CTkButton(self.frame_page2, width=10, height=20, corner_radius=3, text='选择文件', command=self.page2_file3)
        self.page2_button3.grid(row=4, column=7, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_label14 = customtkinter.CTkLabel(self.frame_page2, text='4、请选择课程试卷成绩：')
        self.page2_label14.grid(row=5, column=0, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_entry14 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="请选择试卷成绩文件，表头应该包含“学号”和“成绩”两列",state="normal", width=500, height=20, corner_radius=3)
        self.page2_entry14.grid(row=5, column=1, columnspan=6, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_button4 = customtkinter.CTkButton(self.frame_page2, width=10, height=20, corner_radius=3, text='选择文件', command=self.page2_file4)
        self.page2_button4.grid(row=5, column=7, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_label15 = customtkinter.CTkLabel(self.frame_page2, text='5、选择输出路径：')
        self.page2_label15.grid(row=6, column=0, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_entry15 = customtkinter.CTkEntry(master=self.frame_page2,placeholder_text="请选择或者粘贴输出路径，不选择则输出到考勤同名目录",state="normal", width=500, height=20, corner_radius=3)
        self.page2_entry15.grid(row=6, column=1, columnspan=6, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_button5 = customtkinter.CTkButton(self.frame_page2, width=10, height=20, corner_radius=3, text='选择目录', command=self.page2_dir1)
        self.page2_button5.grid(row=6, column=7, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_label16 = customtkinter.CTkLabel(self.frame_page2, text='6、生成记录成绩单：')
        self.page2_label16.grid(row=7, column=0, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page2_button5 = customtkinter.CTkButton(self.frame_page2, width=10, height=20, corner_radius=3, text='开始处理', command=self.page2_process)
        self.page2_button5.grid(row=7, column=1, padx=(20, 10), pady=(20, 20), sticky="w")
#page3
        self.page3_button = customtkinter.CTkButton(self.sidebar_frame, corner_radius=0, height=40, border_spacing=10, text="上传模版",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), anchor="center", command=self.page3_button_event)
        self.page3_button.grid(row=1, column=2, sticky="ew")
        self.frame_page3 = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        #self.frame_page3.pack(expand=True, fill="both", padx=10, pady=10)
        self.page3_label0 = customtkinter.CTkLabel(self.frame_page3, text='1、请设置成绩占比，务必与系统中一致：')
        self.page3_label0.grid(row=0, column=0, sticky="e")
        self.page3_label1 = customtkinter.CTkLabel(self.frame_page3, text='期末占比(%)：')
        self.page3_label1.grid(row=1, column=0, sticky="e")
        self.page3_entry1 = customtkinter.CTkEntry(master=self.frame_page3,placeholder_text="默认60",state="normal", width=50, height=20, corner_radius=3)
        self.page3_entry1.grid(row=1, column=1, sticky="w")
        self.page3_entry1.insert(0,'60')
        self.page3_label2 = customtkinter.CTkLabel(self.frame_page3, text='期中占比(%)：')
        self.page3_label2.grid(row=1, column=2, sticky="e")
        self.page3_entry2 = customtkinter.CTkEntry(master=self.frame_page3,placeholder_text="默认0",state="normal", width=50, height=20, corner_radius=3)
        self.page3_entry2.grid(row=1, column=3, sticky="w")
        self.page3_entry2.insert(0,'0')
        self.page3_label3 = customtkinter.CTkLabel(self.frame_page3, text='平时成绩占比(%)：')
        self.page3_label3.grid(row=1, column=4, sticky="e")
        self.page3_entry3 = customtkinter.CTkEntry(master=self.frame_page3,placeholder_text="默认20",state="normal", width=50, height=20, corner_radius=3)
        self.page3_entry3.grid(row=1, column=5, sticky="w")
        self.page3_entry3.insert(0,'20')
        self.page3_label4 = customtkinter.CTkLabel(self.frame_page3, text='考勤占比(%)：')
        self.page3_label4.grid(row=1, column=6, sticky="e")
        self.page3_entry4 = customtkinter.CTkEntry(master=self.frame_page3,placeholder_text="默认20",state="normal", width=50, height=20, corner_radius=3)
        self.page3_entry4.grid(row=1, column=7, sticky="w")
        self.page3_entry4.insert(0,'20')
        self.page3_label5 = customtkinter.CTkLabel(self.frame_page3, text='作业占比(%)：')
        self.page3_label5.grid(row=1, column=8, sticky="e")
        self.page3_entry5 = customtkinter.CTkEntry(master=self.frame_page3,placeholder_text="默认0",state="normal", width=50, height=20, corner_radius=3)
        self.page3_entry5.grid(row=1, column=9, sticky="w")
        self.page3_entry5.insert(0,'0')
        self.page3_label6 = customtkinter.CTkLabel(self.frame_page3, text='2、请选择系统中导出的模板文件：')
        self.page3_label6.grid(row=2, column=0, sticky="w")
        self.page3_entry6 = customtkinter.CTkEntry(master=self.frame_page3,placeholder_text="请选择模版文件",state="normal", width=500, height=20, corner_radius=3)
        self.page3_entry6.grid(row=2, column=1,columnspan=11, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page3_button1 = customtkinter.CTkButton(self.frame_page3, width=10, height=20, corner_radius=3, text='选择文件', command=self.page3_file1)
        self.page3_button1.grid(row=2, column=8, padx=(20, 10), pady=(20, 20), sticky="w")
        self.button_notice0 = customtkinter.CTkButton(self.frame_page3, width=10, height=20, corner_radius=3, text='示例',fg_color='red',hover_color="green",command=self.page3_top0)
        self.button_notice0.grid(row=2, column=9, padx=10, pady=10, sticky="w")
        self.page3_toplevel0 = None
        self.page3_label7 = customtkinter.CTkLabel(self.frame_page3, text='3、请选择考核记录单成绩表：')
        self.page3_label7.grid(row=3, column=0, sticky="w")
        self.page3_entry7 = customtkinter.CTkEntry(master=self.frame_page3,placeholder_text="请选择成绩记录表，即“考核记录单”中导出文件",state="normal", width=500, height=20, corner_radius=3)
        self.page3_entry7.grid(row=3, column=1,columnspan=11, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page3_button2 = customtkinter.CTkButton(self.frame_page3, width=10, height=20, corner_radius=3, text='选择文件', command=self.page3_file2)
        self.page3_button2.grid(row=3, column=8, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page3_label8 = customtkinter.CTkLabel(self.frame_page3, text='4、请选择期中成绩：')
        self.page3_label8.grid(row=4, column=0, sticky="w")
        self.page3_entry8 = customtkinter.CTkEntry(master=self.frame_page3,placeholder_text="没有期中成绩则不选择",state="normal", width=500, height=20, corner_radius=3)
        self.page3_entry8.grid(row=4, column=1,columnspan=11, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page3_button3 = customtkinter.CTkButton(self.frame_page3, width=10, height=20, corner_radius=3, text='选择目录', command=self.page3_file3)
        self.page3_button3.grid(row=4, column=8, padx=(20, 10), pady=(20, 20), sticky="w")

        self.page3_label9 = customtkinter.CTkLabel(self.frame_page3, text='5、请选择输出目录：')
        self.page3_label9.grid(row=5, column=0, sticky="w")
        self.page3_entry9 = customtkinter.CTkEntry(master=self.frame_page3,placeholder_text="不选择则输出到模板同目录",state="normal", width=500, height=20, corner_radius=3)
        self.page3_entry9.grid(row=5, column=1,columnspan=11, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page3_button4 = customtkinter.CTkButton(self.frame_page3, width=10, height=20, corner_radius=3, text='选择目录', command=self.page3_dir1)
        self.page3_button4.grid(row=5, column=8, padx=(20, 10), pady=(20, 20), sticky="w")

        self.page3_label10 = customtkinter.CTkLabel(self.frame_page3, text='6、生成上传成绩表：')
        self.page3_label10.grid(row=6, column=0, sticky="w")
        self.page3_button5 = customtkinter.CTkButton(self.frame_page3, width=10, height=20, corner_radius=3, text='开始处理', command=self.page3_process)
        self.page3_button5.grid(row=6, column=1, padx=(20, 10), pady=(20, 20), sticky="w")
#page4
        self.page4_button = customtkinter.CTkButton(self.sidebar_frame, corner_radius=0, height=40, border_spacing=10, text="成绩分析表",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), anchor="center", command=self.page4_button_event)
        self.page4_button.grid(row=1, column=3, sticky="ew")
        self.frame_page4 = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        #self.frame_page4.pack(expand=True, fill="both", padx=10, pady=10)
        self.page4_lable = customtkinter.CTkLabel(self.frame_page4, text='请本地部署 deepseek 或者 kimi，让其按分析表分析即可!')
        self.page4_lable.grid(row=0, column=0)
        self.page4_lable1 = customtkinter.CTkTextbox(self.frame_page4, width=500, corner_radius=0)
        self.page4_lable1.grid(row=1, column=0)
        self.page4_lable1.insert("0.0",'deepseek：https://chat.deepseek.com/')
        self.page4_lable1.configure('True')

        self.page4_lable2 = customtkinter.CTkTextbox(self.frame_page4, width=500, corner_radius=0)
        self.page4_lable2.grid(row=2, column=0)
        self.page4_lable2.insert("0.0", 'kimi：https://kimi.moonshot.cn/')
        self.page4_lable2.configure('True')
#page5
        self.page5_button = customtkinter.CTkButton(self.sidebar_frame, corner_radius=0, height=40, border_spacing=10, text="班主任考勤表",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), anchor="center", command=self.page5_button_event)
        self.page5_button.grid(row=1, column=4, sticky="ew")
        self.frame_page5 = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        #self.frame_page4.pack(expand=True, fill="both", padx=10, pady=10)
        self.page5_lable = customtkinter.CTkLabel(self.frame_page5, text='1.请选择系统导出班级考勤excel文件：')
        self.page5_lable.grid(row=0, column=0, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page5_entry1 = customtkinter.CTkEntry(master=self.frame_page5,placeholder_text="请选择班级考勤文件或者粘贴文件路径",state="normal", width=500, height=20, corner_radius=3)
        self.page5_entry1.grid(row=0, column=1, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page5_button1 = customtkinter.CTkButton(self.frame_page5, width=10, height=20, corner_radius=3,
                                                text='选择文件', command=self.page5_file1)
        self.page5_button1.grid(row=0, column=2, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page5_lable1 = customtkinter.CTkLabel(self.frame_page5, text='提示，请参考下图：（导出后不要更改表格，同一堂前后节迟到按一次计算）',text_color='red')
        self.page5_lable1.grid(row=1, column=0, columnspan=3, padx=(20, 10), sticky="w")
        self.page5_image = customtkinter.CTkImage(Image.open(os.path.join(self.image_path, "class_attendance.png")), size=(700, 267))
        self.page5_image_lable = customtkinter.CTkLabel(self.frame_page5, text="", image=self.page5_image)
        self.page5_image_lable.grid(row=2, column=0, columnspan=3, padx=20, pady=10)
        self.page5_lable2 = customtkinter.CTkLabel(self.frame_page5, text='2.请选择输出目录：')
        self.page5_lable2.grid(row=3, column=0, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page5_entry2 = customtkinter.CTkEntry(master=self.frame_page5,placeholder_text="请选择或者粘贴输出路径，不选择则输出到考勤同名目录",state="normal", width=500, height=20, corner_radius=3)
        self.page5_entry2.grid(row=3, column=1, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page5_button2 = customtkinter.CTkButton(self.frame_page5, width=10, height=20, corner_radius=3,
                                                text='选择目录', command=self.page5_dir1)
        self.page5_button2.grid(row=3, column=2, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page5_lable3 = customtkinter.CTkLabel(self.frame_page5, text='3.处理并输出：')
        self.page5_lable3.grid(row=4, column=0, padx=(20, 10), pady=(20, 20), sticky="w")
        self.page5_button3 = customtkinter.CTkButton(self.frame_page5, width=10, height=20, corner_radius=3,
                                                text='开始处理', command=self.page5_process)
        self.page5_button3.grid(row=4, column=1, padx=(20, 10), pady=(20, 20), sticky="w")
        self.select_frame_by_name("课程考勤")

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.page1_button.configure(fg_color=("gray75", "gray25") if name == "课程考勤" else "transparent")
        self.page2_button.configure(fg_color=("gray75", "gray25") if name == "考核记录单" else "transparent")
        self.page3_button.configure(fg_color=("gray75", "gray25") if name == "上传模版" else "transparent")
        self.page4_button.configure(fg_color=("gray75", "gray25") if name == "成绩分析表" else "transparent")
        self.page5_button.configure(fg_color=("gray75", "gray25") if name == "班主任考勤表" else "transparent")

        # show selected frame
        if name == "课程考勤":
            self.frame_page1.pack(expand=True, fill="both", padx=10, pady=10)
        else:
            self.frame_page1.pack_forget()
        if name == "考核记录单":
            self.frame_page2.pack(expand=True, fill="both", padx=10, pady=10)
        else:
            self.frame_page2.pack_forget()
        if name == "上传模版":
            self.frame_page3.pack(expand=True, fill="both", padx=10, pady=10)
        else:
            self.frame_page3.pack_forget()
        if name == "成绩分析表":
            self.frame_page4.pack(expand=True, fill="both", padx=10, pady=10)
        else:
            self.frame_page4.pack_forget()
        if name == "班主任考勤表":
            self.frame_page5.pack(expand=True, fill="both", padx=10, pady=10)
        else:
            self.frame_page5.pack_forget()
        #文字标签-主题切换
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="显示模式:", anchor="w")
        self.appearance_mode_label.grid(row=0, column=0, padx=10, pady=(10, 10))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, width=100, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=0, column=1, padx=0, pady=(10, 10))
        #文字标签-缩放比例
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=0, column=2, padx=10, pady=(10, 10))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, width=100, values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=0, column=3, padx=0, pady=(10, 10))

    def page1_button_event(self):
        self.select_frame_by_name("课程考勤")

    def page2_button_event(self):
        self.select_frame_by_name("考核记录单")

    def page3_button_event(self):
        self.select_frame_by_name("上传模版")

    def page4_button_event(self):
        self.select_frame_by_name("成绩分析表")

    def page5_button_event(self):
        self.select_frame_by_name("班主任考勤表")

#选取文件
    # def selectfile(self):
    #     self.selected_path = FileExplorer.askfile("选择考勤文件")
    #     self.filetext.delete("0.0", "end")
    #     self.filetext.insert("0.0",self.selected_path)
#选取文件 2
    def page1_file1(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page1_entry1.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page1_entry1.delete(0, customtkinter.END)  # 删除现有文本
            self.page1_entry1.insert(0, self.file_path)  # 插入新文本
            self.page1_entry1.configure(state="disabled")  # 设置为不可编辑状态
    def page1_file2(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page1_entry2.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page1_entry2.delete(0, customtkinter.END)  # 删除现有文本
            self.page1_entry2.insert(0, self.file_path)  # 插入新文本
            self.page1_entry2.configure(state="disabled")  # 设置为不可编辑状态
    def page2_file1(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page2_entry11.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page2_entry11.delete(0, customtkinter.END)  # 删除现有文本
            self.page2_entry11.insert(0, self.file_path)  # 插入新文本
            self.page2_entry11.configure(state="disabled")  # 设置为不可编辑状态
    def page2_file2(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page2_entry12.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page2_entry12.delete(0, customtkinter.END)  # 删除现有文本
            self.page2_entry12.insert(0, self.file_path)  # 插入新文本
            self.page2_entry12.configure(state="disabled")  # 设置为不可编辑状态
    def page2_file3(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page2_entry13.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page2_entry13.delete(0, customtkinter.END)  # 删除现有文本
            self.page2_entry13.insert(0, self.file_path)  # 插入新文本
            self.page2_entry13.configure(state="disabled")  # 设置为不可编辑状态
    def page2_file4(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page2_entry14.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page2_entry14.delete(0, customtkinter.END)  # 删除现有文本
            self.page2_entry14.insert(0, self.file_path)  # 插入新文本
            self.page2_entry14.configure(state="disabled")  # 设置为不可编辑状态

    def page3_file1(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page3_entry6.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page3_entry6.delete(0, customtkinter.END)  # 删除现有文本
            self.page3_entry6.insert(0, self.file_path)  # 插入新文本
            self.page3_entry6.configure(state="disabled")  # 设置为不可编辑状态

    def page3_file2(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page3_entry7.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page3_entry7.delete(0, customtkinter.END)  # 删除现有文本
            self.page3_entry7.insert(0, self.file_path)  # 插入新文本
            self.page3_entry7.configure(state="disabled")  # 设置为不可编辑状态
    def page3_file3(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page3_entry8.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page3_entry8.delete(0, customtkinter.END)  # 删除现有文本
            self.page3_entry8.insert(0, self.file_path)  # 插入新文本
            self.page3_entry8.configure(state="disabled")  # 设置为不可编辑状态

    def page5_file1(self):
        self.file_path = customtkinter.filedialog.askopenfilename(title="请选择文件",filetypes=(("Excel files", "*.xls"), ("Excel files", "*.xlsx")))
        if self.file_path:
            self.page5_entry1.configure(state="normal")  # 先设置为正常状态以便更新文本
            self.page5_entry1.delete(0, customtkinter.END)  # 删除现有文本
            self.page5_entry1.insert(0, self.file_path)  # 插入新文本
            self.page5_entry1.configure(state="disabled")  # 设置为不可编辑状态

#选目录
    def page1_dir1(self):
        self.filename = customtkinter.filedialog.askdirectory()
        self.page1_entry3.delete(0, "end")
        self.page1_entry3.insert(0, self.filename)
    def page2_dir1(self):
        self.filename = customtkinter.filedialog.askdirectory()
        self.page2_entry15.delete(0, "end")
        self.page2_entry15.insert(0, self.filename)

    def page3_dir1(self):
        self.filename = customtkinter.filedialog.askdirectory()
        self.page3_entry9.delete(0, "end")
        self.page3_entry9.insert(0, self.filename)
    def page5_dir1(self):
        self.filename = customtkinter.filedialog.askdirectory()
        self.page5_entry2.delete(0, "end")
        self.page5_entry2.insert(0, self.filename)

    def page1_process(self):
        self.page1_massage = page_process.generate_attendance_sheet(self.date_picker.get_date(),self.page1_entry1.get(),self.page1_entry2.get(),self.page1_entry3.get())
        Message.showinfo('课程考勤表完成提醒', f'你的处理已经完成，输出位置为：\n {self.page1_massage} \n请查看！')

    def page2_process(self):
        q_values = [self.page2_entry1.get(),self.page2_entry2.get(),self.page2_entry3.get(),self.page2_entry4.get()]
        q_values = [int(i) for i in q_values]
        a_values = [self.page2_entry5.get(),self.page2_entry6.get(),self.page2_entry7.get(),self.page2_entry8.get(),self.page2_entry9.get(),self.page2_entry10.get()]
        a_values = [float(i) for i in a_values]
        self.page1_massage = page_process.calculate_grades(self.page2_entry12.get(),self.page2_entry11.get(),q_values,a_values,self.page2_entry13.get(),self.page2_entry14.get(),self.page2_entry15.get())
        Message.showinfo('课程考勤表完成提醒', f'你的处理已经完成，输出位置为：\n {self.page1_massage} \n请查看！')

    def page3_process(self):
        b_values = [self.page3_entry1.get(),self.page3_entry2.get(),self.page3_entry3.get(),self.page3_entry4.get(),self.page3_entry5.get()]
        b_values = [int(i) for i in b_values]

        self.page3_massage = page_process.fill_scores(self.page3_entry6.get(),self.page3_entry7.get(),self.page3_entry8.get(),b_values,self.page3_entry9.get())
        Message.showinfo('课程考勤表完成提醒', f'你的处理已经完成，输出位置为：\n {self.page3_massage} \n请查看！')

    def page5_process(self):
        self.page5_massage = page_process.process_class(self.page5_entry1.get(),self.page5_entry2.get())
        Message.showinfo('班级考勤表完成提醒', f'你的处理已经完成，输出位置为：\n {self.page5_massage} \n请查看！')

    def page1_top0(self):
        if self.page1_toplevel0 is None or not self.page1_toplevel0.winfo_exists():
            self.page1_toplevel0 = topnotice.ToplevelWindow0(self)  # create window if its None or destroyed
        else:
            self.page1_toplevel0.focus()  # if window exists focus it

    def page1_top1(self):
        if self.page1_toplevel1 is None or not self.page1_toplevel1.winfo_exists():
            self.page1_toplevel1 = topnotice.ToplevelWindow1(self)  # create window if its None or destroyed
        else:
            self.page1_toplevel1.focus()  # if window exists focus it

    def page3_top0(self):
        if self.page3_toplevel0 is None or not self.page3_toplevel0.winfo_exists():
            self.page3_toplevel0 = topnotice.ToplevelWindow2(self)  # create window if its None or destroyed
        else:
            self.page3_toplevel0.focus()  # if window exists focus it
    #改变主题模式函数
    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)
    #改变缩放比例函数
    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

if __name__ == "__main__":
    app = App()
    #pywinstyles.apply_style(app, 'acrylic')
    app.mainloop()
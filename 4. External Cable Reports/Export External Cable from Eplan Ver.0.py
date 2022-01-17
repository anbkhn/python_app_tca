import tkinter as tk
from tkinter import ttk, filedialog as fd, messagebox as mg
from types import SimpleNamespace
from typing import DefaultDict
from numpy.core.fromnumeric import transpose
import xlwings as xw
import pandas as pd
import numpy as np
from xlwings import constants
from xlwings.constants import DeleteShiftDirection
import random

class exportCables(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Export External Cables")
        self.geometry("1000x600")
   

        # Chia làm 3 dòng, 1 cột
        self.rowconfigure(0, weight=4)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)
        self.columnconfigure(0, weight=1)

        # Tạo 3 label frames: settings, dữ liệu, thực thi
        setting_lbframe = tk.LabelFrame(self, text="Cấu hình")
        data_lbframe = tk.LabelFrame(self, text="Dữ liệu")
        execute_lbframe = tk.LabelFrame(self, text="Thực thi")

        # Đặt vị trí cho 3 frames này
        lbframelist = (setting_lbframe,data_lbframe,execute_lbframe)
        for i in range(len(lbframelist)):
            lbframelist[i].grid(sticky="NWES", column=0, row=i, padx=(5,5), pady=(5,5))
        
        # Xử lý cấu trúc cho từng frame
        # Cấu hình cột cho frame setting
        setting_lbframe.rowconfigure(0, weight=5) 
        setting_lbframe.rowconfigure(1, weight=1) 
        setting_lbframe.columnconfigure((0,1,2,3,4), weight=1) 

        # Tạo Spare cable
        self.spare_cable = tk.StringVar(value="20")
        spare_cable_lb = ttk.Label(setting_lbframe, text=r"Nhập % core dự phòng:")
        spare_cable_entry = ttk.Entry(setting_lbframe, width=10, textvariable=self.spare_cable, state="enable")

        # Đặt Spare này vào Setting Frame
        spare_cable_lb.grid(sticky="W", column=0, row=1, padx=(5,5), pady=(5,5))
        spare_cable_entry.grid(sticky="E", column=0, row=1, padx=(5,5), pady=(5,5))     

        # Tạo các label frame cho từng loại cáp: cáp nguồn, cáp tín hiệu, cáp trip, cáp CT, cáp VT
        pow_cable_lbframe = tk.LabelFrame(setting_lbframe, text="Cáp mạch nguồn")
        sig_cable_lbframe = tk.LabelFrame(setting_lbframe, text="Cáp mạch tín hiệu")
        trip_cable_lbframe = tk.LabelFrame(setting_lbframe, text="Cáp mạch Trip")
        ct_cable_lbframe = tk.LabelFrame(setting_lbframe, text="Cáp mạch dòng")
        vt_cable_lbframe = tk.LabelFrame(setting_lbframe, text="Cáp mạch áp")

        # Đặt vị trí các label frame này vào label frame Setting
        setframelist = (pow_cable_lbframe, sig_cable_lbframe, trip_cable_lbframe, ct_cable_lbframe, vt_cable_lbframe)
        for i in range(len(setframelist)):
            setframelist[i].grid(sticky="NWES", column=i, row=0, padx=(5,5), pady=(5,5))
        
        # Cấu hình cột cho từng frame trong frame Setting
        for i in setframelist:
            i.columnconfigure(0, weight=3)
            i.columnconfigure(1, weight=1)

        # Tạo các label và check box trong label frame cáp nguồn 
        p2x2_5_lb = ttk.Label(pow_cable_lbframe, text="2x2.5mm2")
        p4x2_5_lb = ttk.Label(pow_cable_lbframe, text="4x2.5mm2")
        # p2x4_lb = ttk.Label(pow_cable_lbframe, text="2x4mm2")
        # p2x6_lb = ttk.Label(pow_cable_lbframe, text="2x6mm2")
        # p4x4_lb = ttk.Label(pow_cable_lbframe, text="4x4mm2")
        # p4x6_lb = ttk.Label(pow_cable_lbframe, text="4x6mm2")
        power_cables_lb_list = (p2x2_5_lb,p4x2_5_lb)

        self.p2x2_5_opt = tk.IntVar(value=1)
        self.p4x2_5_opt = tk.IntVar(value=1)
        # self.p2x4_opt = tk.IntVar(value=0)
        # self.p2x6_opt = tk.IntVar(value=0)
        # self.p4x4_opt = tk.IntVar(value=0)
        # self.p4x6_opt = tk.IntVar(value=0)
        p2x2_5_chk = ttk.Checkbutton(pow_cable_lbframe, variable=self.p2x2_5_opt, onvalue=1, offvalue=0)
        p4x2_5_chk = ttk.Checkbutton(pow_cable_lbframe, variable=self.p4x2_5_opt, onvalue=1, offvalue=0)
        # p2x4_chk = ttk.Checkbutton(pow_cable_lbframe, variable=self.p2x4_opt, onvalue=1, offvalue=0)
        # p2x6_chk = ttk.Checkbutton(pow_cable_lbframe, variable=self.p2x6_opt, onvalue=1, offvalue=0)
        # p4x4_chk = ttk.Checkbutton(pow_cable_lbframe, variable=self.p4x4_opt, onvalue=1, offvalue=0)
        # p4x6_chk = ttk.Checkbutton(pow_cable_lbframe, variable=self.p4x6_opt, onvalue=1, offvalue=0)
        power_cables_chk_list = (p2x2_5_chk,p4x2_5_chk)

        # Cấu hình row cho label frame cáp nguồn
        pow_cable_lbframe.rowconfigure((0,1,2,3,4,5,6,7,8, 9), weight=1)

        # Đặt các label và check box này trong lable frame cáp nguồn
        for i in range(len(power_cables_lb_list)):
            power_cables_lb_list[i].grid(sticky="NWES", column=0, row=i, padx=(5,5), pady=(5,5))
            power_cables_chk_list[i].grid(sticky="NWES", column=1, row=i, padx=(5,5), pady=(5,5))


        # Tạo các label và check box trong label frame cáp tín hiệu 
        s19x1_5_lb = ttk.Label(sig_cable_lbframe, text="19x1.5mm2")
        s14x1_5_lb = ttk.Label(sig_cable_lbframe, text="14x1.5mm2")
        s10x1_5_lb = ttk.Label(sig_cable_lbframe, text="10x1.5mm2")
        s7x1_5_lb = ttk.Label(sig_cable_lbframe, text="7x1.5mm2")
        s4x1_5_lb = ttk.Label(sig_cable_lbframe, text="4x1.5mm2")
        signal_cables_lb_list = (s19x1_5_lb,s14x1_5_lb,s10x1_5_lb,s7x1_5_lb,s4x1_5_lb)

        self.s19x1_5_opt = tk.IntVar(value=1)
        self.s14x1_5_opt = tk.IntVar(value=1)
        self.s10x1_5_opt = tk.IntVar(value=1)
        self.s7x1_5_opt = tk.IntVar(value=1)
        self.s4x1_5_opt = tk.IntVar(value=1)
        s19x1_5_chk = ttk.Checkbutton(sig_cable_lbframe, variable=self.s19x1_5_opt, onvalue=1, offvalue=0)
        s14x1_5_chk = ttk.Checkbutton(sig_cable_lbframe, variable=self.s14x1_5_opt, onvalue=1, offvalue=0)
        s10x1_5_chk = ttk.Checkbutton(sig_cable_lbframe, variable=self.s10x1_5_opt, onvalue=1, offvalue=0)
        s7x1_5_chk = ttk.Checkbutton(sig_cable_lbframe, variable=self.s7x1_5_opt, onvalue=1, offvalue=0)
        s4x1_5_chk = ttk.Checkbutton(sig_cable_lbframe, variable=self.s4x1_5_opt, onvalue=1, offvalue=0)
        signal_cables_chk_list = (s19x1_5_chk,s14x1_5_chk,s10x1_5_chk,s7x1_5_chk,s4x1_5_chk)

        # Cấu hình row cho label frame cáp tín hiệu
        sig_cable_lbframe.rowconfigure((0,1,2,3,4,5), weight=1)

        # Đặt các label và check box này trong lable frame cáp tín hiệu
        for i in range(len(signal_cables_lb_list)):
            signal_cables_lb_list[i].grid(sticky="NWES", column=0, row=i, padx=(5,5), pady=(5,5))
            signal_cables_chk_list[i].grid(sticky="NWES", column=1, row=i, padx=(5,5), pady=(5,5))
        
        # Tạo các label và check box trong label frame cáp trip
        t10x1_5_lb = ttk.Label(trip_cable_lbframe, text="10x1.5mm2")
        t7x1_5_lb = ttk.Label(trip_cable_lbframe, text="7x1.5mm2")
        t4x1_5_lb = ttk.Label(trip_cable_lbframe, text="4x1.5mm2")
        t10x2_5_lb = ttk.Label(trip_cable_lbframe, text="10x2.5mm2")
        t7x2_5_lb = ttk.Label(trip_cable_lbframe, text="7x2.5mm2")
        t4x2_5_lb = ttk.Label(trip_cable_lbframe, text="4x2.5mm2")
        trip_cables_lb_list = (t10x1_5_lb,t7x1_5_lb,t4x1_5_lb,t10x2_5_lb,t7x2_5_lb,t4x2_5_lb)

        self.t10x1_5_opt = tk.IntVar(value=1)
        self.t7x1_5_opt = tk.IntVar(value=1)
        self.t4x1_5_opt = tk.IntVar(value=1)
        self.t10x2_5_opt = tk.IntVar(value=0)
        self.t7x2_5_opt = tk.IntVar(value=0)
        self.t4x2_5_opt = tk.IntVar(value=0)
        t10x1_5_chk = ttk.Checkbutton(trip_cable_lbframe, variable=self.t10x1_5_opt, onvalue=1, offvalue=0)
        t7x1_5_chk = ttk.Checkbutton(trip_cable_lbframe, variable=self.t7x1_5_opt, onvalue=1, offvalue=0)
        t4x1_5_chk = ttk.Checkbutton(trip_cable_lbframe, variable=self.t4x1_5_opt, onvalue=1, offvalue=0)
        t10x2_5_chk = ttk.Checkbutton(trip_cable_lbframe, variable=self.t10x2_5_opt, onvalue=1, offvalue=0)
        t7x2_5_chk = ttk.Checkbutton(trip_cable_lbframe, variable=self.t7x2_5_opt, onvalue=1, offvalue=0)
        t4x2_5_chk = ttk.Checkbutton(trip_cable_lbframe, variable=self.t4x2_5_opt, onvalue=1, offvalue=0)
        trip_cables_chk_list = (t10x1_5_chk,t7x1_5_chk,t4x1_5_chk,t10x2_5_chk,t7x2_5_chk,t4x2_5_chk)

        # Cấu hình row cho label frame cáp tín hiệu
        trip_cable_lbframe.rowconfigure((0,1,2,3,4,5,6), weight=1)

        # Đặt các label và check box này trong lable frame cáp tín hiệu
        for i in range(len(trip_cables_lb_list)):
            trip_cables_lb_list[i].grid(sticky="NWES", column=0, row=i, padx=(5,5), pady=(5,5))
            trip_cables_chk_list[i].grid(sticky="NWES", column=1, row=i, padx=(5,5), pady=(5,5))

        # Tạo các label và check box trong label frame cáp mạch dòng
        ct4x4_lb = ttk.Label(ct_cable_lbframe, text="4x4mm2")
        ct2x4_lb = ttk.Label(ct_cable_lbframe, text="2x4mm2")
        ct_cables_lb_list = (ct4x4_lb,ct2x4_lb)

        self.ct4x4_opt = tk.IntVar(value=1)
        self.ct2x4_opt = tk.IntVar(value=1)
        ct4x4_chk = ttk.Checkbutton(ct_cable_lbframe, variable=self.ct4x4_opt, onvalue=1, offvalue=0)
        ct2x4_chk = ttk.Checkbutton(ct_cable_lbframe, variable=self.ct2x4_opt, onvalue=1, offvalue=0)
        ct_cables_chk_list = (ct4x4_chk,ct2x4_chk)

        # Cấu hình row cho label frame cáp mạch dòng
        ct_cable_lbframe.rowconfigure((0,1,2,3,4,5,6,7,8, 9), weight=1)

        # Đặt các label và check box này trong lable frame cáp mạch dòng
        for i in range(len(ct_cables_lb_list)):
            ct_cables_lb_list[i].grid(sticky="NWES", column=0, row=i, padx=(5,5), pady=(5,5))
            ct_cables_chk_list[i].grid(sticky="NWES", column=1, row=i, padx=(5,5), pady=(5,5))

        # Tạo các label và check box trong label frame cáp mạch áp
        vt4x2_5_lb = ttk.Label(vt_cable_lbframe, text="4x2.5mm2")
        vt2x2_5_lb = ttk.Label(vt_cable_lbframe, text="2x2.5mm2")
        vt_cables_lb_list = (vt4x2_5_lb,vt2x2_5_lb)

        self.vt4x2_5_opt = tk.IntVar(value=1)
        self.vt2x2_5_opt = tk.IntVar(value=1)
        vt4x2_5_chk = ttk.Checkbutton(vt_cable_lbframe, variable=self.vt4x2_5_opt, onvalue=1, offvalue=0)
        vt2x2_5_chk = ttk.Checkbutton(vt_cable_lbframe, variable=self.vt2x2_5_opt, onvalue=1, offvalue=0)
        vt_cables_chk_list = (vt4x2_5_chk,vt2x2_5_chk)

        # Cấu hình row cho label frame cáp mạch áp
        vt_cable_lbframe.rowconfigure((0,1,2,3,4,5,6,7,8, 9), weight=1)

        # Đặt các label và check box này trong lable frame cáp mạch dòng
        for i in range(len(vt_cables_lb_list)):
            vt_cables_lb_list[i].grid(sticky="NWES", column=0, row=i, padx=(5,5), pady=(5,5))
            vt_cables_chk_list[i].grid(sticky="NWES", column=1, row=i, padx=(5,5), pady=(5,5))


        # Cấu hình cột cho frame data
        data_lbframe.rowconfigure((0,1), weight=1) 
        data_lbframe.columnconfigure(0, weight=2)
        data_lbframe.columnconfigure(1, weight=1)
        data_lbframe.columnconfigure(2, weight=4)

        # Tạo các label, button và Entry cho frame data.
        self.data_file_text = tk.StringVar()
        data_label = ttk.Label(data_lbframe, text="File Connection")
        data_button = ttk.Button(data_lbframe, text="Chọn File", command=self.data_choose_file)
        data_entry = ttk.Entry(data_lbframe, width=50, textvariable=self.data_file_text, state="disable")

        self.result_file_text = tk.StringVar()
        export_label = ttk.Label(data_lbframe, text="Form Cáp")
        export_button = ttk.Button(data_lbframe, text="Chọn File", command=self.result_choose_file)
        export_entry = ttk.Entry(data_lbframe, width=50, textvariable=self.result_file_text, state="disable")

        # Đặt vào label Frame Data
        data_label.grid(sticky="EW", column=0, row= 0, padx=5, pady=5)
        data_button.grid(sticky="EW", column=1, row= 0, padx=5, pady=5)
        data_entry.grid(sticky="EW", column=2, row= 0, padx=5, pady=5)

        export_label.grid(sticky="EW", column=0, row= 1, padx=5, pady=5)
        export_button.grid(sticky="EW", column=1, row= 1, padx=5, pady=5)
        export_entry.grid(sticky="EW", column=2, row= 1, padx=5, pady=5)

        # Cấu hình cột cho execute frame
        execute_lbframe.columnconfigure((0,1,2,3,4), weight=1)
        execute_lbframe.rowconfigure(0, weight=1)

        # Tạo các nút ấn trong frame thực thi
        execute_button = ttk.Button(execute_lbframe, text="Xuất Cáp", command=self.exportcables)
        exit_button = ttk.Button(execute_lbframe, text="Thoát", command=self.destroy)

        # Đặt các nút ấn trong frame thực thi
        execute_button.grid(sticky="EW", column=1, row= 0, padx=5, pady=5)
        exit_button.grid(sticky="EW", column=3, row= 0, padx=5, pady=5)

        excel_file_extensions = ['*.xls', '*.xlsx', '*.xlsb',  '*.xlsm']
        self.ftypes = [    
            ('test files', excel_file_extensions), 
            ('All files', '*'), 
        ]

    def data_choose_file(self):
        filename =  fd.askopenfilename(initialdir = "/a",title = "Hãy chọn file connection xuất ra từ eplan",filetypes=self.ftypes)
        self.data_file_text.set(filename)
    
    def result_choose_file(self):
        filename =  fd.askopenfilename(initialdir = "/a",title = "Hãy chọn form để xuất file cáp",filetypes=self.ftypes)
        self.result_file_text.set(filename)
    
    def validate(self):
        # Lấy các thông tin setting từ form
        power_core_list_get_from_form = [4*self.p4x2_5_opt.get(), 2*self.p2x2_5_opt.get()]
        power_name_list_get_from_form = ["4x2.5mm2"*self.p4x2_5_opt.get(), "2x2.5mm2"*self.p2x2_5_opt.get()]
        
        self.power_core_list = [i for i in power_core_list_get_from_form if i]
        self.power_name_list = [i for i in power_name_list_get_from_form if i]

        if len(self.power_core_list) == 0:
            mg.showinfo("Thông báo","Hãy tích chọn cáp nguồn")
            return 0 

        signal_core_list_get_from_form = [
            19*self.s19x1_5_opt.get(), 
            14*self.s14x1_5_opt.get(),
            10*self.s10x1_5_opt.get(),
            7*self.s7x1_5_opt.get(),
            4*self.s4x1_5_opt.get(),
            ]
        signal_name_list_get_from_form = [
            "19x1.5mm2"*self.s19x1_5_opt.get(), 
            "14x1.5mm2"*self.s14x1_5_opt.get(),
            "10x1.5mm2"*self.s10x1_5_opt.get(),
            "7x1.5mm2"*self.s7x1_5_opt.get(),
            "4x1.5mm2"*self.s4x1_5_opt.get(),
            ]

        self.signal_core_list = [i for i in signal_core_list_get_from_form if i]
        self.signal_name_list = [i for i in signal_name_list_get_from_form if i]

        if len(self.signal_core_list) == 0:
            mg.showinfo("Thông báo","Hãy tích chọn cáp tín hiệu")
            return 0  

        trip_core_list_get_from_form = [
            10*self.t10x1_5_opt.get(), 
            7*self.t7x1_5_opt.get(),
            4*self.t4x1_5_opt.get(),
            10*self.t10x2_5_opt.get(),
            7*self.t7x2_5_opt.get(),
            4*self.t4x2_5_opt.get(),
            ]
        trip_name_list_get_from_form = [
            "10x1.5mm2"*self.t10x1_5_opt.get(), 
            "7x1.5mm2"*self.t7x1_5_opt.get(),
            "4x1.5mm2"*self.t4x1_5_opt.get(),
            "10x2.5mm2"*self.t10x2_5_opt.get(),
            "7x2.5mm2"*self.t7x2_5_opt.get(),
            "4x2.5mm2"*self.t4x2_5_opt.get(),
            ]
        self.trip_core_list = [i for i in trip_core_list_get_from_form if i]
        self.trip_name_list = [i for i in trip_name_list_get_from_form if i]

        if len(self.trip_core_list) > len(set(self.trip_core_list)):
            mg.showinfo("Thông báo","Hãy tích chọn một loại tiết diện cho mạch Trip") 
            return 0 

        if len(self.trip_core_list) == 0:
            mg.showinfo("Thông báo","Hãy tích chọn cáp mạch Trip") 
            return 0 

        ct_core_list_get_from_form = [4*self.ct4x4_opt.get(), 2*self.ct2x4_opt.get()]
        ct_name_list_get_from_form = ["4x4mm2"*self.ct4x4_opt.get(), "2x4mm2"*self.ct2x4_opt.get()]

        self.ct_core_list = [i for i in ct_core_list_get_from_form if i]
        self.ct_name_list = [i for i in ct_name_list_get_from_form if i]

        if len(self.ct_core_list) == 0:
            mg.showinfo("Thông báo","Hãy tích chọn cáp mạch dòng") 
            return 0 


        vt_core_list_get_from_form = [4*self.vt4x2_5_opt.get(), 2*self.vt2x2_5_opt.get()]
        vt_name_list_get_from_form = ["4x2.5mm2"*self.vt4x2_5_opt.get(), "2x2.5mm2"*self.vt2x2_5_opt.get()]

        self.vt_core_list = [i for i in vt_core_list_get_from_form if i]
        self.vt_name_list = [i for i in vt_name_list_get_from_form if i]

        if len(self.vt_core_list) == 0:
            mg.showinfo("Thông báo","Hãy tích chọn cáp mạch áp") 
            return 0 

        if self.spare_cable.get().isnumeric():
            self.spare_no_from_form = float(self.spare_cable.get())/100
        else:
            mg.showinfo("Thông báo","Hãy điền đúng dạng số") 
            return 0 
        
        self.cab_dict = {0:self.signal_core_list, 1:self.power_core_list, 2:self.trip_core_list, 3:self.ct_core_list, 4:self.vt_core_list}
        self.cab_dict_name = {0:self.signal_name_list, 1:self.power_name_list, 2:self.trip_name_list, 3:self.ct_name_list, 4:self.vt_name_list}
        return 1

    def exportcables(self):

        if(self.validate()):
            # Kiểm tra lại các thông tin trên form 
            if self.data_file_text.get() and self.result_file_text.get():

                # Sử dụng câu lệnh with này để tránh luôn chạy tiến trình chạy Excel trong task manager
                with xw.App(visible=False) as app:

                    wb = xw.Book(self.data_file_text.get())
                    # Kết nối đến sheet đầu tiên của file dự trù
                    sh = wb.sheets[0]

                    # Tìm dòng cuối của dữ liệu trong file dự trù
                    last_row = sh.range(f'E{sh.cells.last_cell.row}').end('up').row
                    
                    # Lấy tất cả các dữ liệu trong file dự trù vào dataframe
                    df = pd.DataFrame(sh[f'E3:F{last_row}'].value, columns=["source", "target"])

                    # Tách cột source thành cột tủ và thiết bị
                    df[["cub_s","dev_s"]] = df["source"].str.split("-", expand=True).drop(columns=[2,3])

                    # Tách cột target thành cột tủ và thiết bị
                    df[["cub_t","dev_t"]] = df["target"].str.split("-", expand=True).drop(columns=[2,3])

                    # Sort dữ liệu theo cột tủ nguồn trước sau đó theo cột tủ đích sau
                    df = df.sort_values(["cub_s","cub_t"])

                    # Xóa dữ liệu dây nội bộ tủ
                    df = df.drop(df[df["cub_s"] == df["cub_t"]].index)

                    # Đặt lại index cho dataframe
                    df.index = np.arange(1, len(df)+1)

                    # Liệt kê các tủ trong cột cub_s
                    cub_s_unique = df["cub_s"].unique()

                    # Tạo 1 hàm kiểm tra loại cáp
                    def type_cab_check(terminal):
                        # Nếu hàng kẹp bắt đầu bằng X4 nghĩa là hàng kẹp dòng
                        if terminal.startswith("X4"):

                            # Trả về giá trị 3 đúng với self.cab_dict
                            return 3
                        # Nếu hàng kẹp bắt đầu bằng X5 nghĩa là hàng kẹp áp
                        elif terminal.startswith("X5"):
                            return  4
                        
                        # Nếu hàng kẹp bắt đầu bằng XAC hoặc XDC nghĩa là hàng kẹp nguồn
                        elif terminal.startswith("XAC") or terminal.startswith("XDC"):
                            return 1
                        
                        # Nếu hàng kẹp bắt đầu bằng X10 hoặc X20 nghĩa là hàng kẹp trip
                        elif terminal.startswith("X10") or terminal.startswith("X20"):
                            return 2

                        # Nếu hàng kẹp bắt đầu bằng P nghĩa là hàng kẹp áp hoặc dòng đang nối đến CT và VT
                        elif terminal.startswith("P"):
                            return 7
                        # Tất cả các hàng kẹp còn lại thuộc hàng kẹp tín hiệu
                        else:
                            return 0
                    # Hàm kiểm tra core để tìm số loại cáp cần sử dụng, số lượng core đã sử dụng, và số lượng core còn dư
                    # Đưa vào hàm truyền 4 tham số
                    # no_terminal: số lượng hàng kẹp; core_practical: list core thực tế, core_available: list core có thể dùng được, core_spare: list core luôn sẵn sàng dự trù
                    def checkcore(no_terminal, core_practical, core_available, core_spare, cab_name_practical):
                        use_core = []
                        core_quantity = []
                        spare_quantity = []
                        use_cab = []

                        # Kiểm tra đến khi nào mà biến no_terminal lớn hơn giá trị nhỏ nhất trong list core_available
                        while (no_terminal >= min(core_available)):

                            # Làm 1 vòng chạy từ đầu danh sách đến cuối danh sách
                            for i in range(len(core_available)):

                                # Nếu số lượng hàng kẹp mà lớn hơn giá trị tìm được trong vòng lặp đầu tiên thì sẽ thực thi lệnh và thoát ngay
                                if no_terminal >= core_available[i]:

                                    # Ghi loại core được sử dụng trong trường hợp này
                                    use_core.append(core_practical[i])

                                    # Ghi loại cáp được sử dụng trong trường hợp này
                                    use_cab.append(cab_name_practical[i])

                                    # Ghi số lượng core được sử dụng trong trường hợp này
                                    core_quantity.append(int(int(no_terminal)/int(core_available[i])))

                                    # Ghi số lượng core còn dư trong trường hợp này
                                    spare_quantity.append(int(core_spare[i]))

                                    # Thay đổi lại giá trị hàng kẹp để tìm ra hàng kẹp còn dư rồi lại tiếp tục vòng lặp
                                    no_terminal = no_terminal-int(int(no_terminal)/int(core_available[i]))*int(core_available[i])

                                    # Thoát vòng lặp for để tiếp tục vòng lặp while
                                    break
                        # Nếu trong trường hợp mà số lượng hàng kẹp >0 à nhỏ hơn giá trị nhỏ nhất 
                        if no_terminal > 0 and no_terminal < min(core_available):

                            # Loại core được sử dụng chính là core có giá trị nhỏ nhất trong danh sách
                            use_core.append(min(core_practical)) 

                            # Loại core được sử dụng chính là core có giá trị nhỏ nhất trong danh sách
                            use_cab.append(cab_name_practical[len(cab_name_practical)-1]) 

                            # Số lượng core nhỏ nhất lúc đó chắc chắn là 1
                            core_quantity.append(1)

                            # Tính toán lại số lượng core còn dư
                            spare_quantity.append(int(min(core_available))-no_terminal+ int(min(core_spare)))

                        # Trả lại giá trị cho các danh sách
                        return use_core, core_quantity, spare_quantity, use_cab

                    df_cub_s = pd.DataFrame()

                    # Chạy vòng lặp cho các tủ
                    for cub_s in cub_s_unique:


                        # Tạo dataframe riêng cho từng tủ
                        dfa = df[df["cub_s"] == cub_s]

                        # Tìm kiếm các tủ duy nhất được kết nối cáp với tủ đang chạy vòng lặp
                        cub_t_unique = dfa["cub_t"].unique()

                        # Chạy vòng lặp cho các tủ target
                        for cub_t in cub_t_unique:
                            
                            # Tạo dataframe với từng tủ target 
                            dfb = dfa[dfa["cub_t"]==cub_t]
                        
                            # Sort dfb theo cột dev_s
                            dfb = dfb.sort_values("dev_s")

                            # Tạo thêm 2 cột trong dfb với tên là ter_s: Tên hàng kẹp và pin_s: Tên chân đấu 
                            dfb[["ter_s","pin_s"]]=dfb["dev_s"].str.split(":",expand = True)

                            # Tạo thêm 2 cột trong dfb với tên là ter_t: Tên hàng kẹp và pin_t: Tên chân đấu 
                            dfb[["ter_t","pin_t"]]=dfb["dev_t"].str.split(":",expand = True)

                            # Tìm tất cả các loại hàng kẹp trong dataframe dfb
                            terminal_s_list = dfb["ter_s"].unique()

                            # Tìm tất cả các loại hàng kẹp loại tín hiệu
                            signal_list = [i for i in terminal_s_list if type_cab_check(i) == 0]        

                            df_cub_t = pd.DataFrame()

                            if len(signal_list):
                                dftemp = dfb[dfb["ter_s"].isin(signal_list)]

                                # Đếm số lượng hàng kẹp
                                ter_count = dftemp["ter_s"].count()

                                # Kiểm tra loại cáp được sử dụng cho loại hàng kẹp này
                                core_practical = self.cab_dict[0]

                                cab_name_practical = self.cab_dict_name[0]

                                core_available = np.floor(np.array(core_practical)*(1-self.spare_no_from_form))

                                core_spare = np.array(core_practical)-core_available

                                # Tính toán cáp xem sử dụng 1 danh sách list các loại cáp nào, số lượng bao nhiêu, spare bao nhiêu, tên cáp là gì
                                cab_use, quan_cab_use, quan_spare_cab, cab_name = checkcore(ter_count,core_practical, core_available, core_spare, cab_name_practical)
                                
                                core_use_column_list = []
                                cab_name_list = []
                                for i in range(len(cab_use)):
                                    temp = []
                                    core_use_column_list.extend(list(np.arange(1, cab_use[i]+1-quan_spare_cab[i]))*quan_cab_use[i])
                                    temp.append(cab_name[i])
                                    cab_name_list = cab_name_list + temp*quan_cab_use[i]
                                

                                dftemp["core_use"] = core_use_column_list

                                dftemp.loc[dftemp["core_use"] == 1,"cab_name"] = cab_name_list

                                df_cub_t = df_cub_t.append(dftemp)

                            #  Chạy vòng lặp cho các hàng kẹp khác hàng kẹp tín hiệu
                            for i in terminal_s_list:
                                if type_cab_check(i) != 0 and type_cab_check(i) !=7 : 
                                    dftemp = dfb[dfb["ter_s"] == i]      
                                # else:
                                #     dftemp = dfb[dfb["ter_s"].isin(signal_list)]

                                    # Đếm số lượng hàng kẹp
                                    ter_count = dftemp["ter_s"].count()

                                    # Kiểm tra loại cáp được sử dụng cho loại hàng kẹp này
                                    core_practical = self.cab_dict[type_cab_check(i)]

                                    cab_name_practical = self.cab_dict_name[type_cab_check(i)]

                                    # Kiểm tra số core cáp thực tế được sử dụng
                                    core_available = core_practical

                                    # Tính toán số core spare
                                    core_spare = np.array(core_practical)-np.array(core_available)

                                    # Tính toán cáp xem sử dụng 1 danh sách list các loại cáp nào, số lượng bao nhiêu, spare bao nhiêu
                                    cab_use, quan_cab_use, quan_spare_cab, cab_name = checkcore(ter_count,core_practical, core_available, core_spare, cab_name_practical)
                                    
                                    # Đưa kết quả các lõi sử dụng vào 1 list
                                    core_use_column_list = []
                                    cab_name_list = []
                                    for i in range(len(cab_use)):
                                        temp = []
                                        core_use_column_list.extend(list(np.arange(1, cab_use[i]+1-quan_spare_cab[i]))*quan_cab_use[i])
                                        temp.append(cab_name[i])
                                        cab_name_list = cab_name_list + temp*quan_cab_use[i]

                                    # Tạo 1 cột core_use
                                    dftemp["core_use"] = core_use_column_list
                                   
                                    dftemp.loc[dftemp["core_use"] == 1,"cab_name"] = cab_name_list

                                    # Nối dftemp vào df_cub_t
                                    df_cub_t = df_cub_t.append(dftemp)
                                # Xử lý trường hợp nối với CT, VT ngoài                            
                                elif type_cab_check(i) != 0 and type_cab_check(i) ==7 : 

                                    dftemp = dfb[dfb["ter_s"] == i]  

                                    terminal_t_list = dftemp["ter_t"].unique()

                                    # Đưa kết quả các lõi sử dụng vào 1 list
                                    core_use_column_list = []
                                    cab_name_list = []

                                    for i in terminal_t_list:

                                        ter_count = dftemp[dftemp["ter_t"]==i]["ter_t"].count()

                                        # Kiểm tra loại cáp được sử dụng cho loại hàng kẹp này
                                        core_practical = self.cab_dict[type_cab_check(i)]

                                        cab_name_practical = self.cab_dict_name[type_cab_check(i)]

                                        # Kiểm tra số core cáp thực tế được sử dụng
                                        core_available = core_practical

                                        # Tính toán số core spare
                                        core_spare = np.array(core_practical)-np.array(core_available)
                                    
                                        # Tính toán cáp xem sử dụng 1 danh sách list các loại cáp nào, số lượng bao nhiêu, spare bao nhiêu
                                        cab_use, quan_cab_use, quan_spare_cab, cab_name = checkcore(ter_count,core_practical, core_available, core_spare, cab_name_practical)
                                        for i in range(len(cab_use)):
                                            temp = []
                                            core_use_column_list.extend(list(np.arange(1, cab_use[i]+1-quan_spare_cab[i]))*quan_cab_use[i])
                                            temp.append(cab_name[i])
                                            cab_name_list = cab_name_list + temp*quan_cab_use[i]

                                    # Tạo 1 cột core_use
                                    dftemp["core_use"] = core_use_column_list
                                
                                    dftemp.loc[dftemp["core_use"] == 1,"cab_name"] = cab_name_list

                                    # Nối dftemp vào df_cub_t
                                    df_cub_t = df_cub_t.append(dftemp)

                            # Nối các df_cub_t vào df_cub_s
                            df_cub_s = df_cub_s.append(df_cub_t)

                    

                    
                    
                    for i in ["name","cab_struct","ref","note","cab_gland","len"]:
                        df_cub_s[i] = ""

                                       

                    cab_count = df_cub_s[df_cub_s["cab_name"].notna()]["cab_name"].count()
                    cab_list_count = list(np.arange(1,cab_count+1))
                    cab_list_count_leading_zero = [str(i).zfill(3) for i in cab_list_count]
                    name_cab_list = [f"W{i}" for i in cab_list_count_leading_zero]

                    for i in range(len(df_cub_s[df_cub_s["cab_name"].notna()]["cab_name"].index)):

                        df_cub_s.loc[df_cub_s[df_cub_s["cab_name"].notna()]["cab_name"].index[i],"name"]= name_cab_list[i]

                    # Tạo 1 dataframe cho riêng cáp list
                    df_cable_list = df_cub_s[["name", "cab_struct", "cab_name","cab_gland","cub_s","cub_t","len", "note"]]
                    df_cable_list = df_cable_list[df_cable_list["cab_name"].notna()]
                    cable_list_unique = df_cable_list["cab_name"].unique()
                    cable_gland_template = ["PG21","PG29","PG25"]
                    cable_gland_list_template = [cable_gland_template[random.randint(0,2)] for i in range(len(cable_list_unique))]
            
                    cab_gland_list_formula = [f"=VLOOKUP(C{i},$J$3:$K${len(cable_list_unique)+2},2,0)" for i in range(3,df_cable_list["name"].count()+3)]
                    df_cable_list["cab_gland"] = cab_gland_list_formula

                    cub_unique = list(set(list(df_cable_list["cub_s"].unique()) + list(df_cable_list["cub_t"].unique())))
                    cub_unique_remove_equal = [str(i).replace("=","") for i in cub_unique]
                    cub_unique_remove_equal.sort()
                    # Thay các kí tự = thành các kí tự "'="
                    for i in ["source","target", "cub_s","cub_t"]:
                        df_cub_s[i] = df_cub_s[i].str.replace("=","'=")
                        
                    for i in ["cub_s","cub_t"]:
                        df_cable_list[i] = df_cable_list[i].str.replace("=","'=")

                    # Xóa các cột thừa trong df_cub_s
                    df_cub_s = df_cub_s.drop(columns=["cub_s", "dev_s", "cub_t", "dev_t", "ter_s", "pin_s","ter_t", "pin_t","cab_gland","len"])

                    # Thêm và thay đổi vị trí cột
                    df_cub_s = df_cub_s[["name","cab_struct","cab_name","core_use","source","target","ref","note"]]
                    # Mở file result
                    wb_result = xw.Book(self.result_file_text.get())

                    sh_result = wb_result.sheets[0]
                    sh_result.name = "cable connection"
                    
                    # Xóa các sheet khác nếu có
                    for sh in wb_result.sheets:
                        if sh != sh_result:
                            sh.delete()

                    # Tạo sheet mới ngay sau sheet result
                    new_sh = wb_result.sheets.add(name="cable list",after=sh_result.name) 

                    # Tìm dòng cuối của dữ liệu trong file kết quả
                    last_row_result = sh_result.range(f'E{sh_result.cells.last_cell.row}').end('up').row


                    # Xóa hết nội dung trong sheet đầu tiên
                    sh_result.range(f'1:{last_row_result}').api.Delete(DeleteShiftDirection.xlShiftUp) 

                    # Merge dòng 1
                    sh_result.range("A1:H1").merge()
                    new_sh.range("A1:H1").merge()
                    # Đặt tên cho các dòng tiêu đề
                    sh_result["A1"].value = "CABLE CONNECTION"
                    new_sh["A1"].value = "CABLE LIST"

                    sh_result["A2"].value = ["NAME","CABLE STRUCTURE","TYPE","CORE","FROM","TO","REFERENCE","NOTE"]
                    new_sh["A2"].value = ["NAME","CABLE STRUCTURE","TYPE","CABLE GLAND","FROM","TO","LENGTH","NOTE"]
                    new_sh["J2"].value = ["Cáp","Ốc Siết Cáp","Tủ","PG21","PG25","PG29"]
                    # Ghi kết quả vào A3
                    sh_result["A3"].value = df_cub_s.values.tolist()
                    
                    new_sh["A3"].value = df_cable_list.values.tolist()
                    new_sh["J3"].options(transpose=True).value = cable_list_unique
                    new_sh["K3"].options(transpose=True).value = cable_gland_list_template
                    new_sh["L3"].options(transpose=True).value = cub_unique_remove_equal


                    offset = 2
                    gland_formulas=[]


                    for i in cub_unique_remove_equal:
                        a = [
                            f"=COUNTIFS($D$3:$D${len(df_cable_list)+offset},$M$2,$E$3:$E${len(df_cable_list)+offset},\"*{i}\") + COUNTIFS($D$3:$D${len(df_cable_list)+offset},$M$2,$F$3:$F${len(df_cable_list)+offset},\"*{i}\")",
                            f"=COUNTIFS($D$3:$D${len(df_cable_list)+offset},$N$2,$E$3:$E${len(df_cable_list)+offset},\"*{i}\") + COUNTIFS($D$3:$D${len(df_cable_list)+offset},$N$2,$F$3:$F${len(df_cable_list)+offset},\"*{i}\")",
                            f"=COUNTIFS($D$3:$D${len(df_cable_list)+offset},$O$2,$E$3:$E${len(df_cable_list)+offset},\"*{i}\") + COUNTIFS($D$3:$D${len(df_cable_list)+offset},$O$2,$F$3:$F${len(df_cable_list)+offset},\"*{i}\")",
                        ]

                        gland_formulas.append(a)

                    new_sh["M3"].value = gland_formulas

                    # Định dạng lại các ô dữ liệu
                    sh_result[f'A1:H{len(df_cub_s)+offset}'].api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
                    sh_result[f'A1:H{len(df_cub_s)+offset}'].api.VerticalAlignment = constants.VAlign.xlVAlignCenter

                    sh_result[f'A1:H{len(df_cub_s)+offset}'].font.size = 11
                    sh_result[f'A3:H{len(df_cub_s)+offset}'].font.bold = False
                    sh_result[f'A1:H2'].font.bold = True
                    sh_result[f'A1:H{len(df_cub_s)+offset}'].font.name = "Time New Roman"

                    new_sh[f'A1:O{len(df_cable_list)+offset}'].api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
                    new_sh[f'A1:O{len(df_cable_list)+offset}'].api.VerticalAlignment = constants.VAlign.xlVAlignCenter

                    new_sh[f'A1:O{len(df_cable_list)+offset}'].font.size = 11
                    new_sh[f'A3:O{len(df_cable_list)+offset}'].font.bold = False
                    new_sh[f'A1:H2'].font.bold = True
                    new_sh[f'A1:O{len(df_cable_list)+offset}'].font.name = "Time New Roman"


                    

                    # Ép các cột trong sheet mới giãn ô tự động
                    sh_result["A:H"].autofit()
                    new_sh["A:O"].autofit()

                    # Vẽ khung cho bảng dữ liệu excel
                    for i in range(7,13):
                        sh_result[f'A2:H{len(df_cub_s)+offset}'].api.Borders(i).LineStyle = 1
                        sh_result[f'A2:H{len(df_cub_s)+offset}'].api.Borders(i).Weight = 2
                        new_sh[f'A2:H{len(df_cable_list)+offset}'].api.Borders(i).LineStyle = 1
                        new_sh[f'A2:H{len(df_cable_list)+offset}'].api.Borders(i).Weight = 2

                    wb_result.save()

                    mg.showinfo("Thông báo","Đã xong!!!") 
            else:
                mg.showinfo("Thông báo","Hãy chọn các file Connection và form xuất cáp trước") 

root = exportCables()
root.mainloop()

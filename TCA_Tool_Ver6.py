import tkinter as tk
from tkinter import ttk, filedialog as fd, messagebox as mg
import xlwings as xw
import pandas as pd
import numpy as np
from xlwings import constants
from xlwings.constants import DeleteShiftDirection
import sys
import datetime

class tcaTool(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Tool TCA Rev.6")
        self.rowconfigure((0,1), weight=1)
        self.columnconfigure(0, weight=1)

        labelframe1 = tk.LabelFrame(self, text="Tool Xuất nhà cung ứng, kiểm tra tồn kho, tính toán sản xuất")
        labelframe1.grid(sticky="NWES", column=0, row=0, padx=(5,5), pady=(5,5))

        data_frame = ttk.Frame(labelframe1, padding=(10,10))
        data_frame.grid(sticky="NWES", column=0, row=0)
        data_frame.columnconfigure(0, weight=2)
        data_frame.columnconfigure(1, weight=1)
        data_frame.columnconfigure(2, weight=4)
        data_frame.rowconfigure((0,1,2,3), weight=1)

        button_frame = ttk.Frame(labelframe1, padding=(10,10))
        button_frame.grid(sticky="NWES", column=0, row=1)
        button_frame.columnconfigure((0,1,2,3,4,5), weight=1)

        self.data_file_text = tk.StringVar()
        self.database_file_text = tk.StringVar()
        self.inventory_file_text = tk.StringVar()
        self.manufacture_file_text = tk.StringVar()

        data_label = ttk.Label(data_frame, text="File Dự trù")
        data_button = ttk.Button(data_frame, text="Chọn File", command=self.data_choose_file)
        data_entry = ttk.Entry(data_frame, width=50, textvariable=self.data_file_text, state="disable")

        database_label = ttk.Label(data_frame, text="File Thông tin VTTB")
        database_button = ttk.Button(data_frame, text="Chọn File", command=self.database_choose_file)
        database_entry = ttk.Entry(data_frame, width=50, textvariable=self.database_file_text, state="disable")

        inventory_label = ttk.Label(data_frame, text="File Tồn kho")
        inventory_button = ttk.Button(data_frame, text="Chọn File", command=self.inventory_choose_file)
        inventory_entry = ttk.Entry(data_frame, width=50, textvariable=self.inventory_file_text, state="disable")

        manufacture_label = ttk.Label(data_frame, text="File QLSX")
        manufacture_button = ttk.Button(data_frame, text="Chọn File", command=self.manufacture_choose_file)
        manufacture_entry = ttk.Entry(data_frame, width=50, textvariable=self.manufacture_file_text, state="disable")

        clear_button = ttk.Button(button_frame, text="Xóa sheet", command=self.clear)
        findERP_button = ttk.Button(button_frame, text="Tìm ERP Code", command=self.findERP)
        purchase_button = ttk.Button(button_frame, text="Xuất nhà cung ứng", command=self.purchase)
        check_inventory_button = ttk.Button(button_frame, text="Kiểm tra tồn kho", command=self.check_inventory)
        show_buying_item_button = ttk.Button(button_frame, text="Tạo sheet mua", command=self.generate_buying_sheet)
        calculate_manufacture_button = ttk.Button(button_frame, text="Tính toán SX", command=self.calculate_manufacture)

        data_label.grid(sticky="EW", column=0, row= 0, padx=5, pady=5)
        data_button.grid(sticky="EW", column=1, row= 0, padx=5, pady=5)
        data_entry.grid(sticky="EW", column=2, row= 0, padx=5, pady=5)

        database_label.grid(sticky="EW", column=0, row= 1, padx=5, pady=5)
        database_button.grid(sticky="EW", column=1, row= 1, padx=5, pady=5)
        database_entry.grid(sticky="EW", column=2, row= 1, padx=5, pady=5)

        inventory_label.grid(sticky="EW", column=0, row= 2, padx=5, pady=5)
        inventory_button.grid(sticky="EW", column=1, row= 2, padx=5, pady=5)
        inventory_entry.grid(sticky="EW", column=2, row= 2, padx=5, pady=5)

        manufacture_label.grid(sticky="EW", column=0, row= 3, padx=5, pady=5)
        manufacture_button.grid(sticky="EW", column=1, row= 3, padx=5, pady=5)
        manufacture_entry.grid(sticky="EW", column=2, row= 3, padx=5, pady=5)

        clear_button.grid(sticky="EW", column=0, row= 0, padx=5, pady=5)
        findERP_button.grid(sticky="EW", column=1, row= 0, padx=5, pady=5)
        purchase_button.grid(sticky="EW", column=2, row= 0, padx=5, pady=5)
        check_inventory_button.grid(sticky="EW", column=3, row= 0, padx=5, pady=5)
        show_buying_item_button.grid(sticky="EW", column=4, row= 0, padx=5, pady=5)
        calculate_manufacture_button.grid(sticky="EW", column=5, row= 0, padx=5, pady=5)


        labelframe2 = tk.LabelFrame(self, text="Tool tổng hợp các file dự trù xuất giao hàng")
        labelframe2.grid(sticky="NWES", column=0, row=1, padx=(5,5), pady=(5,5))

        self.data_file_text_2 = tk.StringVar()
        self.result_file_text_2 = tk.StringVar()

        data_frame_2 = ttk.Frame(labelframe2, padding=(10,10))
        data_frame_2.grid(sticky="NWES", column=0, row=0)
        data_frame_2.columnconfigure(0, weight=2)
        data_frame_2.columnconfigure(1, weight=1)
        data_frame_2.columnconfigure(2, weight=4)
        data_frame_2.rowconfigure((0,1), weight=1)

        button_frame_2 = ttk.Frame(labelframe2, padding=(10,10))
        button_frame_2.grid(sticky="NWES", column=0, row=1)
        button_frame_2.columnconfigure(0, weight=1)

        data_label_2 = ttk.Label(data_frame_2, text="Các File Dự trù")
        data_button_2 = ttk.Button(data_frame_2, text="Chọn Files", command=self.data_choose_multiple_file)
        data_entry_2 = ttk.Entry(data_frame_2, width=50, textvariable=self.data_file_text_2, state="disable")

        result_label_2 = ttk.Label(data_frame_2, text="Form Dự trù tổng")
        result_button_2 = ttk.Button(data_frame_2, text="Chọn File", command=self.result_choose_file)
        result_entry_2 = ttk.Entry(data_frame_2, width=50, textvariable=self.result_file_text_2, state="disable")

        export_button_2 = ttk.Button(button_frame_2, text="Xuất file tổng hợp", command=self.exportCollection)

        data_label_2.grid(sticky="EW", column=0, row= 0, padx=5, pady=5)
        data_button_2.grid(sticky="EW", column=1, row= 0, padx=5, pady=5)
        data_entry_2.grid(sticky="EW", column=2, row= 0, padx=5, pady=5)

        result_label_2.grid(sticky="EW", column=0, row= 1, padx=5, pady=5)
        result_button_2.grid(sticky="EW", column=1, row= 1, padx=5, pady=5)
        result_entry_2.grid(sticky="EW", column=2, row= 1, padx=5, pady=5)

        export_button_2.grid(sticky="NSEW", column=0, row= 0, padx=5, pady=5)


        excel_file_extensions = ['*.xls', '*.xlsx', '*.xlsb',  '*.xlsm']
        self.ftypes = [    
            ('test files', excel_file_extensions), 
            ('All files', '*'), 
        ]

    def data_choose_file(self):
        filename =  fd.askopenfilename(initialdir = "/a",title = "Hãy chọn file Dự trù",filetypes=self.ftypes)
        self.data_file_text.set(filename)
   
    def database_choose_file(self):
        filename =  fd.askopenfilename(initialdir = "/a",title = "Hãy chọn file Database",filetypes=self.ftypes)
        self.database_file_text.set(filename)
    
    def inventory_choose_file(self):
        filename =  fd.askopenfilename(initialdir = "/a",title = "Hãy chọn file Tồn kho",filetypes=self.ftypes)
        self.inventory_file_text.set(filename)

    def manufacture_choose_file(self):
        filename =  fd.askopenfilename(initialdir = "/a",title = "Hãy chọn file QLSX",filetypes=self.ftypes)
        self.manufacture_file_text.set(filename)    

    def data_choose_multiple_file(self):
        self.mulfilename =  fd.askopenfilenames(initialdir = "/a",title = "Hãy chọn các file Dự trù",filetypes=self.ftypes)
        self.data_file_text_2.set(self.mulfilename)
    
    def result_choose_file(self):
        filename =  fd.askopenfilename(initialdir = "/a",title = "Hãy chọn form Dự trù",filetypes=self.ftypes)
        self.result_file_text_2.set(filename)

    def findERP(self):
        if self.database_file_text.get() and self.data_file_text.get():
            try:
                # Đọc file excel
                wb = xw.Book(self.data_file_text.get())
                wb_erp = xw.Book(self.database_file_text.get())

                # Tạo biến sh_main lưu giữ giá trị của tên Dự trù và biến sh_main_erp lưu giữ giá trị tên sheet đầu tiên của database
                sh_main = wb.sheets[0]
                sh_main_erp = wb_erp.sheets[0]

                # Find the last row
                last_row = sh_main.range(f'B{sh_main.cells.last_cell.row}').end('up').row
                last_row_erp = sh_main_erp.range(f'B{sh_main_erp.cells.last_cell.row}').end('up').row

                # Tạo Dataframe lấy tất cả dữ liệu trong bảng Dự trù
                df = pd.DataFrame(sh_main[f'D14:D{last_row}'].value, columns=["order"])
                df_erp = pd.DataFrame(
                    sh_main_erp[f'B3:I{last_row_erp}'].value, 
                    columns=["order","des_vi","des_en","man_vi", "man_en","unit_vi","unit_en", "erp"]
                    )
                
                # Đóng file database
                wb_erp.close()

                # Tạo dataframe merge giữa 2 Dataframe
                result = pd.merge(df, df_erp, how="left")

                # Lấy giá trị trong bảng erp để đưa vào cột ERP trong file dự trù
                sh_main["H14"].options(transpose=True).value = result["erp"].values.tolist()

                mg.showinfo("Thông báo", "Đã hoàn thành!")

            except Exception as e:
                _, _, error_tb = sys.exc_info()
                error_str = self.findERP.__qualname__ + ": " + str(e) + ", Line error: " + str(error_tb.tb_lineno)
                mg.showerror("Error", error_str)
        else:
            mg.showinfo("Thông báo","Hãy chọn file Dự trù và file Database trước") 

    def purchase(self):
        if self.data_file_text.get():
            try:
                wb = xw.Book(self.data_file_text.get())

                # Xóa các trang sheet khác sheet Dự trù
                self.delete_data_sheet(workbook=wb)

                # Tạo biến sh_main lưu giữ giá trị của tên Dự trù
                sh_main = wb.sheets[0]

                # Tìm dòng cuối theo cột B trong sheet Dự trù
                last_row = sh_main.range(f'B{sh_main.cells.last_cell.row}').end('up').row
                
                # Tạo Dataframe lấy tất cả dữ liệu trong bảng Dự trù
                df = pd.DataFrame(sh_main[f'B14:H{last_row}'].value, columns=["des","man","order","unit", "quan","note","erp"])

                # Lên danh sách tất cả các nhà cung cấp có trong sheet Dự Trù
                man_list = [i for i in df.groupby("man").groups]

                # Sắp xếp các tên nhà cung cấp theo danh sách giảm dần
                man_list.sort(reverse=True)

                # Thực hiện vòng lặp cho từng nhà cung cấp
                for man in man_list:
                    
                    # Tạo Dataframe cho riêng từng nhà cung cấp
                    dfa = df.groupby("man").get_group(man)

                    # Tạo phương thức kiểm tra xem chỗ unit có phải tủ hay không?
                    units = dfa["unit"].unique()

                    unit_is_cuble = False

                    for u in units:
                        if str(u).lower().__contains__("tủ"):
                            unit_is_cuble = True
                            break

                    # Sử dụng tính năng Dataframe để tính tổng số lượng các mã Ordercode 
                    dfb = dfa.groupby("order").agg(quan=("quan","sum"), erp=("erp","first"), des = ("des", "first"), unit = ("unit", "first"), note = ("note", "first")).reset_index()

                    if unit_is_cuble:
                        dfb = dfb.dropna(subset=["note"])
                        dfb = dfb[~dfb["note"].str.strip().isin(["", "None"])]

                    print(dfb)
                    
                    if len(dfb):

                        # Tạo thêm cột Manufacturer
                        dfb["man"] = man
                        
                        # Tạo thêm cột Số thứ tự
                        dfb["no"] = np.arange(len(dfb.index)) + 1

                        # Tạo thêm cột Note 1: Cột này sẽ là ghi chú từ bên hàm check tồn kho
                        dfb["note1"] = np.nan

                        # Tạo thêm cột Inventory: 
                        dfb["inventory"] = np.nan

                        # Tạo thêm cột Inventory: 
                        dfb["buying"] = np.nan

                        # Sắp xếp lại các cột trong Dataframe
                        dfb = dfb[["no","des","man","order","unit","quan","erp","note1","inventory","buying","note"]]

                        # Thêm ký tự "'" vào trước dữ liệu trong cột order
                        dfb['order'] = "'" + dfb['order'].astype(str)

                        # Tạo biến tiêu đề cho từng sheet mới
                        title = ['No','Description','Manufacturer','Order Code','Unit','Quantity','ERP Code', 'Note1', 'Inventory', 'Buying','Note2']
                        
                        # Tạo sheet mới ngay sau sheet main Dự trù
                        man_name = man.replace("/","_")
                        new_sh = wb.sheets.add(name=man_name,after=sh_main.name) 
                        
                        # Tạo các tiêu đề cho sheet mới từ dòng 1
                        new_sh['A1'].value = title
                        
                        # Lấy toàn bộ dữ liệu của Dataframe cho vào sheet mới
                        new_sh["A2"].value = dfb.values.tolist()
                        
                        # Định dạng dòng đầu tiên bôi đậm
                        new_sh["A1:K1"].api.Font.Bold = True
                        
                        # Định dạng các ô dữ liệu trong sheet mới phải được căn lề giữa
                        new_sh["A:K"].api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
                        new_sh["A:K"].api.VerticalAlignment = constants.VAlign.xlVAlignCenter
                        
                        # Ép các cột trong sheet mới giãn ô tự động
                        new_sh["A:K"].autofit()
                        
                        """
                        # Borders(id) with id:
                                                7: Left Border
                                                8: Top Border
                                                9: Bottom Border
                                                10: Right Border
                                                11: Inner Vertical Border
                                                12: Inner Horizontal Border
                        Border.LineStyle = id
                                                1: Straight line
                                                2: Dotted Line
                                                3: Double-dot Chain Line
                                                4: Dot-dash Line
                                                5: Dot-Dot-Dash Line
                        """
                        # Vẽ kẻ khung cho từng sheet mới
                        for i in range(7,13):
                            new_sh[f'A1:K{len(dfb.index)+1}'].api.Borders(i).LineStyle = 1
                            new_sh[f'A1:K{len(dfb.index)+1}'].api.Borders(i).Weight = 2
                mg.showinfo("Thông báo", "Đã hoàn thành!")
                
            except Exception as e:
                _, _, error_tb = sys.exc_info()
                error_str = self.purchase.__qualname__ + ": " + str(e) + ", Line error: " + str(error_tb.tb_lineno)
                mg.showerror("Error", error_str)
        else:
            mg.showinfo("Thông báo","Hãy chọn file Dự trù trước") 
    
    def check_inventory(self):
        if self.inventory_file_text.get() and self.data_file_text.get():
            try:
                wb = xw.Book(self.data_file_text.get())
                wb_inv = xw.Book(self.inventory_file_text.get())

                # Định danh cho biến sh_main bằng sheet đầu tiên của file Dự trù
                sh_main = wb.sheets[0]

                # Định danh cho biến sh_main_inv bằng sheet đầu tiền của file Tồn kho
                sh_main_inv = wb_inv.sheets[0]

                # Tìm dòng cuối dữ liệu trong file Tồn kho
                last_row_inv = sh_main_inv.range(f'B{sh_main_inv.cells.last_cell.row}').end('up').row

                # Tạo Dataframe lấy tất cả dữ liệu trong file tồn kho
                df_inv = pd.DataFrame(
                    sh_main_inv[f'B6:K{last_row_inv}'].value,
                    columns=["erp","C","D","E","F","G","H","I","J","quan_inv"]
                    )

                # Lấy Dataframe cho dữ liệu ERP và khối lượng trong kho
                df_inv = df_inv[["erp","quan_inv"]]

                # Đóng workbook Tồn kho
                wb_inv.close()

                # Dùng vòng lặp cho từng sheet khác với sheet dự trù
                for sh in wb.sheets:
                    if sh != sh_main:

                        # Tìm dòng cuối cùng của sheet
                        last_row = sh.range(f'B{sh.cells.last_cell.row}').end('up').row
                        found_one = False
                        if last_row == 2:
                            found_one = True
                            last_row += 1

                        # Lấy Dataframe cho dữ liệu ERP và khối lượng hiện hữu
                        df = pd.DataFrame(sh[f'F2:G{last_row}'].value, columns=["quan","erp"])

                        # Đổi vị trí cột
                        df = df[["erp","quan"]]

                        # Đổi các biến nan thành số 0 và chuyển định dạng số ERP sang int để loại bỏ phẩy động
                        df['erp'] = df['erp'].fillna(0).astype(int)

                        # Đổi định dạng của tất cả các biến trong cột ERP sang dạng string
                        df['erp'] = df['erp'].astype(str)

                        # Tạo Dataframme để merge 2 Dataframe lại với nhau
                        result = pd.merge(df, df_inv, on="erp", how="left")

                        # Đổi tất cả các biến trong cột khối lượng trong tồn kho nếu là NAN thì chuyển sang 0
                        result["quan_inv"] = result["quan_inv"].fillna(0)

                        # Tạo lệnh điều kiện cho cột Note
                        conditions_note = [
                            (result["quan"] == 0),
                            (result["quan_inv"] == 0),
                            (result["quan"]-result["quan_inv"] <= 0),
                            (result["quan"]-result["quan_inv"] > 0)
                            ]

                        # Tạo giá trị tương ứng với từng lệnh điều kiện của cột Note
                        values_note = ["Không có khối lượng","Không có tồn kho","Lấy đủ trong kho", "Lấy một phần trong kho" ]

                        # Tạo cột note trong Dataframe result
                        result['note'] = np.select(conditions_note, values_note)
                        
                        # Tạo lệnh điều kiện cho cột Inventory
                        conditions_inv = [
                            (result["quan"]-result["quan_inv"] <= 0),
                            (result["quan"]-result["quan_inv"] > 0),
                            ]

                        # Tạo giá trị tương ứng với từng lệnh điều kiện của cột Inventory
                        values_inv = [result["quan"], result["quan_inv"]]

                        # Tạo cột inventory trong Dataframe result
                        result['inventory'] = np.select(conditions_inv, values_inv)

                        # Tạo lệnh điều kiện cho cột Buying
                        conditions_buy = [
                            (result["quan"]-result["quan_inv"] <= 0),
                            (result["quan"]-result["quan_inv"] > 0),
                            (result["inventory"] == 0)
                            ]

                        # Tạo giá trị tương ứng với từng lệnh điều kiện của cột Buying
                        values_buy = [0, result["quan"]-result["quan_inv"],result["quan"]]

                        # Tạo cột buying trong Dataframe result
                        result['buying'] = np.select(conditions_buy, values_buy)

                        # Lấy 3 cột note, inventory và buying để gửi dữ liệu vào excel
                        result = result[["note", "inventory", "buying"]]
                        
                        # Gửi dữ liệu vào excel
                        sh["H2"].value = result.values.tolist()

                        # Tự động giãn cột
                        sh["A:J"].autofit()

                        for i in range(2, last_row+1):
                            if sh[f'H{i}'].value == "Lấy đủ trong kho":
                                sh[f'H{i}'].color = (0, 204, 0) # Màu xanh
                            elif sh[f'H{i}'].value == "Lấy một phần trong kho":
                                sh[f'H{i}'].color = (255, 255, 0) # Màu vàng
                            else:
                                sh[f'H{i}'].color = (255, 0, 0) # Màu đỏ
                        
                        # Xóa dòng cuối nếu chỉ trong bảng chỉ có 1 thiết bị
                        if found_one:
                            sh[f'{last_row}:{last_row}'].api.Delete(DeleteShiftDirection.xlShiftUp)

                mg.showinfo("Thông báo", "Đã hoàn thành!")
            except Exception as e:
                _, _, error_tb = sys.exc_info()
                error_str = self.check_inventory.__qualname__ + ": " + str(e) + ", Line error: " + str(error_tb.tb_lineno)
                mg.showerror("Error", error_str)
        else:
            mg.showinfo("Thông báo","Hãy chọn file Dự trù và file Tồn kho trước")   

    def exportCollection(self):
        # Kiểm tra xem đã có chọn file hay chưa
        if self.data_file_text_2.get() and self.result_file_text_2.get():
            # Tạo trước 1 dataframe rỗng nhằm tổng hợp tất cả các dataframe đọc được từ các file dự trù
            df_combine =pd.DataFrame(columns=["des","man","order","unit", "quan","note","erp"])

            # Sử dụng câu lệnh with này để tránh luôn chạy tiến trình chạy Excel trong task manager
            with xw.App(visible=False) as app:
                
                try:
                # Tạo vòng lặp cho từng file được chọn
                    for file in self.mulfilename:
                        # Kết nối mở file dự trù
                        wb = xw.Book(file)

                        # Kết nối đến sheet đầu tiên của file dự trù
                        sh = wb.sheets[0]

                        # Tìm dòng cuối của dữ liệu trong file dự trù
                        last_row = sh.range(f'B{sh.cells.last_cell.row}').end('up').row

                        # Lấy tất cả các dữ liệu trong file dự trù vào dataframe
                        df = pd.DataFrame(sh[f'B14:H{last_row}'].value, columns=["des","man","order","unit", "quan","note","erp"])

                        # Nối dataframe của file dự trù vào dataframe combine
                        # df_combine = df_combine.append(df) --> Deprecated
                        df_combine = pd.concat([df_combine, df], ignore_index=True)

                        # Save the workbook
                        # wb.save()      
                
                    # Đặt lại địa chỉ index của dataframe combine tính từ 0    
                    df_combine.index = np.arange(0, len(df_combine))  

                    # Lọc tìm tất cả các tên tủ có trong dataframe combine (bắt đầu bằng dấu =) đưa vào trong 1 dataframe cubicles
                    df_cubicles = df_combine[df_combine['order'].str.contains("=", na = False)]['order'] 

                    # Tạo các list: Các tủ duy nhất, địa chỉ dòng của các tủ có trong dataframe combine, tên của các tủ có trong dataframe cubicles  
                    unique_cubicles, index_cubicles, name_cubicles = df_cubicles.unique(), df_cubicles.index.tolist(), df_cubicles.tolist()

                    # Sort lại tên tủ 
                    unique_cubicles.sort()
                    
                    # Tạo ra một dataframe rỗng là dataframe kết quả cuối cùng
                    df_result =pd.DataFrame(columns=["des","man","order","unit", "quan","note","erp"])

                    # Chạy vòng lặp cho từng tủ
                    for cubicle in unique_cubicles:
                        # Tạo 1 dataframe cho từng tủ
                        df_each_cubicle_with_header =pd.DataFrame(columns=["des","man","order","unit", "quan","note","erp"])

                        # Lấy thông tin dòng tên tủ
                        header = df_combine[df_combine['order'] == cubicle].iloc[0].to_dict()
                        header_list = []
                        header_list.append(header)
                        df_cubicle_header = pd.DataFrame.from_records(header_list)

                        # Tạo thêm 1 dataframe của tủ mà có chứa dòng có tên tủ
                        # df_each_cubicle_with_header = df_each_cubicle_with_header.append(df_combine[df_combine['order']==cubicle].iloc[0])
                        df_each_cubicle_with_header = pd.concat([df_each_cubicle_with_header,df_cubicle_header], ignore_index=True)

                        # Đổi dấu = thành dấu '= để khi ghi vào file excel không bị báo lỗi do excel bị nhầm tưởng dấu = mang ý nghĩa của 1 hàm toán học
                        df_each_cubicle_with_header["order"] = df_each_cubicle_with_header["order"].str.replace("=","'=")

                        # Đặt index cho tên tủ luôn = 0
                        df_each_cubicle_with_header.index = np.arange(0, len(df_each_cubicle_with_header))

                        # Tạo thêm 1 dataframe của tủ nhưng lại không chứa dòng có tên tủ
                        df_each_cubicle_without_header =pd.DataFrame(columns=["des","man","order","unit", "quan","note","erp"])

                        # Bắt đầu quá trình sử dụng vòng lặp để tìm dữ liệu cho tủ
                        for i in range(len(name_cubicles)):
                            # Nếu tên tủ trong dataframe cubicles trùng với tên tủ đang chạy
                            if name_cubicles[i] == cubicle:
                                # Chỉ số bắt đầu
                                index_start = index_cubicles[i]+1

                                # Chỉ số kết thúc
                                if i < len(name_cubicles)-1:
                                    index_end = index_cubicles[i+1]
                                else:
                                    index_end = len(df_combine)-1

                                # Sử dụng chỉ số bắt đầu và kết thúc để lấy tất cả các dòng thuộc cùng 1 tủ, đưa vào dataframe
                                df_temp = df_combine[index_start:index_end]

                                # Thêm các dataframe tạm thời này vào dataframe không có header
                                # df_each_cubicle_without_header = df_each_cubicle_without_header.append(df_temp)
                                df_each_cubicle_without_header = pd.concat([df_each_cubicle_without_header, df_temp], ignore_index=True)
                        
                        # Sử dụng groupby để tính tổng số lượng các thiết bị trong một tủ
                        df_each_cubicle_without_header = df_each_cubicle_without_header.groupby('order').agg(des = ("des", "min"),man = ("man", "min"),order = ("order", "min"),  unit = ("unit", "min"),quan=("quan","sum"))

                        # Sort theo tên nhà cung cấp
                        df_each_cubicle_without_header = df_each_cubicle_without_header.sort_values('man')  

                        # Đặt lại địa chỉ index của dataframe tính từ 1, tên tủ sẽ có địa chỉ dòng là 0
                        df_each_cubicle_without_header.index = np.arange(1, len(df_each_cubicle_without_header)+1)

                        # Nối dataframe không có header vào dataframe có header để giữ cấu trúc: tên tủ ở dòng đầu tien, sau đó là danh sách các thiết bị trong tủ ở phía dưới
                        # df_each_cubicle_with_header =  df_each_cubicle_with_header.append(df_each_cubicle_without_header)
                        df_each_cubicle_with_header = pd.concat([df_each_cubicle_with_header, df_each_cubicle_without_header], ignore_index=True)

                        # Nối các dataframe có header vào dataframe kết quả cuối cùng. Đây chính là dataframe có thể in ra excel
                        # df_result = df_result.append(df_each_cubicle_with_header)
                        df_result = pd.concat([df_result,df_each_cubicle_with_header], ignore_index=True)
                    
                    # Tạo thêm cột Số thứ tự trong dataframe result
                    df_result['no'] = df_result.index.tolist()

                    # Điều chỉnh lại vị trí các cột trong dataframe theo đúng form
                    df_result = df_result[["no","des","man","order","unit", "quan","note", "erp"]]

                    # Mở file form Dự trù
                    wb = xw.Book(self.result_file_text_2.get())

                    # Kết nối đến sheet đầu tiên
                    sh = wb.sheets[0]

                    #==================================================================================================
                    # Thêm yêu cầu mới
                    # Xóa các sheet khác nếu có
                    for sh_temp in wb.sheets:
                        if sh_temp != sh:
                            sh_temp.delete()

                    # Tạo sheet mới ngay sau sheet result
                    new_sh = wb.sheets.add(name="summarized list",after=sh.name) 

                    df_product_list = df_result[~df_result["order"].str.contains("=", na = False)]

                    df_product_list = df_product_list.groupby('order').agg(no = ("no", "min") ,des = ("des", "min"),man = ("man", "min"),order = ("order", "min"),  unit = ("unit", "min"),quan=("quan","sum"))
                    
                    df_product_list = df_product_list[["no","des","man","order","unit", "quan"]]

                    df_product_list = df_product_list.sort_values('man') 

                    df_product_list["no"] = list(np.arange(1, len(df_product_list)+1))

                    new_sh["A1"].value = ["STT", "Mô Tả", "Nhà Sản Xuất", "Mã Hiệu", "Đơn vị", "Số Lượng"]
                    
                    new_sh["A2"].value=df_product_list.values.tolist()

                    # Định dạng lại các ô dữ liệu
                    new_sh[f'A1:F{len(df_product_list)+1}'].api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
                    new_sh[f'A14:F{len(df_product_list)+1}'].api.VerticalAlignment = constants.VAlign.xlVAlignCenter

                    new_sh[f'A1:F{len(df_product_list)+1}'].font.size = 11
                    new_sh[f'A1:F{len(df_product_list)+1}'].font.bold = False
                    new_sh[f'A1:F{len(df_product_list)+1}'].font.name = "Time New Roman"

                    new_sh["A:F"].autofit()

                    new_sh['A1:F1'].font.bold = True

                    # Vẽ khung cho bảng dữ liệu excel
                    for i in range(7,13):
                        new_sh[f'A1:F{len(df_product_list)+1}'].api.Borders(i).LineStyle = 1
                        new_sh[f'A1:F{len(df_product_list)+1}'].api.Borders(i).Weight = 2

                    #==================================================================================================

                    # Tìm dòng cuối của dữ liệu trong file dự trù
                    last_row = sh.range(f'B{sh.cells.last_cell.row}').end('up').row

                    # Xóa hết nội dung trong sheet đầu tiên
                    if last_row >= 14:
                        sh.range(f'14:{last_row}').api.Delete(DeleteShiftDirection.xlShiftUp) 

                    # Điền giá trị từ dataframe result vào cell A14
                    sh["A14"].value=df_result.values.tolist()

                    offset = 13
                    
                    # Định dạng lại các ô dữ liệu
                    sh[f'A14:G{len(df_result)+offset}'].api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
                    sh[f'A14:G{len(df_result)+offset}'].api.VerticalAlignment = constants.VAlign.xlVAlignCenter

                    sh[f'A14:G{len(df_result)+offset}'].font.size = 11
                    sh[f'A14:G{len(df_result)+offset}'].font.bold = False
                    sh[f'A14:G{len(df_result)+offset}'].font.name = "Time New Roman"

                    # Ép các cột trong sheet mới giãn ô tự động
                    sh["A:G"].autofit()
                    
                    # Vẽ khung cho bảng dữ liệu excel
                    for i in range(7,13):
                        sh[f'A14:G{len(df_result)+offset}'].api.Borders(i).LineStyle = 1
                        sh[f'A14:G{len(df_result)+offset}'].api.Borders(i).Weight = 2

                    # Đánh lại địa chỉ dòng cho df_result
                    df_result.index = np.arange(0, len(df_result))

                    # Tìm lại địa chỉ các dòng của tủ
                    bold_list = df_result[df_result['order'].str.contains("=", na = False)]['order'].index.tolist()

                    # Bôi đậm và tô màu cho các dòng của tủ
                    for i in bold_list:
                        sh[f'A{i+offset + 1}:G{i+offset + 1}'].font.bold = True
                        sh[f'A{i+offset + 1}:G{i+offset + 1}'].color = (0, 204, 0)
                    
                    # Save
                    wb.save() 

                    # Đóng file
                    wb.close()  

                    # Sau khi hoàn thiện thì tạo thông báo
                    mg.showinfo("Thông báo",f"Đã hoàn thành!!!.\nKiểm tra file tại đường dẫn:\n{self.result_file_text_2.get()}")   

                except Exception as e:
                    _, _, error_tb = sys.exc_info()
                    error_str = self.exportCollection.__qualname__ + ": " + str(e) + ", Line error: " + str(error_tb.tb_lineno)
                    mg.showerror("Error", error_str) 
        # Nếu chưa chọn thì xuất hiện thông báo
        else:
            mg.showinfo("Thông báo","Hãy chọn các file Dự trù và form Dự trù trước")  

    def generate_buying_sheet(self):
        if self.data_file_text.get():
            try:
                wb = xw.Book(self.data_file_text.get())

                # Đặt sh_main là sheet đầu tiên: Sheet Dự trù
                sh_main = wb.sheets[0]

                # Xóa sheet buying_sheet nếu đã tồn tại
                if wb.sheets.count > 2:
                    if wb.sheets[1].name == "buying_sheet":
                        wb.sheets[1].delete() 

                    # Tạo sheet mới ngay sau sheet main Dự trù
                    new_sh = wb.sheets.add(name="buying_sheet",after=sh_main.name) 

                    # Tạo biến tiêu đề cho từng sheet mới
                    title = ['No','Description','Manufacturer','Order Code','Unit','Quantity','ERP Code', 'Note', 'Inventory', 'Buying']

                    # Tạo Dataframe tổng
                    df_combine =pd.DataFrame(columns=title)

                    # Chạy qua tất cả các sheet
                    for sh in wb.sheets:
                        if sh != wb.sheets[0] and sh != wb.sheets[1]:
                            # Tìm dòng cuối cùng của sheet
                            last_row = sh.range(f'B{sh.cells.last_cell.row}').end('up').row
                            found_one = False
                            if last_row == 2:
                                found_one = True
                                last_row += 1
                            
                            # Lấy Dataframe cho dữ liệu ERP và khối lượng hiện hữu
                            df = pd.DataFrame(sh[f'A2:J{last_row}'].value, columns=title)

                            # Lọc các dữ liệu với Buying != 0
                            df_buying = df[df['Buying'] > 0]

                            # Ghép các dữ liệu này vào Dataframe combine
                            if len(df_buying):
                                df_combine = pd.concat([df_combine, df_buying], ignore_index=True)
                                #df_combine = df_combine.append(df_buying)

                            #print(df_combine)
                    
                    # Lọc lại dữ liệu: bỏ qua các thành phần tủ TCA đi
                    df_combine_filter = df_combine[df_combine["Manufacturer"].str.lower().str.contains('tca') == False]

                    # Đánh thứ tự cho cột
                    df_combine_filter["No"] = np.arange(len(df_combine_filter.index)) + 1

                    # Dán dữ liệu vào trang mới
                    new_sh["A1"].value = title
                    new_sh["A2"].value=df_combine_filter.values.tolist()

                    # Định dạng dòng đầu tiên bôi đậm
                    new_sh["A1:j1"].api.Font.Bold = True
                        
                    # Định dạng các ô dữ liệu trong sheet mới phải được căn lề giữa
                    new_sh["A:J"].api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
                    new_sh["A:J"].api.VerticalAlignment = constants.VAlign.xlVAlignCenter

                    # Tự động giãn cột
                    new_sh["A:J"].autofit()

                    # Tìm lại dòng cuối
                    last_row = len(df_combine_filter.index) + 1

                    # Bôi màu
                    for i in range(2, last_row+1):
                        if new_sh[f'H{i}'].value == "Lấy đủ trong kho":
                            new_sh[f'H{i}'].color = (0, 204, 0) # Màu xanh
                        elif new_sh[f'H{i}'].value == "Lấy một phần trong kho":
                            new_sh[f'H{i}'].color = (255, 255, 0) # Màu vàng
                        else:
                            new_sh[f'H{i}'].color = (255, 0, 0) # Màu đỏ

                    # Vẽ kẻ khung cho từng sheet mới
                    for i in range(7,13):
                        new_sh[f'A1:J{last_row}'].api.Borders(i).LineStyle = 1
                        new_sh[f'A1:J{last_row}'].api.Borders(i).Weight = 2

                    mg.showinfo("Thông báo", "Đã hoàn thành!")
                else:
                    mg.showinfo("Thông báo", "Hãy xuất nhà cung ứng trước")
            except Exception as e:
                _, _, error_tb = sys.exc_info()
                error_str = self.generate_buying_sheet.__qualname__ + ": " + str(e) + ", Line error: " + str(error_tb.tb_lineno)
                mg.showerror("Error", error_str)
        else:
            mg.showinfo("Thông báo","Hãy chọn file Dự trù trước") 

    def calculate_manufacture(self):
        if self.manufacture_file_text.get() and self.data_file_text.get():
            try:

                # Mở file Dự trù
                wb = xw.Book(self.data_file_text.get())

                # Tạo biến sh_main lưu giữ giá trị của tên Dự trù
                sh_main = wb.sheets[0]

                # Lấy các thông tin về ngày giờ liên quan tới sản xuất và hợp đồng
                project_name = sh_main.range("B11").value
                
                if str(project_name).strip() == "":
                    mg.showinfo("Thông báo","Hãy đặt tên dự án vào ô B11 trong file Dự trù") 
                    return

                contract_date = sh_main.range("G9").value

                if not isinstance(contract_date, datetime.date):
                    mg.showinfo("Thông báo","Ngày ký hợp đồng ở ô G9 trong file Dự trù không đúng đính dạng") 
                    return

                delivery_date_to_customers = sh_main.range("G10").value

                if not isinstance(delivery_date_to_customers, datetime.date):
                    mg.showinfo("Thông báo","Ngày giao hàng ở ô G10 trong file Dự trù không đúng đính dạng") 
                    return

                if contract_date > delivery_date_to_customers:
                    mg.showinfo("Thông báo","Hãy đảm bảo thời gian ký hợp đồng nhỏ hơn thời gian giao hàng trong file Dự trù") 
                    return
                
                special_delivery_date = sh_main.range("G11").value
                if not isinstance(special_delivery_date, datetime.date):
                    mg.showinfo("Thông báo","Ngày hàng đặc biệt về ở ô G11 trong file Dự trù không đúng đính dạng") 
                    return

                if special_delivery_date > delivery_date_to_customers:
                    mg.showinfo("Thông báo","Hãy đảm bảo thời gian hàng đặc biệt về nhỏ hơn thời gian giao hàng trong file Dự trù") 
                    return

                # Tìm dòng cuối theo cột B trong sheet Dự trù
                last_row = sh_main.range(f'B{sh_main.cells.last_cell.row}').end('up').row
                
                # Tạo Dataframe lấy tất cả dữ liệu trong bảng Dự trù
                df = pd.DataFrame(sh_main[f'B14:H{last_row}'].value, columns=["des","man","order","unit", "quan","note","erp"])

                # Lên danh sách tất cả các nhà cung cấp có trong sheet Dự Trù
                man_list = [i for i in df.groupby("man").groups]

                # Thực hiện vòng lặp cho từng nhà cung cấp
                for man in man_list:
                    
                    # Tạo Dataframe cho riêng từng nhà cung cấp
                    dfa = df.groupby("man").get_group(man)

                    # Tạo phương thức kiểm tra xem chỗ unit có phải tủ hay không?
                    units = dfa["unit"].unique()

                    unit_is_cuble = False

                    for u in units:
                        if str(u).lower().__contains__("tủ"):
                            unit_is_cuble = True
                            break
                    if unit_is_cuble:
                        # Sử dụng tính năng Dataframe để tính tổng số lượng các loại tủ
                        dfb = dfa.groupby("note").agg(quan=("quan","sum")).reset_index()
                        man_list.remove(man)
                        break
                
                print(dfb)

                # Mở file QLSX
                wb_man = xw.Book(self.manufacture_file_text.get())

                # Định danh cho biến sh_data_hanghoa bằng sheet đầu tiền của file QLSX
                sh_data_hanghoa = wb_man.sheets[0]

                # Tìm dòng cuối dữ liệu trong sh_data_hanghoa
                last_row_data_hanghoa = sh_data_hanghoa.range(f'B{sh_data_hanghoa.cells.last_cell.row}').end('up').row

                # Tạo Dataframe lấy tất cả dữ liệu trong sheet hang hoa
                df_data_hanghoa = pd.DataFrame(
                    sh_data_hanghoa[f'B2:D{last_row_data_hanghoa}'].value,
                    columns=["manufacturer","delivery_time","special"]
                    )
                # Trước hết, thay NaN bằng chuỗi rỗng "" để tránh lỗi khi gọi .str
                df_data_hanghoa["special"] = df_data_hanghoa["special"].fillna("")

                # Tạo một Series đã cắt khoảng trắng
                special_stripped = df_data_hanghoa["special"].str.strip()

                special_mfg_list = df_data_hanghoa.loc[
                    special_stripped != "", 
                    "manufacturer"
                ].tolist()

                # Định danh cho biến sh_data_sanxuat bằng sheet thứ hai của file QLSX
                sh_data_sanxuat = wb_man.sheets[1]

                # Tìm dòng cuối dữ liệu trong sh_data_sanxuat
                last_row_data_sanxuat = sh_data_sanxuat.range(f'B{sh_data_sanxuat.cells.last_cell.row}').end('up').row

                # Tạo Dataframe lấy  dữ liệu loai tu trong sheet san xuat
                df_data_loaitu = pd.DataFrame(
                    sh_data_sanxuat[f'C2:D{last_row_data_sanxuat}'].value,
                    columns=["note","manufacturer_time"]
                    )

                # Ghép 2 dataframe 
                df_merged_loaitu = pd.merge(dfb, df_data_loaitu, how='left', on='note')

                df_merged_loaitu['hour_total'] = df_merged_loaitu['quan'] * df_merged_loaitu['manufacturer_time']

                # Tính toán tổng số giờ phải sản xuất cho tất cả các loại tủ
                total_manufacturer_hour = df_merged_loaitu["hour_total"].sum()


                # Kiểm tra trong file excel QLSX đã có sheet tên dự án chưa, nếu có rồi thì xóa
                for sh in wb_man.sheets:
                    if sh.name.lower() == str(project_name).lower():
                        sh.delete()

                # Tạo sheet project name sau sheet data san xuat:
                new_sh = wb_man.sheets.add(name=project_name,after=sh_data_sanxuat.name) 
                template_title = ["Tên dự án","Ngày ký hợp đồng","Ngày giao hàng", "Tổng số giờ sản xuất", "Ngày bắt đầu sản xuất"]
                template_data = [project_name, contract_date, delivery_date_to_customers, total_manufacturer_hour,f"=C3-C4/'{sh_data_sanxuat.name}'!K2"]

                new_sh.range('A1').options(transpose=True).value = template_title
                new_sh.range('C1').options(transpose=True).value = template_data
                new_sh.range('C2:C3').number_format = "dd-mm-yyyy"
                new_sh.range('C5').number_format = "dd-mm-yyyy"

                purchase_schedule_buying_list = []
                purchase_schedule_order_list = []


                for i in range(0, len(man_list)):

                    purchase_buying_schedule = f"=VLOOKUP(B{i+8},'{sh_data_hanghoa.name}'!$B$2:$D${last_row_data_hanghoa},2,0) * 7"

                    if (man_list[i] not in special_mfg_list):
                        purchase_order_schedule = f"=$C$5-C{i+8}" 
                    else:
                        purchase_order_schedule = f"='[{wb.name}]{sh_main.name}'!$G$11 - C{i+8}"

                    purchase_schedule_buying_list.append(purchase_buying_schedule)

                    purchase_schedule_order_list.append(purchase_order_schedule)

                new_sh.range('B8').options(transpose=True).value = man_list
                new_sh.range('C8').options(transpose=True).value = purchase_schedule_buying_list
                new_sh.range('D8').options(transpose=True).value = purchase_schedule_order_list
                new_sh.range(f"D8:D{len(man_list)+7}").number_format = "dd-mm-yyyy"

                # Gán tiêu đề
                new_sh.range("A6").value = "STT"
                new_sh.range("B6").value = "Hãng sản xuất"
                new_sh.range("C6").value = "Thời gian"
                new_sh.range("D6").value = "Ngày đặt hàng"

                # Gộp ô (merge)
                new_sh.range("A6:A7").api.Merge()
                new_sh.range("B6:B7").api.Merge()
                new_sh.range("C6:C7").api.Merge()
                new_sh.range("D6:D7").api.Merge()

                

                # Sắp xếp dữ liệu theo thứ tự ngày đặt hàng tăng dần
                rng_to_sort = new_sh.range(f"A8:D{len(man_list)+7}")

                rng_to_sort.api.Sort(
                    Key1=new_sh.range("D8").api,
                    Order1=1,                # 1 = xlAscending, 2 = xlDescending
                    Orientation=1,           # 1 = xlTopToBottom
                    Header=0                 # 0 = xlGuess, 1 = xlNo, 2 = xlYes
                )

                weeks = self.generate_weeks(contract_date, delivery_date_to_customers)

                if len(weeks) == 0: return

                # 3) Ghi các tuần (Tuần 1, Tuần 2, ...) ở hàng 7 và ngày bắt đầu ở hàng 8
                #    Bắt đầu từ cột E = 5
                start_col = 5
                for i, (w_num, w_start_date) in enumerate(weeks, start=0):
                    col = start_col + i
                    # Hàng 6: "Tuần x"
                    new_sh.range((6, col)).value = f"Tuần {w_num}"
                    # Hàng 7: ngày bắt đầu
                    new_sh.range((7, col)).value = w_start_date
                    # Định dạng hiển thị ngày
                    new_sh.range((7, col)).number_format = "dd-mm-yyyy"

                # Đánh số thứ tự
                for idx, row in enumerate(range(8, len(man_list)+8), start=1):
                    new_sh.range((row, 1)).value = idx

                # 5) Đọc dữ liệu từng mặt hàng, đánh dấu x
                #    Cột C (thời gian - ngày), Cột D (ngày bắt đầu)
                #    Ghi "x" vào các cột tuần nếu thời gian chồng lấn
                for row in range(8, len(man_list)+8):
                    duration = new_sh.range((row, 3)).value  # cột C
                    start_dt = new_sh.range((row, 4)).value  # cột D

                    # Nếu giá trị trống hoặc không hợp lệ, dừng
                    if not (isinstance(start_dt, datetime.datetime) and isinstance(duration, (int, float))):
                        # Nếu ô trống thì có thể break (nếu bạn muốn dừng ngay) hoặc continue
                        continue

                    # Tính ngày bắt đầu và ngày kết thúc
                    start_date_item = start_dt.date()
                    end_date_item = start_date_item + datetime.timedelta(days=duration - 1)

                    # Lặp qua các tuần (cột E..)
                    for i, (w_num, w_start_date) in enumerate(weeks, start=0):
                        col = start_col + i
                        w_start = w_start_date
                        w_end   = w_start_date + datetime.timedelta(days=6)

                        # Kiểm tra chồng lấn
                        if self.overlap(start_date_item, end_date_item, w_start, w_end):
                            new_sh.range((row, col)).value = "x"
                        else:
                            # Nếu bạn muốn xoá các giá trị cũ
                            new_sh.range((row, col)).value = ""
                            # pass
                

                # Định dạng dòng đầu tiên bôi đậm
                new_sh[f"A6:{self.col_num_to_letter(len(weeks)+4)}7"].api.Font.Bold = True

                new_sh[f"A1:A5"].api.HorizontalAlignment = constants.HAlign.xlHAlignLeft

                # Định dạng các ô dữ liệu trong sheet mới phải được căn lề giữa
                new_sh[f"A:{self.col_num_to_letter(len(weeks)+4)}"].api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
                new_sh[f"A:{self.col_num_to_letter(len(weeks)+4)}"].api.VerticalAlignment = constants.VAlign.xlVAlignCenter
                
                
                """
                # Borders(id) with id:
                                        7: Left Border
                                        8: Top Border
                                        9: Bottom Border
                                        10: Right Border
                                        11: Inner Vertical Border
                                        12: Inner Horizontal Border
                Border.LineStyle = id
                                        1: Straight line
                                        2: Dotted Line
                                        3: Double-dot Chain Line
                                        4: Dot-dash Line
                                        5: Dot-Dot-Dash Line
                """
                # Vẽ kẻ khung cho từng sheet mới
                for i in range(7,13):
                    
                    new_sh[f'A1:C5'].api.Borders(i).LineStyle = 1
                    new_sh[f'A1:C5'].api.Borders(i).Weight = 2

                    new_sh[f'A6:{self.col_num_to_letter(len(weeks) + 4)}{len(man_list) + 7}'].api.Borders(i).LineStyle = 1
                    new_sh[f'A6:{self.col_num_to_letter(len(weeks) + 4)}{len(man_list) + 7}'].api.Borders(i).Weight = 2

                new_sh.range("A1:B1").api.Merge()
                new_sh.range("A2:B2").api.Merge()
                new_sh.range("A3:B3").api.Merge()
                new_sh.range("A4:B4").api.Merge()
                new_sh.range("A5:B5").api.Merge()

                # Ép các cột trong sheet mới giãn ô tự động
                new_sh[f"A:{self.col_num_to_letter(len(weeks)+4)}"].autofit()

                wb_man.save()

            except Exception as e:
                _, _, error_tb = sys.exc_info()
                error_str = self.calculate_manufacture.__qualname__ + ": " + str(e) + ", Line error: " + str(error_tb.tb_lineno)
                mg.showerror("Error", error_str)
        else:
            mg.showinfo("Thông báo","Hãy chọn file Dự trù và file QLSX trước")   

    def clear(self):
        if self.data_file_text.get():
            try:
                wb = xw.Book(self.data_file_text.get())
                self.delete_data_sheet(wb)
            except:
                pass       
        else:
            mg.showinfo("Thông báo","Hãy chọn file Dự trù trước")  

    def delete_data_sheet(self, workbook):
        name_for_main_sheet = ("dự trù", "du tru")
        for sh in workbook.sheets:
            if sh.name.lower() not in name_for_main_sheet:
                sh.delete()  
    
    def generate_weeks(self, start_date: datetime.date, end_date: datetime.date):
        """
        Sinh danh sách (week_num, monday_of_week) bắt đầu từ start_date
        (đã được làm tròn/lùi về Thứ Hai) cho đến khi vượt quá end_date.
        """

        # 1) Kiểm tra tính hợp lệ
        if start_date > end_date:
            # Tuỳ ý: hoặc trả về danh sách trống
            return []
            # Hoặc raise ValueError("start_date phải <= end_date")

        # 2) Đảm bảo start_date lùi về Thứ Hai (nếu bạn muốn tuần bắt đầu Thứ Hai)
        while start_date.weekday() != 0:  # 0 = Monday
            start_date -= datetime.timedelta(days=1)

        # 3) Duyệt cho đến khi vượt quá end_date
        current_monday = start_date
        week_num = 1
        weeks = []
        while current_monday <= end_date:
            weeks.append((week_num, current_monday))
            week_num += 1
            current_monday += datetime.timedelta(days=7)

        return weeks

    def overlap(self, date_start, date_end, week_start, week_end):
        """
        Trả về True nếu đoạn [date_start, date_end]
        giao nhau với đoạn [week_start, week_end], ngược lại False.

        - date_start, date_end có thể là datetime.date hoặc datetime.datetime
        - week_start, week_end cũng tương tự
        """
        # Chuyển mọi thứ về dạng date
        if isinstance(date_start, datetime.datetime):
            date_start = date_start.date()
        if isinstance(date_end, datetime.datetime):
            date_end = date_end.date()
        if isinstance(week_start, datetime.datetime):
            week_start = week_start.date()
        if isinstance(week_end, datetime.datetime):
            week_end = week_end.date()

        return not (date_end < week_start or date_start > week_end)


    def col_num_to_letter(self,col_num: int) -> str:
        """
        1 -> 'A', 2 -> 'B', ... 26 -> 'Z', 27 -> 'AA', 28 -> 'AB', ...
        """
        letters = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters


root = tcaTool()
root.mainloop()
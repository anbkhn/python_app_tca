import tkinter as tk
from tkinter import ttk, filedialog as fd, messagebox as mg
import xlwings as xw
import pandas as pd
import numpy as np
from xlwings import constants
from xlwings.constants import DeleteShiftDirection

class tcaTool(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Tool TCA")
        self.rowconfigure((0,1), weight=1)
        self.columnconfigure(0, weight=1)

        labelframe1 = tk.LabelFrame(self, text="Tool Xuất nhà cung ứng, kiểm tra tồn kho")
        labelframe1.grid(sticky="NWES", column=0, row=0, padx=(5,5), pady=(5,5))

        data_frame = ttk.Frame(labelframe1, padding=(10,10))
        data_frame.grid(sticky="NWES", column=0, row=0)
        data_frame.columnconfigure(0, weight=2)
        data_frame.columnconfigure(1, weight=1)
        data_frame.columnconfigure(2, weight=4)
        data_frame.rowconfigure((0,2), weight=1)

        button_frame = ttk.Frame(labelframe1, padding=(10,10))
        button_frame.grid(sticky="NWES", column=0, row=1)
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        button_frame.columnconfigure(2, weight=1)

        self.data_file_text = tk.StringVar()
        self.database_file_text = tk.StringVar()
        self.inventory_file_text = tk.StringVar()

        data_label = ttk.Label(data_frame, text="File Dự trù")
        data_button = ttk.Button(data_frame, text="Chọn File", command=self.data_choose_file)
        data_entry = ttk.Entry(data_frame, width=50, textvariable=self.data_file_text, state="disable")

        database_label = ttk.Label(data_frame, text="File Thông tin VTTB")
        database_button = ttk.Button(data_frame, text="Chọn File", command=self.database_choose_file)
        database_entry = ttk.Entry(data_frame, width=50, textvariable=self.database_file_text, state="disable")

        inventory_label = ttk.Label(data_frame, text="File Tồn kho")
        inventory_button = ttk.Button(data_frame, text="Chọn File", command=self.inventory_choose_file)
        inventory_entry = ttk.Entry(data_frame, width=50, textvariable=self.inventory_file_text, state="disable")

        clear_button = ttk.Button(button_frame, text="Xóa sheet", command=self.clear)
        findERP_button = ttk.Button(button_frame, text="Tìm ERP Code", command=self.findERP)
        purchase_button = ttk.Button(button_frame, text="Xuất nhà cung ứng", command=self.purchase)
        check_inventory_button = ttk.Button(button_frame, text="Kiểm tra tồn kho", command=self.check_inventory)
        # exit_button = ttk.Button(button_frame, text="Thoát", command=self.destroy)

        data_label.grid(sticky="EW", column=0, row= 0, padx=5, pady=5)
        data_button.grid(sticky="EW", column=1, row= 0, padx=5, pady=5)
        data_entry.grid(sticky="EW", column=2, row= 0, padx=5, pady=5)

        database_label.grid(sticky="EW", column=0, row= 1, padx=5, pady=5)
        database_button.grid(sticky="EW", column=1, row= 1, padx=5, pady=5)
        database_entry.grid(sticky="EW", column=2, row= 1, padx=5, pady=5)

        inventory_label.grid(sticky="EW", column=0, row= 2, padx=5, pady=5)
        inventory_button.grid(sticky="EW", column=1, row= 2, padx=5, pady=5)
        inventory_entry.grid(sticky="EW", column=2, row= 2, padx=5, pady=5)

        clear_button.grid(sticky="EW", column=0, row= 0, padx=5, pady=5)
        findERP_button.grid(sticky="EW", column=1, row= 0, padx=5, pady=5)
        purchase_button.grid(sticky="EW", column=2, row= 0, padx=5, pady=5)
        check_inventory_button.grid(sticky="EW", column=3, row= 0, padx=5, pady=5)
        # exit_button.grid(sticky="EW", column=4, row= 0, padx=5, pady=5)


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
            except:
                pass
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

                # Đổi tên nhà cung cấp để tạo tên cho từng sheet, chuyển tên sang in thường để tránh việc viết sai ký tự
                df["man"] = df["man"].str.replace(pat="/", repl="_")
                df["man"] = df["man"].str.lower()

                # Lên danh sách tất cả các nhà cung cấp có trong sheet Dự Trù
                man_list = [i for i in df.groupby("man").groups]

                # Sắp xếp các tên nhà cung cấp theo danh sách giảm dần
                man_list.sort(reverse=True)

                # Thực hiện vòng lặp cho từng nhà cung cấp
                for man in man_list:
                    
                    # Tạo Dataframe cho riêng từng nhà cung cấp
                    dfa = df.groupby("man").get_group(man)
                    
                    # Sử dụng tính năng Dataframe để tính tổng số lượng các mã Ordercode 
                    dfb = dfa.groupby("order").agg(quan=("quan","sum"), erp=("erp","min"), des = ("des", "min"), unit = ("unit", "min")).reset_index()
                    
                    # Tạo thêm cột Manufacturer
                    dfb["man"] = man
                    
                    # Tạo thêm cột Số thứ tự
                    dfb["no"] = np.arange(len(dfb.index)) + 1
                    
                    # Sắp xếp lại các cột trong Dataframe
                    dfb = dfb[["no","des","man","order","unit","quan","erp"]]

                    # Tạo biến tiêu đề cho từng sheet mới
                    title = ['No','Description','Manufacturer','Order Code','Unit','Quantity','ERP Code', 'Note', 'Inventory', 'Buying']
                    
                    # Tạo sheet mới ngay sau sheet main Dự trù
                    new_sh = wb.sheets.add(name=man,after=sh_main.name) 
                    
                    # Tạo các tiêu đề cho sheet mới từ dòng 1
                    new_sh['A1'].value = title
                    
                    # Lấy toàn bộ dữ liệu của Dataframe cho vào sheet mới
                    new_sh["A2"].value = dfb.values.tolist()
                    
                    # Định dạng dòng đầu tiên bôi đậm
                    new_sh["A1:j1"].api.Font.Bold = True
                    
                    # Định dạng các ô dữ liệu trong sheet mới phải được căn lề giữa
                    new_sh["A:J"].api.HorizontalAlignment = constants.HAlign.xlHAlignCenter
                    new_sh["A:J"].api.VerticalAlignment = constants.VAlign.xlVAlignCenter
                    
                    # Ép các cột trong sheet mới giãn ô tự động
                    new_sh["A:J"].autofit()
                    
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
                        new_sh[f'A1:J{len(dfb.index)+1}'].api.Borders(i).LineStyle = 1
                        new_sh[f'A1:J{len(dfb.index)+1}'].api.Borders(i).Weight = 2

                total_count = df["quan"].sum()
                total_data_list_count = df.groupby("man").sum()["quan"].sum()
                if total_count == total_data_list_count:
                    total_result = "ĐẠT"
                else:
                    total_result = "KHÔNG ĐẠT"
                mg.showinfo(
                    "Kết quả", 
                    f"""
                    Tổng số lượng sheet Dự trù là {total_count}\n
                    Tổng số lượng liệt kê là {total_data_list_count}\n
                    Kết quả: {total_result} 
                    """
                    )  
                
            except:
                pass
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
                last_row_inv = sh_main_inv.range(f'B{sh_main_inv.cells.last_cell.row}').end('up').row - 1

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
            except:
                pass
        else:
            mg.showinfo("Thông báo","Hãy chọn file Dự trù và file Tồn kho trước")   

    def exportCollection(self):
        # Kiểm tra xem đã có chọn file hay chưa
        if self.data_file_text_2.get() and self.result_file_text_2.get():
            # Tạo trước 1 dataframe rỗng nhằm tổng hợp tất cả các dataframe đọc được từ các file dự trù
            df_combine =pd.DataFrame(columns=["des","man","order","unit", "quan","note","erp"])

            # Sử dụng câu lệnh with này để tránh luôn chạy tiến trình chạy Excel trong task manager
            with xw.App(visible=False) as app:

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
                    df_combine = df_combine.append(df)

                    # Save the workbook
                    wb.save()      

                # Đặt lại địa chỉ index của dataframe combine tính từ 0    
                df_combine.index = np.arange(0, len(df_combine))  

                # Lọc tìm tất cả các tên tủ có trong dataframe combine (bắt đầu bằng dấu =) đưa vào trong 1 dataframe cubicles
                df_cubicles = df_combine[df_combine['order'].str.contains("=", na = False)]['order'] 

                # Tạo các list: Các tủ duy nhất, địa chỉ dòng của các tủ có trong dataframe combine, tên của các tủ có trong dataframe cubicles  
                unique_cubicles, index_cubicles, name_cubicles = df_cubicles.unique(), df_cubicles.index.tolist(), df_cubicles.tolist()

                unique_cubicles.sort()
                
                # Tạo ra một dataframe rỗng là dataframe kết quả cuối cùng
                df_result =pd.DataFrame(columns=["des","man","order","unit", "quan","note","erp"])

                # Chạy vòng lặp cho từng tủ
                for cubicle in unique_cubicles:
                    # Tạo 1 dataframe cho từng tủ
                    df_each_cubicle_with_header =pd.DataFrame(columns=["des","man","order","unit", "quan","note","erp"])

                    # Tạo thêm 1 dataframe của tủ mà có chứa dòng có tên tủ
                    df_each_cubicle_with_header = df_each_cubicle_with_header.append(df_combine[df_combine['order']==cubicle].iloc[0])

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
                            df_each_cubicle_without_header = df_each_cubicle_without_header.append(df_temp)
                    
                    # Sử dụng groupby để tính tổng số lượng các thiết bị trong một tủ
                    df_each_cubicle_without_header = df_each_cubicle_without_header.groupby('order').agg(des = ("des", "min"),man = ("man", "min"),order = ("order", "min"),  unit = ("unit", "min"),quan=("quan","sum"))

                    # Sort theo tên nhà cung cấp
                    df_each_cubicle_without_header = df_each_cubicle_without_header.sort_values('man')  

                    # Đặt lại địa chỉ index của dataframe tính từ 1, tên tủ sẽ có địa chỉ dòng là 0
                    df_each_cubicle_without_header.index = np.arange(1, len(df_each_cubicle_without_header)+1)

                    # Nối dataframe không có header vào dataframe có header để giữ cấu trúc: tên tủ ở dòng đầu tien, sau đó là danh sách các thiết bị trong tủ ở phía dưới
                    df_each_cubicle_with_header =  df_each_cubicle_with_header.append(df_each_cubicle_without_header)

                    # Nối các dataframe có header vào dataframe kết quả cuối cùng. Đây chính là dataframe có thể in ra excel
                    df_result = df_result.append(df_each_cubicle_with_header)
                
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
        # Nếu chưa chọn thì xuất hiện thông báo
        else:
            mg.showinfo("Thông báo","Hãy chọn các file Dự trù và form Dự trù trước")  

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

root = tcaTool()
root.mainloop()
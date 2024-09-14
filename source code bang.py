import os
from tkinter import *
from tkinter import filedialog, messagebox
from docx import Document

# Hàm cập nhật nội dung bảng đầu tiên với ngày và serial RTU do người dùng nhập vào mà giữ nguyên định dạng
def update_first_table(doc, test_date, serial_rtu):
    first_table = doc.tables[0]
    for row in first_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if "<DATE>" in run.text:
                        run.text = run.text.replace("<DATE>", test_date)
                    if "<SERIAL RTU>" in run.text:
                        run.text = run.text.replace("<SERIAL RTU>", serial_rtu)

# Hàm giữ lại các bảng chứa ký tự đặc biệt và xóa các bảng còn lại
def remove_tables_with_special_characters(input_file, output_file, special_tags, test_date, serial_rtu):
    doc = Document(input_file)
    update_first_table(doc, test_date, serial_rtu)
    
    tables_to_delete = []
    for table_index, table in enumerate(doc.tables):
        if table_index == 0 or table_index == 1 or table_index == len(doc.tables) - 1:
            continue
        if len(table.rows) > 1:
            first_cell_text = table.rows[1].cells[0].text.strip()
            if not any(tag in first_cell_text for tag in special_tags):
                tables_to_delete.append(table)
        else:
            tables_to_delete.append(table)
    
    for table in tables_to_delete:
        table._element.getparent().remove(table._element)

    doc.save(output_file)

# Hàm tạo thư mục nếu chưa tồn tại
def create_output_folder(output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

# Hàm xử lý khi nhấn nút "Xuất văn bản"
def export_file():
    try:
        selected_file_name = listbox.get(ACTIVE)
        input_file = os.path.join('template', selected_file_name)

        special_tags = special_tags_entry.get().split()
        test_date = test_date_entry.get()
        serial_rtu = serial_rtu_entry.get()

        if not special_tags or not test_date or not serial_rtu:
            messagebox.showerror("Lỗi", "Vui lòng điền đầy đủ thông tin!")
            return

        output_file_name = f"BienBan_RMU_{selected_file_name.split('.')[0]}_{serial_rtu}.docx"
        output_folder = 'output'
        create_output_folder(output_folder)
        output_file = os.path.join(output_folder, output_file_name)

        remove_tables_with_special_characters(input_file, output_file, special_tags, test_date, serial_rtu)
        messagebox.showinfo("Thành công", f"Đã lưu file vào '{output_file}'")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")

# Tạo cửa sổ chính
root = Tk()
root.title("Tool tạo biên bản thí nghiệm điện")
root.geometry("500x400")
root.iconbitmap("tool.ico")

# Tạo nhãn và listbox để chọn file
Label(root, text="Chọn file Word từ danh sách:").pack(pady=5)
listbox = Listbox(root, width=50, height=6)
listbox.pack(pady=5)

# Liệt kê file trong thư mục "template"
if not os.path.exists('template'):
    os.makedirs('template')
files = [f for f in os.listdir('template') if f.endswith('.docx')]
for file in files:
    listbox.insert(END, file)

# Ô nhập ký tự đặc biệt
Label(root, text="Nhập ký tự đặc biệt (cách nhau bởi dấu cách):").pack(pady=5)
special_tags_entry = Entry(root, width=50)
special_tags_entry.pack(pady=5)

# Ô nhập ngày thử nghiệm
Label(root, text="Nhập ngày thử nghiệm (YYYY-MM-DD):").pack(pady=5)
test_date_entry = Entry(root, width=50)
test_date_entry.pack(pady=5)

# Ô nhập số serial RTU
Label(root, text="Nhập số serial RTU:").pack(pady=5)
serial_rtu_entry = Entry(root, width=50)
serial_rtu_entry.pack(pady=5)

# Nút xuất văn bản
Button(root, text="Xuất văn bản", command=export_file).pack(pady=10)

# Dòng chữ tác giả
Label(root, text="Phần mềm được tạo bởi Thinhlh", font=("Arial", 10, "italic")).pack(side=BOTTOM, pady=10)

# Chạy chương trình
root.mainloop()

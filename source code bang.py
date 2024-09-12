import re
from docx import Document

# Hàm xóa các bảng chứa cặp ký tự đặc biệt
def remove_tables_with_custom_tags(input_file, output_file, user_tags):
    # Mở file Word
    doc = Document(input_file)

    # Duyệt qua từng bảng trong tài liệu và đánh dấu các bảng cần xóa
    tables_to_delete = []
    
    for table in doc.tables:
        table_text = ' '.join([' '.join([cell.text for cell in row.cells]) for row in table.rows])  # Ghép nội dung của tất cả các ô trong bảng

        # Kiểm tra nếu bảng chứa cặp tag mà người dùng nhập vào
        if any(f"/{tag}" in table_text and f"{tag}/" in table_text for tag in user_tags):
            # Nếu bảng chứa cặp tag cần xóa, thêm bảng vào danh sách cần xóa
            tables_to_delete.append(table)

    # Xóa các bảng đã tìm thấy
    for table in tables_to_delete:
        table._element.getparent().remove(table._element)

    # Lưu file kết quả
    doc.save(output_file)

# Nhập tên file đầu vào và file đầu ra
input_file = 'input.docx'  # Tên file Word đầu vào
output_file = 'output.docx'  # Tên file Word đầu ra

# Nhập danh sách các tag mà người dùng muốn xóa bảng
user_tags = input("Nhập các ký tự mà bạn muốn xóa bảng chứa chúng, cách nhau bởi khoảng trắng (ví dụ: k1 k2 k3): ").split()

# Gọi hàm xử lý
remove_tables_with_custom_tags(input_file, output_file, user_tags)

print(f"Đã xóa các bảng chứa các ký tự đặc biệt. Kết quả được lưu vào '{output_file}'.")

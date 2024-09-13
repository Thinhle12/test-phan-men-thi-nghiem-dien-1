import re
from docx import Document

# Hàm xóa các bảng chứa ký tự đặc biệt "/kx" trong hàng đầu tiên
def remove_tables_with_custom_tags(input_file, output_file, user_tag_prefix):
    # Mở file Word
    doc = Document(input_file)

    # Sử dụng biểu thức chính quy để tìm ký tự đặc biệt dạng "/kx" trong hàng đầu tiên
    pattern = re.compile(rf"/{user_tag_prefix}\d+")
    
    # Duyệt qua từng bảng trong tài liệu và đánh dấu các bảng cần xóa
    tables_to_delete = []
    
    for table in doc.tables:
        # Ghép nội dung của hàng đầu tiên trong bảng
        first_row_text = ' '.join([cell.text for cell in table.rows[0].cells])
        
        # Kiểm tra nếu hàng đầu tiên chứa ký tự đặc biệt theo định dạng "/kx"
        if pattern.search(first_row_text):
            # Nếu hàng đầu tiên chứa ký tự đặc biệt, thêm bảng vào danh sách cần xóa
            tables_to_delete.append(table)

    # Xóa các bảng đã tìm thấy
    for table in tables_to_delete:
        table._element.getparent().remove(table._element)

    # Sau khi xóa các bảng chứa ký tự đặc biệt, kiểm tra các bảng còn lại
    for table in doc.tables:
        # Lấy ô đầu tiên của cột đầu tiên
        first_column_text = ' '.join([row.cells[0].text for row in table.rows])
        # Tạo biểu thức chính quy để tìm ký tự đặc biệt dạng "/kx" nhưng x không phải là x của người dùng
        pattern_other = re.compile(rf"/{user_tag_prefix}\d+")
        
        # Duyệt qua từng hàng trong bảng và xóa hàng đầu tiên nếu ô đầu tiên có ký tự đặc biệt
        if any(pattern_other.search(cell) for cell in first_column_text.split()):
            # Xóa hàng đầu tiên của bảng
            table.rows[0]._element.getparent().remove(table.rows[0]._element)
    
    # Lưu file kết quả
    doc.save(output_file)

# Nhập tên file đầu vào và file đầu ra
input_file = 'input.docx'  # Tên file Word đầu vào
output_file = 'output.docx'  # Tên file Word đầu ra

# Nhập tiền tố ký tự đặc biệt từ người dùng (ví dụ: "k")
user_tag_prefix = input("Nhập tiền tố ký tự đặc biệt (ví dụ: k): ")

# Gọi hàm xử lý
remove_tables_with_custom_tags(input_file, output_file, user_tag_prefix)

print(f"Đã xóa các bảng có hàng đầu tiên chứa ký tự đặc biệt theo định dạng /{user_tag_prefix}x và các hàng đầu tiên trong bảng còn lại chứa ký tự đặc biệt khác. Kết quả được lưu vào '{output_file}'.")

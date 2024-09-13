import re
from docx import Document

# Hàm xóa các bảng chứa ký tự đặc biệt trong ô đầu tiên của hàng thứ hai
def remove_tables_with_special_characters(input_file, output_file, special_tags):
    # Mở file Word
    doc = Document(input_file)
    special_tags_uppercase = special_tags.uppercase
    # Duyệt qua từng bảng trong tài liệu và đánh dấu các bảng cần xóa
    tables_to_delete = []
    
    for table_index, table in enumerate(doc.tables):
        print(f"Kiểm tra bảng {table_index + 1}...")
        
        # Kiểm tra xem bảng có đủ 2 hàng hay không (phải có ít nhất 2 hàng để xét hàng thứ hai)
        if len(table.rows) > 1:
            # Lấy nội dung của ô đầu tiên trong hàng thứ hai
            first_cell_text = table.rows[1].cells[0].text.strip()

            # In ra nội dung của ô đầu tiên trong hàng thứ hai để kiểm tra
            print(f"Nội dung ô đầu tiên của hàng thứ hai: '{first_cell_text}'")

            # Kiểm tra nếu ô đầu tiên của hàng thứ hai chứa bất kỳ ký tự đặc biệt nào mà người dùng nhập vào
            if any(tag in first_cell_text for tag in special_tags_uppercase):
                print(f"Bảng {table_index + 1} có chứa ký tự đặc biệt và sẽ bị xóa.")
                # Nếu có, thêm bảng vào danh sách cần xóa
                tables_to_delete.append(table)
            else:
                print(f"Bảng {table_index + 1} không chứa ký tự đặc biệt.")
        else:
            print(f"Bảng {table_index + 1} không có đủ 2 hàng, bỏ qua.")
    
    # Xóa các bảng đã tìm thấy
    for table in tables_to_delete:
        table._element.getparent().remove(table._element)

    # Lưu file kết quả
    doc.save(output_file)
    print(f"Đã lưu kết quả vào '{output_file}'.")

# Nhập tên file đầu vào và file đầu ra
input_file = 'input.docx'  # Tên file Word đầu vào
output_file = 'output.docx'  # Tên file Word đầu ra

# Nhập danh sách các ký tự đặc biệt mà người dùng muốn tìm và xóa bảng chứa chúng
special_tags = input("Nhập các ký tự đặc biệt, cách nhau bởi khoảng cách (ví dụ: k1 k2 k3): ").split()

# Gọi hàm xử lý
remove_tables_with_special_characters(input_file, output_file, special_tags)

print(f"Đã xóa các bảng có hàng thứ hai chứa ký tự đặc biệt. Kết quả được lưu vào '{output_file}'.")

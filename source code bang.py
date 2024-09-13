import re
from docx import Document

# Hàm giữ lại các bảng chứa ký tự đặc biệt và xóa các bảng còn lại
def remove_tables_with_special_characters(input_file, output_file, special_tags):
    # Mở file Word
    doc = Document(input_file)

    # Duyệt qua từng bảng trong tài liệu và đánh dấu các bảng cần giữ lại
    
    tables_to_delete = []
    
    for table_index, table in enumerate(doc.tables):
        print(f"Kiểm tra bảng {table_index + 1}...")

        # Trường hợp bảng đầu tiên, bảng thứ hai và bảng cuối cùng luôn được giữ lại
        if table_index == 0 or table_index == 1 or table_index == len(doc.tables) - 1:
            print(f"Bảng {table_index + 1} là bảng đặc biệt (bảng đầu, bảng thứ hai hoặc bảng cuối cùng) và sẽ được giữ lại.")
            
            continue

        # Kiểm tra xem bảng có đủ 2 hàng hay không (phải có ít nhất 2 hàng để xét hàng thứ hai)
        if len(table.rows) > 1:
            # Lấy nội dung của ô đầu tiên trong hàng thứ hai
            first_cell_text = table.rows[1].cells[0].text.strip()

            # In ra nội dung của ô đầu tiên trong hàng thứ hai để kiểm tra
            print(f"Nội dung ô đầu tiên của hàng thứ hai: '{first_cell_text}'")

            # Kiểm tra nếu ô đầu tiên của hàng thứ hai chứa bất kỳ ký tự đặc biệt nào mà người dùng nhập vào
            if any(tag in first_cell_text for tag in special_tags):
                print(f"Bảng {table_index + 1} có chứa ký tự đặc biệt và sẽ được giữ lại.")
                
            else:
                print(f"Bảng {table_index + 1} không chứa ký tự đặc biệt và sẽ bị xóa.")
                tables_to_delete.append(table)
        else:
            print(f"Bảng {table_index + 1} không có đủ 2 hàng, sẽ bị xóa.")

    
    for table in tables_to_delete:
        table._element.getparent().remove(table._element)

    # Lưu file kết quả
    doc.save(output_file)
    print(f"Đã lưu kết quả vào '{output_file}'.")

# Nhập tên file đầu vào và file đầu ra
input_file = 'input.docx'  # Tên file Word đầu vào
output_file = 'output.docx'  # Tên file Word đầu ra

# Nhập danh sách các ký tự đặc biệt mà người dùng muốn tìm
special_tags = input("Nhập các ký tự đặc biệt, cách nhau bởi khoảng cách (ví dụ: k1 k2 k3): ").split()

# Gọi hàm xử lý
remove_tables_with_special_characters(input_file, output_file, special_tags)

print(f"Đã xóa các bảng không có ký tự đặc biệt. Kết quả được lưu vào '{output_file}'.")

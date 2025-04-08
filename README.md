Cấu Hình Power Automate Flow
Để tự động hóa quy trình, bạn cần tạo một flow trong Power Automate với các bước sau:

Bước 1: Kích Hoạt và Đặt Biến
Sử dụng trigger "Manually trigger a flow" để thử nghiệm (hoặc trigger phù hợp khác).
Thêm hành động "Initialize variable" để đặt columnNames (danh sách cột cần sao chép, ví dụ: ["col1", "col2", "col3"]).
Bước 2: Chạy Script Lấy Dữ Liệu
Thêm hành động "Excel Online (Business) - Run script" để chạy GetColumnData trên file Excel gốc.
Cấu hình:
Location: Vị trí file Excel gốc (OneDrive hoặc SharePoint).
Document Library: Thư viện chứa file.
File: Chọn file Excel gốc.
Script: Chọn GetColumnData.
Parameters: Đặt columnNames là biến columnNames đã khởi tạo.
Kết quả sẽ được lưu trong biến body('Run_script')['result'], là một mảng 2D cho mỗi cột.
Bước 3: Dán Dữ Liệu Vào File Mới
Thêm hành động "Excel Online (Business) - Run script" để chạy PasteColumnData trên file Excel mới.
Cấu hình:
Location: Vị trí file Excel mới.
Document Library: Thư viện chứa file.
File: Chọn file Excel mới (đảm bảo file mới đã có sẵn header với cùng tên cột).
Script: Chọn PasteColumnData.
Parameters:
Name: data, Value: body('Run_script')['result'] (kết quả từ bước trước).
Name: columnNames, Value: @variables('columnNames')
.
function main(workbook: ExcelScript.WORKBOOK; columnNames: string[]): string[][][] {
    let sheet = worksheet.getActiveWorksheet();
    let allData = [];
    for (let name of columnNames) {
        try {
            let range = sheet.getRangeByColumns(name);
            if (!range || range.getCellCount() == 0) {
                throw new Error(`Column ${name} not found or empty.`);
            }
            let dataRange = range.getOffsetRange(1); // Bắt đầu từ hàng 2, loại bỏ header
            let data = dataRange.getValues();
            allData.push(data);
        } catch (error) {
            console.error(`Error getting data for column ${name}: ${error.message}`);
            // Bỏ qua cột này nếu có lỗi
        }
    }
    return allData;
}
function main(workbook: ExcelScript.WORKBOOK; data: string[][][]; columnNames: string[]) {
    let sheet = worksheet.getActiveWorksheet();
    for (let i = 0; i < data.length; i++) {
        let name = columnNames[i];
        try {
            let range = sheet.getRangeByColumns(name);
            if (!range || range.getCellCount() == 0) {
                throw new Error(`Column ${name} not found or empty.`);
            }
            let dataRange = range.getOffsetRange(1); // Bắt đầu từ hàng 2
            dataRange.setValues(data[i]);
        } catch (error) {
            console.error(`Error pasting data for column ${name}: ${error.message}`);
            // Bỏ qua cột này nếu có lỗi
        }
    }
}

# Excel Compare PRO 7.1.3

Tool gồm hai chức năng độc lập:

1. **So sánh dữ liệu**: so sánh hai file và xuất báo cáo trạng thái.
2. **Dò và ghi dữ liệu**: dò theo một hoặc nhiều cặp khóa, sau đó ghi một hoặc
   nhiều cột từ file nguồn sang bản sao của file đích.

## Chạy source

```bash
pip install -r requirements.txt
python main.py
```

Tool hỗ trợ `.xlsx`. File kết quả luôn được lưu thành file mới để bảo vệ file gốc.

Phiên bản 6.3 bổ sung tự tìm dòng tiêu đề, nguồn/đích chung file, tìm nhanh tên
cột, kiểm tra cấu hình và xem trước 20 dòng mà không ghi dữ liệu.

Phiên bản 6.4 bổ sung nhật ký từng ô thay đổi, danh sách không tìm thấy, báo cáo
khóa nguồn trùng, tô vàng ô cập nhật và hoàn tác từ nhật ký sang một file mới.

Phiên bản 7.0 bổ sung hồ sơ cấu hình JSON có thể dùng lại, Sheet tổng hợp chỉ số,
và gợi ý khớp gần đúng cho dòng không tìm thấy. Gợi ý chỉ phục vụ kiểm tra và
không bao giờ tự động ghi dữ liệu.

Phiên bản 7.0.1 thay danh sách chọn cột bằng combobox có giới hạn chiều cao,
thanh cuộn dọc, hỗ trợ con lăn chuột và điều hướng bàn phím.

Phiên bản 7.1 bổ sung checkbox chung tạo báo cáo phụ (mặc định tắt) và danh sách
tác vụ chạy lô. Mỗi tác vụ có file/Sheet nguồn, file/Sheet đích, nhiều khóa và
nhiều cặp cột lấy/ghi độc lập.

Phiên bản 7.1.1 đồng bộ chức năng So sánh dữ liệu: chọn Sheet A/B, dòng tiêu đề
A/B, tự tìm tiêu đề và nạp lại cột; không còn buộc đọc Sheet đầu tiên/dòng 1.

Phiên bản 7.1.2 bổ sung vùng checkbox chọn các trường xuất từ File A/B. Mặc định
không chọn; chỉ trường được tích mới xuất vào báo cáo so sánh.

Phiên bản 7.1.3 sắp xếp báo cáo theo khối File A → File B → kết quả, giữ đúng
thứ tự cột trong từng Sheet. Cột trùng tên được ghi rõ hậu tố [A]/[B].

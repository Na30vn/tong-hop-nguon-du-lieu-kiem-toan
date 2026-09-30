# Dashboard tổng hợp nguồn dữ liệu phục vụ kiểm toán

Trang web tĩnh nằm trong thư mục `dist`. Dữ liệu được tạo từ các thư mục hồ sơ đơn vị ở thư mục cha và thư mục `CV yêu cầu gốc`.

Để quét lại file sau khi bổ sung hoặc thay đổi hồ sơ, chạy `cap-nhat-dashboard.cmd`. Script sẽ đọc lại Excel/PDF, cập nhật số liệu và sao chép các tệp gốc vào khu vực tải xuống của website.

Các nhận xét kiểm soát chất lượng trong `build_data.py` là các điểm tổng hợp cần được rà soát nghiệp vụ; khi đơn vị gửi bản điều chỉnh, cập nhật lại các nhận xét này trước khi xuất bản phiên bản mới.

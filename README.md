# Phân Tích & Dự Báo Cổ Phiếu HNX30

Dự án cung cấp công cụ phân tích định lượng cho rổ cổ phiếu **HNX30** trên thị trường chứng khoán Việt Nam. Notebook thực hiện từ xử lý dữ liệu, tính toán các chỉ số tài chính, phân nhóm cổ phiếu đến dự báo giá bằng mô hình **ARIMA**.



## Tổng Quan
- **Xử lý dữ liệu:** Làm sạch và hợp nhất dữ liệu giá, khối lượng và chỉ số thị trường.  
- **Phân tích rủi ro:** Tính Beta, Sharpe Ratio, Max Drawdown.  
- **Danh mục đầu tư:** Phân loại cổ phiếu thành nhóm **Ổn định (Low Beta)** và **Mạo hiểm (High Beta)**.  
- **Dự báo:** ARIMA với biến ngoại sinh (lợi suất trái phiếu 10 năm) dự báo giá cổ phiếu hàng tháng.  



## Dữ Liệu Đầu Vào

* `HNX30_HNXINDEX_gop.xlsx`: Giá và khối lượng HNX30.
* `HNXINDEX_Lichsu_Clean.xlsx`: Lịch sử HNX-INDEX.
* `10ybondyield_filled.xlsx`: Lợi suất trái phiếu 10 năm.



## Quy Trình Phân Tích

1. Đồng bộ dữ liệu giá cổ phiếu, chỉ số và lãi suất.
2. Tính lợi suất định kỳ.
3. Hồi quy OLS tính Beta.
4. Xây dựng danh mục: Top 5 Beta thấp nhất (Phòng thủ), Top 5 Beta cao nhất (Tấn công).
5. Trực quan hóa hiệu quả đầu tư.
6. Dự báo giá bằng ARIMA với Grid Search tham số tối ưu.



## Kết Quả

* **Excel:** Beta tích lũy, xếp hạng hiệu quả, dự báo giá từng cổ phiếu.
* **Biểu đồ:** So sánh Beta, lợi nhuận tích lũy, Rủi ro-Lợi nhuận, CAPM, Sharpe Ratio, Max Drawdown.
* **Dự báo ARIMA:** Biểu đồ giá thực tế so với dự báo.



## Ghi Chú

* Dữ liệu từ 2020 đến nay.
* ARIMA sử dụng biến ngoại sinh là **Lợi suất trái phiếu** để nâng cao độ chính xác.

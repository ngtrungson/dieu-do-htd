# EMS Simulator: Nộp bài & Công bố điểm trực tuyến

## Thành phần

| File | Vai trò |
|---|---|
| `index.html` | Trang mô phỏng (đưa lên GitHub Pages, thay cho file cũ) |
| `apps-script/Code.gs` | Máy chủ nhận bài, chấm lại điểm, ghi vào Google Sheets |
| `tools/grader.js` | Mô hình tính toán và thang chấm, dùng chung cho cả trang web lẫn máy chủ |
| `tools/server.gs` | Mã nguồn phần máy chủ (phần ghép vào Code.gs) |
| `tools/build.py` | Đồng bộ `grader.js` vào `index.html` và sinh `Code.gs` |

## Cài đặt máy chủ (một lần, khoảng 10 phút)

1. Tạo một **Google Sheet** mới, ví dụ đặt tên "Điểm Điều độ HTĐ 2026".
2. Vào **Tiện ích mở rộng → Apps Script**. Xóa nội dung mẫu, dán toàn bộ `apps-script/Code.gs` vào rồi bấm Lưu.
3. Chọn hàm `setup` trên thanh công cụ, bấm **Chạy** và cấp quyền khi Google hỏi. Sheet sẽ có 4 trang tính: `BaiNop`, `BangDiem`, `DanhSach`, `CauHinh`.
4. Bấm **Triển khai → Tùy chọn triển khai mới → Ứng dụng web** và đặt:
   - *Thực thi với tư cách*: **Tôi**
   - *Người có quyền truy cập*: **Bất kỳ ai**
5. Sao chép **URL ứng dụng web** (dạng `https://script.google.com/macros/s/.../exec`).
6. Mở `index.html`, tìm dòng `const SUBMIT_API_URL = '';` rồi dán URL vào giữa hai dấu nháy.
7. Đẩy `index.html` lên repo `ngtrungson/dieu-do-htd` trên GitHub. Trang sẽ có hiệu lực sau 1–2 phút.

> Khi sửa `Code.gs` về sau, cần vào **Triển khai → Quản lý triển khai → Chỉnh sửa → Phiên bản mới** thì thay đổi mới có hiệu lực. URL giữ nguyên.

## Điều khiển hằng ngày (trang tính `CauHinh`)

Mọi thay đổi ở đây có hiệu lực ngay, không cần triển khai lại.

| Tham số | Ý nghĩa |
|---|---|
| `MO_NOP_BAI` | `TRUE` nhận bài, `FALSE` khóa nộp bài |
| `HAN_NOP` | Ví dụ `2026-10-30 23:59`. Để trống nếu không giới hạn |
| `SO_LAN_NOP_TOI_DA` | Số lượt nộp của mỗi sinh viên (mặc định 3) |
| `CACH_TINH_DIEM` | `cao_nhat` lấy lần nộp điểm cao nhất, `lan_cuoi` lấy lần nộp cuối |
| `CONG_BO_DIEM` | `TRUE` công khai bảng điểm trên web ngay khi sinh viên nộp |
| `KIEM_TRA_DANH_SACH` | `TRUE` chỉ nhận MSSV có trong trang tính `DanhSach` |
| `AN_MSSV` | `TRUE` che bớt chữ số MSSV trên bảng điểm công khai |

## Chấm điểm và điều chỉnh

- Máy chủ **tự chấm lại** từ thông số thanh trượt mà sinh viên gửi lên. Điểm hiển thị trên trình duyệt chỉ là điểm dự kiến, nên sinh viên không sửa được điểm.
- `BaiNop`: mỗi dòng là một lần nộp, kèm chi tiết chấm và bản ghi dạng JSON để đối chiếu.
- `BangDiem`: mỗi sinh viên một dòng. Muốn sửa điểm (ví dụ chấm lại phần nhận xét), nhập vào cột **Điểm GV điều chỉnh**. Bảng điểm trên web sẽ dùng giá trị này và gắn dấu `*`. Cột này được giữ nguyên khi sinh viên nộp lại.
- Thang điểm: 9 kịch bản × 10 điểm + Nhận xét 10 điểm = 100, quy về hệ 10. Chi tiết xem tab *Hướng dẫn Thực hành* trên trang.

## Sửa kịch bản hoặc thang điểm

Chỉ sửa trong `tools/grader.js`, sau đó chạy:

```
python tools/build.py
```

Lệnh này cập nhật đồng thời `index.html` và `apps-script/Code.gs`. Sau đó dán lại Code.gs vào Apps Script và tạo phiên bản triển khai mới.

## Giới hạn cần biết

- Không có đăng nhập: ai biết MSSV của người khác cũng có thể nộp thay. Bật `KIEM_TRA_DANH_SACH` sẽ hạn chế MSSV lạ. Mọi lần nộp đều được lưu kèm thời gian trong `BaiNop` để đối chiếu.
- Mức trừ điểm khi dùng "Gợi ý OPF" dựa vào trình duyệt báo lên, nên chỉ có tác dụng răn đe, không phải biện pháp chống gian lận.
- Điểm nhận xét được chấm sơ bộ theo độ dài. Giảng viên nên đọc lại trong cột `Nhận xét của SV`.

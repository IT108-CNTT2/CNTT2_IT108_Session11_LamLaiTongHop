# Phần 1 – Hoàn thiện SRS bằng Heading Styles và Auto TOC

## 1. Mở rộng tiêu đề thành 5 cấp

Có thể bổ sung:

```text
1. Giới thiệu
1.1. Mục tiêu hệ thống
1.1.1. Đối tượng sử dụng

2. Chức năng hệ thống
2.1. Quản lý mượn sách
2.1.1. Ghi nhận phiếu mượn

3. Yêu cầu phi chức năng
```

## 2. Gán Heading

- `1. Giới thiệu`, `2. Chức năng hệ thống`, `3. Yêu cầu phi chức năng` → **Heading 1**
- `1.1`, `2.1` → **Heading 2**
- `1.1.1`, `2.1.1` → **Heading 3**

## 3. Nội dung mô tả

### Yêu cầu chức năng

Hệ thống cho phép nhân viên quản lý việc mượn và trả sách. Hệ thống cũng hỗ trợ tra cứu thông tin sách và tình trạng mượn.

### Yêu cầu phi chức năng

Hệ thống cần hoạt động ổn định và dễ sử dụng. Dữ liệu người dùng và dữ liệu mượn sách cần được lưu trữ an toàn.

## 4. Auto TOC

Vào:

```text
References → Table of Contents → Automatic Table
```

Sau khi thêm nội dung làm thay đổi số trang:

```text
Right Click TOC → Update Field → Update entire table
```

Khi đó số trang trong mục lục sẽ tự cập nhật lại.

---

# Phần 2 – Chuẩn hóa dữ liệu và AutoFilter

## 1. Ba lỗi trong bảng

| Lỗi | Vị trí | Quy tắc bị vi phạm |
|---|---|---|
| Có hàng trống giữa dữ liệu | Giữa các bản ghi | Không để hàng trống trong vùng dữ liệu |
| Có ô bị Merge | Dòng `Tổng cộng` | Không Merge Cells trong bảng dữ liệu |
| Dòng `Tổng cộng` nằm trong vùng dữ liệu | Cuối bảng | Mỗi hàng chỉ nên là một Record |

## 2. Sửa bảng

Cách sửa:

- Xóa hàng trống.
- Bỏ Merge Cells.
- Chuyển dòng `Tổng cộng` ra ngoài vùng dữ liệu.
- Mỗi cột có một tiêu đề riêng.
- Mỗi hàng chỉ chứa một bản ghi.

## 3. Record và Field

- **Record:** một hàng dữ liệu của một lần mượn sách.
- **Field:** một cột dữ liệu.

Ví dụ:

```text
Mã sách = S001
```

là một Field của Record chứa thông tin lần mượn đó.

## 4. AutoFilter

Chọn một ô trong bảng rồi nhấn:

```text
Ctrl + Shift + L
```

Sau đó tại cột **Tình trạng**:

```text
Filter → chỉ chọn "Quá hạn" → OK
```

Excel sẽ chỉ hiển thị các bản ghi có tình trạng `Quá hạn`.

---

# Phần 3 – Nối dữ liệu bằng VLOOKUP

## 1. Công thức lấy Tên Sách

Giả sử `Mã Sách` nằm ở ô `C2`:

```excel
=VLOOKUP(C2,DanhMucSach!$A$2:$C$4,2,FALSE)
```

Trong đó:

- `C2`: Mã Sách cần tìm.
- `$A$2:$C$4`: bảng tra cứu.
- `2`: lấy dữ liệu ở cột Tên Sách.
- `FALSE`: tìm chính xác.

## 2. Mã sách S004

Với mã:

```text
S004
```

VLOOKUP sẽ trả về:

```text
#N/A
```

Vì `S004` không tồn tại trong bảng `DanhMucSach`.

## 3. Dùng IFERROR

```excel
=IFERROR(VLOOKUP(C2,DanhMucSach!$A$2:$C$4,2,FALSE),"Không tìm thấy sách")
```

Nếu mã sách không tồn tại thì sẽ hiện:

```text
Không tìm thấy sách
```

thay vì lỗi `#N/A`.

---

# Phần 4 – Slide báo cáo tiến độ

## 1. Nội dung slide

### Assertion Title

```text
Phần lớn lượt mượn vẫn đang ở trạng thái đang mượn
```

### Con số nổi bật

```text
3/5 lượt mượn đang ở trạng thái "Đang mượn"
```

Có thể ghi thêm:

```text
1 lượt quá hạn
```

### Biểu đồ

Em chọn **Bar Chart**.

Lý do: dữ liệu gồm các nhóm trạng thái riêng như:

- Đang mượn
- Quá hạn
- Đã trả

Bar Chart giúp so sánh số lượng giữa các nhóm dễ nhìn hơn.

Không chọn Line Chart vì dữ liệu không phải chuỗi theo thời gian.

## 2. Rule 10/20/30

Theo quy tắc:

- Tổng bài trình bày: **10 slide**
- Tổng thời gian: **20 phút**
- Slide này chiếm: **1/10 slide**
- Thời gian trung bình:

```text
20 / 10 = 2 phút
```

Vậy slide này dự kiến trình bày khoảng **2 phút**.

Font chữ nên từ khoảng **30pt trở lên** để dễ đọc.

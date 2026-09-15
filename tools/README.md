# tools/ — Bộ dựng giáo trình Word

Tái tạo file `Ho_so_bai_giang_lai_xe_2026_v1.0.docx` từ các bản nháp Markdown trong `drafts/`
và kho ảnh của file gốc `HSBG.docx`.

## Yêu cầu

```bash
pip install python-docx lxml
apt-get install -y libreoffice-writer poppler-utils   # chỉ cần khi muốn đo số trang mục lục
```

## Chạy

```bash
cd /đường/dẫn/HoSoBaiGiang
rm -rf build && mkdir -p build
unzip -q HSBG.docx -d build/x                 # lấy kho ảnh của file gốc
export HSBG_WORK="$PWD/build"
python3 tools/extents.py > "$HSBG_WORK/extents.json"   # đo kích thước hiển thị gốc của 35 ảnh
bash tools/iterate.sh                          # dựng + đo số trang, lặp tới khi ổn định
```

Kết quả: `Ho_so_bai_giang_lai_xe_2026_v1.0.docx` ở thư mục gốc dự án.

Nếu không có LibreOffice, chạy thẳng `python3 tools/build.py` — file Word vẫn đúng,
mục lục để trống số trang và Word sẽ tự điền khi mở (đã bật `updateFields`).

## Các file

| File | Vai trò |
|---|---|
| `style.py` | Bảng màu, font, hộp thông tin, viền bảng, hàm chèn ảnh (kể cả **EMF**) |
| `render.py` | Bộ render Markdown → Word: heading, bảng, danh sách, hộp thông tin, ảnh, chú thích |
| `build.py` | Ráp tài liệu: bìa, trang hướng dẫn, mục lục, 10 phần nội dung, phụ lục, header/footer |
| `measure.py` | Đọc PDF, đo số trang thật của từng heading → `toc.json` |
| `iterate.sh` | Lặp build → đo → build cho tới khi số trang mục lục ổn định |
| `extents.py` | Trích kích thước hiển thị gốc của từng ảnh trong `HSBG.docx` |
| `extract.py` | Trích toàn văn `HSBG.docx` ra text có đánh dấu đoạn/bảng (dùng ở Bước 1) |
| `imgmap.py` | Lập bản đồ 37 vị trí chèn ảnh của `HSBG.docx` (dùng ở Bước 1) |
| `toc.json` | Số trang đã đo của 97 heading — dữ liệu dựng sẵn cho mục lục |

## Ghi chú kỹ thuật

- **EMF**: 10 sơ đồ trong file gốc ở định dạng EMF. `python-docx` không hỗ trợ sẵn, nên
  `style.add_picture_any()` chèn image part và dựng XML `w:drawing` thủ công. Word đọc EMF
  native nên không cần chuyển đổi, giữ nguyên chất lượng vector.
- **Kích thước ảnh**: lấy đúng `wp:extent` của file gốc, chỉ thu nhỏ khi vượt bề rộng vùng chữ
  (6,1 inch), giữ nguyên tỷ lệ.
- **Mục lục**: vừa là field `TOC` (Word tự cập nhật bằng Ctrl+A → F9), vừa có nội dung dựng sẵn
  trong vùng kết quả của field để đọc được ngay mà không cần cập nhật.
- **Font**: Arial. Nếu máy không có Arial, Word thay bằng font tương đương; tiếng Việt vẫn đúng dấu.

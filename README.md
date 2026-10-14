# Quý ông người pháp: ôn thi Chuẩn kỹ năng CNTT cơ bản

Web ôn thi trắc nghiệm từ 6 file PDF (718 câu, 6 module). Chọn đáp án, biết đúng sai ngay,
nếu sai thì hiện đáp án đúng. Có thi thử có tính giờ và danh sách câu sai để ôn lại.

## Chạy

```bash
npm install        # cài phụ thuộc (chỉ dùng khi phát triển và kiểm thử)
npm run build      # sinh font, icon, dữ liệu câu hỏi vào thư mục web
npm run serve      # mở http://127.0.0.1:4173
npm test           # unit test (node:test) + E2E (Playwright, dùng Edge có sẵn)
npm run shots      # chụp ảnh các màn hình ở 320/375/768/1024/1440, sáng và tối
```

Cũng có thể mở thẳng `index.html` bằng trình duyệt, không cần server.

## Cấu trúc

- `index.html`, `css/`, `js/`: giao diện thuần HTML/CSS/JS, không cần bước build.
- `js/core.js`: logic thuần (chấm điểm, tiến độ, chọn đề), có unit test.
- `data/questions.js`: dữ liệu sinh từ PDF. `tools/parse_pdfs.py` trích PDF sang JSON
  (đáp án đúng là chữ cái màu đỏ trong PDF; Module 2 câu 33 được khoanh tròn nên khai báo thủ công).
- Tiến độ lưu trong localStorage của trình duyệt; bài thi đang làm dở lưu trong sessionStorage.

## Hướng dẫn thi thực hành (tự luận)

7 đề, 21 câu Word, Excel, PowerPoint tại `#/thuc-hanh`. Mỗi câu có đề bài dạng chữ, các bước
giải chi tiết theo từng công cụ, công thức Excel (có nút chép), danh sách tự kiểm tra và lỗi thường gặp.

- Dữ liệu nằm ở `data/practical/de1.json` đến `de7.json`, `npm run build` gói thành `data/practical.js`.
- Đề được chuyển từ ảnh chụp mờ và file bài làm mẫu. Chỗ đọc không chắc chắn được ghi trong mục
  "Ghi chú về độ chính xác của đề" của từng câu. Đề 4 thiếu ảnh trang 2 nên Câu 3 chưa có đề bài.
- `tests/unit/practical-data.test.js` kiểm tra cấu trúc dữ liệu và quy ước trình bày của cả 7 đề.

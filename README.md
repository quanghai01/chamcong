# Attendance Pro

Hệ thống xử lý chấm công tự động từ file Excel.

## Cách chạy cục bộ
1. Cài đặt dependencies: `npm install`
2. Khởi chạy server: `npm start`
3. Truy cập: `http://localhost:3000`

## Cách triển khai (Deploy)
Dự án này đã được tối ưu hóa để triển khai lên các nền tảng như **Render**, **Heroku** hoặc **Railway**.

### Triển khai lên Render:
1. Đưa mã nguồn lên GitHub.
2. Tạo một "Web Service" mới trên Render.
3. Kết nối với repo GitHub của bạn.
4. Render sẽ tự động nhận diện:
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`

### Triển khai lên Vercel:
1. Cài đặt Vercel CLI hoặc kết nối qua Vercel Dashboard.
2. Chạy lệnh `vercel` (Vercel sẽ tự cấu hình cho project Node.js).

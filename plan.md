# Build Plan: Trang Chia Sẻ Template Văn Phòng

## Stack
- Vanilla HTML + CSS + JS thuần — không framework
- Dữ liệu template lưu trong JSON (templates.json)
- Search bằng ngôn ngữ tự nhiên (keyword matching)
- Lọc theo category

## Deploy
- Source code: repo `sslakd/van-phong`
- Build output → `sslakd/Jason-hut` repo → `van-phong/` subfolder
- URL: `https://sslakd.com/van-phong/`

## File Structure
```
/ (project root)
├── index.html          # Trang chủ
├── assets/
│   ├── style.css       # CSS
│   └── app.js          # JS logic
├── data/
│   └── templates.json  # Dữ liệu template
├── templates/           # Thư mục chứa file template (để download)
│   ├── excel/
│   └── word/
├── CNAME               # Custom domain
└── README.md
```

## Template Categories
1. **Kế toán - Tài chính** — Bảng lương, thu chi, hóa đơn, báo cáo tài chính
2. **Quản lý dự án** — Gantt chart, task tracker, timesheet, KPI
3. **Nhân sự** — Chấm công, đơn xin nghỉ phép, đánh giá nhân viên
4. **Kinh doanh - Bán hàng** — Đơn hàng, invoice, CRM, báo giá
5. **Hành chính - Văn thư** — Công văn, biên bản họp, giấy tờ
6. **Marketing** — Calendar content, KPI marketing, brief
7. **Cá nhân** — Budget, todo list, habit tracker
8. **Học tập** — Điểm số, lịch học, bài tập

## Natural Language Search
- User gõ: "mẫu bảng lương", "theo dõi task", "tính thuế"
- Match trên tên, mô tả, từ khóa
- Fuzzy matching cơ bản

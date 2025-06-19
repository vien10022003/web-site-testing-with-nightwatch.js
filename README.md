# Kiểm thử chức năng trang web với Nightwatch.js

## Cài đặt

Để bắt đầu, cài đặt Nightwatch.js và các dependency cần thiết trong terminal của project:

```bash
npm install
```


## Thay thế file excel mà bạn muốn kiểm thử

Mặc định trong code là đang sử dụng file `nightwatch/ACTVN_TestCases (1).xlsx` để chạy test case, bạn có thể thay thế file này, hoặc vào sửa code để trỏ vào file mà bạn muốn test

## Chạy kiểm thử

Để thực hiện kiểm thử, sử dụng lệnh sau trong terminal:
```bash
npx nightwatch nightwatch/excelTest.js
```

## Mô tả

Đây là một project đơn giản để kiểm thử chức năng của trang web bằng Nightwatch.js. 
File `excelTest.js` chứa các test case được mô tả trong file Excel tương ứng.
Kết quả kiểm thử sẽ được viết vào file `nightwatch/test-data-result.xlsx` (đừng mở file này khi chạy kiểm thử vì có thể sẽ xung đột không ghi được vào file)

## Cần trợ giúp?

Nếu có bất kỳ thắc mắc nào, vui lòng liên hệ với chúng tôi

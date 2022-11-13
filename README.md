# GoogleTranslateXL
 Hàm dịch và phát hiện ngôn ngữ siêu nhanh

Hàm dưới đây sẽ giúp dịch thuật và phát hiện ngôn ngữ cho Office và VBA Editor
Code VBA hoạt động yêu cầu có Internet để dịch thông qua Google Translate​
Dịch ra nhiều thứ tiếng và nhiều thứ tiếng ra tiếng Việt​
Ưu điểm: Khi viết code hoặc copy code tham khảo qua mạng, thường sẽ xuất hiện những thuật ngữ mới, vậy nên cần đến dịch thuật.
Với dữ liệu Excel thì khá nhiều ngôn ngữ nên việc dịch thuật là đương nhiên.
-------------------------------------------------------------------------------------

-------------------------------------------------------------------------
1. Dịch
- Điền mã ngôn ngữ Mặc định và mã ngôn ngữ cần dịch

Ví dụ 1: Có 4 tham số: Từ để dịch - Mã ngôn ngữ nguồn - Mã ngôn ngữ đích - Cách đọc (nếu có)(Bỏ trống ->False)
=GoogleTranslate("Hello","en","vi", False)
=GoogleTranslate("Hello","Anh","Việt", False)
Kết quả: "Xin chào"
Ví dụ 2: Biến thứ hai để trống thì ngôn ngữ phát hiện và dịch tự động
=GoogleTranslate("xin chào", ,"zh-cn")
=GoogleTranslate("xin chào", ,"Trung")
Kết quả: "你好"
Biến thứ 3 là ngôn ngữ cần dịch để trống thì mặc định là tiếng Việt
Ví dụ 3: Biến thứ tư là True thì lấy cách đọc của từ đã được dịch (nếu có - thường là chữ tượng hình)
=GoogleTranslate("xin chào", ,"zh-cn", True)
"Nǐ hǎo"
*Lưu ý: Nếu File tạo để sử dụng trên Google Spreadsheet thì không nên điền tham số thứ 4.
=GoogleTranslate("Hello","en","vi")


=GoogleTranslate("Hello","en","vi", False , True, False, " -_/")
Được thêm 3 tham số sau cùng dựa trên hàm DetachText bao gồm:


+ Tham số 6 - hDetach: Cho phép tách chuỗi liên tục có dấu phân cách hoặc Chữ In hoa, mặc định là False.
+ Tham số 7 - hSpecial: Cho phép thêm dấu cách vào khi gặp ký tự đặc biệt, mặc định là False.
+ Tham số 8 - CharRemove: Nhập các ký tự cần phân cách, mặc định là " -_".

- Hàm DetachText: sẽ tách chuỗi liên tục được ngăn cách bởi ký tự khác dấu " " (dấu cách) , hoặc Âm viết hoa.
Hàm này sẽ giúp dịch các hàm , các comment trong Cửa sổ lập trình VBE.

2. Phát hiện ngôn ngữ :

- Hàm GoogleDetectLang (mới cập nhật):

=GoogleDetectLang("Hello")
+ Tham số - Chuỗi cần nhập để phát hiện ngôn ngữ

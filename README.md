# TranslateXL v2.49 - Add-in Dịch ngôn ngữ cho Excel
 Hàm dịch và phát hiện đa ngôn ngữ

-------------------------------------------------------------------------
[Nhấn tải TranslateXL](https://github.com/SanbiVN/TranslateXL/releases/download/TranslateXL/TranslateXL_v2.49.zip)  

[![Tổng tải xuống](https://img.shields.io/github/downloads/SanbiVN/TranslateXL/total.svg)](https://github.com/SanbiVN/TranslateXL/releases/download/TranslateXL/TranslateXL_v2.49.zip)

Hàm dưới đây sẽ giúp dịch thuật và phát hiện ngôn ngữ cho Office và VBA Editor
Code VBA hoạt động yêu cầu có Internet để dịch thông qua Google Translate​
Dịch ra nhiều thứ tiếng và nhiều thứ tiếng ra tiếng Việt​
Ưu điểm: Khi viết code hoặc copy code tham khảo qua mạng, thường sẽ xuất hiện những thuật ngữ mới, vậy nên cần đến dịch thuật.
Với dữ liệu Excel thì khá nhiều ngôn ngữ nên việc dịch thuật là đương nhiên.



![translatexl](https://github.com/SanbiVN/TranslateXL/assets/58664571/85c4cbcb-ab36-4e76-a59f-275b5f493299)


### Khi tải về gồm tệp .xlsm và add-in .xlam:
Tệp xlsm để dịch trực tiếp, hoặc sao chép mã sang dự án mới.​
Add-in .xlam để dịch nhanh với hàm và phím tắt​
​
## HƯỚNG DẪN CÀI ĐẶT ADD-IN

Sau khi tải Add-in với đuôi .xlam, nếu đuôi .xlsm thì hãy mở với Excel và lưu thành đuôi .xlam​
lưu vào thư mục phù hợp. Hãy cài đặt một trong hai cách dưới đây:​
Cách 1: Mở trình quản lý Add-in, cửa sổ hiện lên, chọn Browser..., tìm đến thư mục.​
Cách 2: Mở thư mục XLStart và tạo Shortcut, sau khi thư mục mở lên, nhấn chuột​
phải vào thư mục -> chọn New (Mới) -> chọn Shortcut, cửa sổ hiện lên chọn Browser....., tìm đến thư mục.​
​
## HƯỚNG DẪN SỬ DỤNG DỊCH NGỮ
​
### Cách xem danh sách ngôn ngữ và ID, bằng cách gõ hàm: ```=TranslateLanguages()​```
Sau khi gõ danh sách sẽ được in ra ô Excel​
​
### Dịch ngôn ngữ bằng cách gõ hàm:​
Với hàm `Translate` và hàm `Translate2`, với `Translate2` sau khi dịch thì ô gõ hàm sẽ được xóa đi chỉ còn lại từ đã được dịch. Nếu hàm Translate chạy trong Add-in thì sẽ tương tự.​
​
Hàm `Translate` có thể dịch Chuỗi, mảng, hoặc cả vùng ô.​
(*Lưu ý: nếu sử dụng Add-in, ô gõ hàm sẽ tự động được xóa như `Translate2`)​

### Có 3 tham số cơ bản, và Có 4 tham số bổ trợ:​
1. **Source** - Từ ngữ dịch​
2. **FromLanguage** - Ngôn ngữ nguồn​
3. **ToLanguage** - Mã ngôn ngữ đích​
4. **SkipOnlyAlphabets** - Bỏ qua chuỗi chỉ gồm ký tự Aphabets - Ascii​
5. **hDetach** - Tách các từ nối nhau (ví dụ: HelloWorldVietNam)​
6. **hSpecial** - Các ký tự đặc biệt​
7. **RemoveCharacters** - Nhập các ký tự cần bỏ qua trước khi dịch: "-_*&"​
​
Ví dụ 1:​
```=Translate("Hello","en","vi")​```
```=Translate("Hello","Anh","Việt")​```

Ví dụ 2: Đối số thứ hai để trống thì ngôn ngữ phát hiện và dịch tự động nhận biết ngôn ngữ nguồn​
​
```=Translate("xin chào", ,"zh-cn")​```
```=Translate("xin chào", ,"Trung")​```

Đối số thứ 3 là ngôn ngữ cần dịch để trống thì mặc định là tiếng Việt​
Hàm LanguageID tìm id của ngôn ngữ:​
```VBA=LanguageId("Trung")​```
```VBA=LanguageId("Việt")​```
​
Lưu ý: khi gõ hàm dịch bị xóa đi, hay gõ lại ngay vị trí ô trước đó ```=Translate()``` để gọi lại

Hàm DetectLang phát hiện ngôn ngữ: ```=DetectLang("Hello")​```
​
### Dịch ngữ sử dụng phím tắt `CTRL+ALT+T​`
​
Chọn một vùng ô cần dịch và nhấn phím tắt. Sau khi nhấn sẽ có thông báo hỏi​
"Bạn có muốn kết quả dịch trả về vị trí mới?", chọn Xác nhận​
Sau khi dịch, để trả lại các từ ngữ ban đầu hãy nhấn Undo hoặc nhấn CTRL+Z​
​
### Cài đặt mặc định:​
1. Ngôn ngữ nguồn: Auto - Tự động nhận dạng ngôn ngữ.​
2. Mã ngôn ngữ đích: Vi - Tiếng Việt​
3. Bỏ qua từ chỉ có ký tự Ascii: 0 (*Khi dịch từ tiếng Việt, hoặc ngôn ngữ tượng hình sang ngôn ngữ khác)​
4. Tách từ nối liền: 0​
​
Hãy sử dụng hàm cài đặt sau để cài đặt:​
```=TranslateSet("auto","vi", 1, 1)​```
​
- Đổi phím tắt mặc định gõ hàm: ```=TranslateSetKeys("^+%r")​```. Trong đó ^CTRL, +SHIFT , %ALT ​
- Nếu ký tự R viết hoa, có nghĩa là có SHIFT, tương đương +r​
- Các phím đặc biệt cần có cặp ngoặc nhọn ví dụ phím Home ```=TranslateSetKeys("^+{HOME}")​```
​
### Phiên bản cập nhật:​
- Trình tự động tìm kiếm bản cập nhật mới nhất tại Github​
- Để tắt gõ hàm: ```=TranslateUpdateOff()​```
- Để bật gõ hàm: ```=TranslateUpdateOn()```

## [Language support](https://cloud.google.com/translate/docs/languages)

Language|ISO-639 code
------- | -------
Afrikaans|af
Albanian|sq
Amharic|am
Arabic|ar
Armenian|hy
Assamese|as
Aymara|ay
Azerbaijani|az
Bambara|bm
Basque|eu
Belarusian|be
Bengali|bn
Bhojpuri|bho
Bosnian|bs
Bulgarian|bg
Catalan|ca
Cebuano|ceb
Chinese (Simplified)|zh-CN or zh (BCP-47)
Chinese (Traditional)|zh-TW (BCP-47)
Corsican|co
Croatian|hr
Czech|cs
Danish|da
Dhivehi|dv
Dogri|doi
Dutch|nl
English|en
Esperanto|eo
Estonian|et
Ewe|ee
Filipino (Tagalog)|fil
Finnish|fi
French|fr
Frisian|fy
Galician|gl
Georgian|ka
German|de
Greek|el
Guarani|gn
Gujarati|gu
Haitian Creole|ht
Hausa|ha
Hawaiian|haw
Hebrew|he or iw
Hindi|hi
Hmong|hmn
Hungarian|hu
Icelandic|is
Igbo|ig
Ilocano|ilo
Indonesian|id
Irish|ga
Italian|it
Japanese|ja
Javanese|jv or jw
Kannada|kn
Kazakh|kk
Khmer|km
Kinyarwanda|rw
Konkani|gom
Korean|ko
Krio|kri
Kurdish|ku
Kurdish (Sorani)|ckb
Kyrgyz|ky
Lao|lo
Latin|la
Latvian|lv
Lingala|ln
Lithuanian|lt
Luganda|lg
Luxembourgish|lb
Macedonian|mk
Maithili|mai
Malagasy|mg
Malay|ms
Malayalam|ml
Maltese|mt
Maori|mi
Marathi|mr
Meiteilon (Manipuri)|mni-Mtei
Mizo|lus
Mongolian|mn
Myanmar (Burmese)|my
Nepali|ne
Norwegian|no
Nyanja (Chichewa)|ny
Odia (Oriya)|or
Oromo|om
Pashto|ps
Persian|fa
Polish|pl
Portuguese (Portugal, Brazil)|pt
Punjabi|pa
Quechua|qu
Romanian|ro
Russian|ru
Samoan|sm
Sanskrit|sa
Scots Gaelic|gd
Sepedi|nso
Serbian|sr
Sesotho|st
Shona|sn
Sindhi|sd
Sinhala (Sinhalese)|si
Slovak|sk
Slovenian|sl
Somali|so
Spanish|es
Sundanese|su
Swahili|sw
Swedish|sv
Tagalog (Filipino)|tl
Tajik|tg
Tamil|ta
Tatar|tt
Telugu|te
Thai|th
Tigrinya|ti
Tsonga|ts
Turkish|tr
Turkmen|tk
Twi (Akan)|ak
Ukrainian|uk
Urdu|ur
Uyghur|ug
Uzbek|uz
Vietnamese|vi
Welsh|cy
Xhosa|xh
Yiddish|yi
Yoruba|yo
Zulu|zu

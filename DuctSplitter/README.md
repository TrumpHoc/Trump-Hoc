# TD.DuctSplitter – Add-in Revit chia đốt ống gió + dim Union

Ribbon: tab **TD MEP** › panel **Ống gió**

| Nút | Chức năng |
|---|---|
| **Chia đốt ống gió** | Chia hàng loạt ống gió (Duct) theo chiều dài đốt nhập trước, nối bằng **Union** theo Routing Preferences của kiểu ống (tương đương lệnh *Split Duct* trên giao diện Revit), sau đó **tự động dim các Union** trong view mặt bằng đang mở. Nếu view đang mở là **3D** thì chỉ chia, **không dim**. |
| **Dim Union** | Dim lại các Union trên các tuyến ống đã chia sẵn (chọn ống trong mặt bằng/mặt cắt). |

## 1. Hộp thoại "Chia đốt ống gió"

| Thông số | Mặc định | Ghi chú |
|---|---|---|
| Ống chữ nhật / tròn / oval (mm) | 1200 / 1200 / 1200 | Chiều dài 1 đốt, **đo theo tim ống** (LocationCurve), giữa 2 connector của ống |
| Đốt cuối tối thiểu (mm) | 100 | Phần dư cuối < giá trị này thì **gộp vào đốt trước** (không cắt mẩu vụn) |
| Bắt đầu đo từ | Tự động | *Tự động*: đo từ đầu đang nối fitting/thiết bị (nếu chỉ 1 đầu nối); nếu cả 2 hoặc không đầu nào nối → từ điểm Start của ống |
| Phạm vi | Ống đang chọn | Hoặc toàn bộ ống gió trong view hiện hành. Ống mềm (Flex Duct) không xử lý |
| Tự động dim Union | Bật | Bị khoá khi view hiện hành là 3D / sheet / schedule… |
| Dim cả 2 đầu tuyến | Bật | Chuỗi dim: đầu ống – Union – … – Union – cuối ống |
| Khoảng cách dim (mm giấy) | 8 | Từ mép ống tới đường dim = ½ bề rộng ống + giá trị × tỉ lệ view |
| Kiểu dim | Mặc định dự án | Danh sách các Linear Dimension Type trong dự án |

Thông số được lưu lại cho lần sau tại `%APPDATA%\TD.DuctSplitter\settings.txt`.

**Ví dụ kiểm tra số:** ống dài 5000 mm, đốt 1200, đốt cuối tối thiểu 100 → điểm chia tại 1200 / 2400 / 3600 / 4800; phần dư 200 ≥ 100 nên giữ → 5 đoạn: 4×1200 + 200 = 5000 ✔.
Nếu ống dài 4850 → điểm 4800 cho phần dư 50 < 100 → bỏ → 4 đoạn: 3×1200 + 1250 = 4850 ✔.

## 2. Cách dim hoạt động

1. Với mỗi tuyến đã chia, add-in lấy tham chiếu **mặt phẳng tâm của family Union** vuông góc trục ống
   (Center Left/Right, Front/Back hoặc Elevation – tự chọn theo hướng family).
2. Nếu family Union không có mặt phẳng tham chiếu phù hợp → bám **mặt đầu ống** tại mối nối.
3. Nếu Revit từ chối tạo dim → thử lại toàn bộ bằng mặt đầu ống.
4. Bỏ qua dim với: ống đứng / ống dốc (không song song mặt phẳng view), tuyến không hiển thị đủ trong view.

## 3. Điều kiện để có Union

Kiểu ống (Duct Type) › **Routing Preferences** › mục **Unions** phải có family Union (ví dụ bích nối TDC).
Nếu thiếu, add-in vẫn chia ống nhưng **nối trực tiếp** 2 đoạn (không có Union) và báo trong bảng kết quả; dim vẫn tạo được, bám mặt đầu ống.

## 4. Build

Không cần cài Revit để build (dùng gói NuGet `Nice3point.Revit.Api.*` – reference assembly).

```powershell
dotnet build DuctSplitter\DuctSplitter.csproj -c Release -p:RevitVersion=2024
```

| Revit | Framework | SDK cần có |
|---|---|---|
| 2022 – 2024 | .NET Framework 4.8 | .NET SDK 8 |
| 2025 – 2026 | .NET 8 (`net8.0-windows`) | .NET SDK 8, build trên Windows |
| 2027 | .NET 10 (`net10.0-windows`) | .NET SDK 10, build trên Windows |

Kết quả: `DuctSplitter\bin\Release\Revit<năm>\`.
GitHub Actions (`.github/workflows/build-ductsplitter.yml`) tự build đủ 2022–2027 và đóng gói sẵn ở mục *Artifacts*.

## 5. Cài đặt

1. Đóng Revit.
2. Copy vào `%APPDATA%\Autodesk\Revit\Addins\<năm>\`:
   ```
   TD.DuctSplitter.addin
   TD.DuctSplitter\TD.DuctSplitter.dll
   ```
3. Mở Revit → chọn **Always Load** khi được hỏi → xuất hiện tab **TD MEP**.

## 6. Giới hạn đã biết

- Chiều dài đốt tính theo tim ống giữa 2 connector; nếu family Union có chiều dài (khác 0) thì Revit sẽ cắt bớt ống tương ứng – đốt thực tế ngắn hơn đúng bằng chiều dài Union.
- Điểm chia không tránh vị trí tap/rẽ nhánh trên thân ống (takeoff) – nên kiểm tra lại các ống có tap.
- Ống bị **Pin** sẽ bị bỏ qua. Với file Workshared, cần quyền sở hữu (borrow) ống.
- Toàn bộ thao tác gom trong 1 lần **Undo**.

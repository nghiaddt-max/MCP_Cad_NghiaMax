# MCP_Cad_NghiaMax — Kết nối Claude với AutoCAD

MCP server cho phép Claude đọc và vẽ trên bản vẽ đang mở trong AutoCAD, thông qua giao diện COM (ActiveX) của Windows.

Phiên bản 0.2.1 — RIKCO JSC.

> Tên repo và gói phát hành là **MCP_Cad_NghiaMax**. Bên trong vẫn giữ các tên kỹ thuật cũ để máy đã cài không phải đăng ký lại:
> - gói Python `autocad_mcp`, lệnh `autocad-mcp.exe`;
> - tên server khi đăng ký với Claude: `autocad`.

## 1. Yêu cầu

| Thành phần | Yêu cầu |
|---|---|
| Hệ điều hành | Windows 10/11 64-bit |
| AutoCAD | Bản **full**: AutoCAD, AutoCAD MEP, Civil 3D… (đã thử trên AutoCAD 2024). **AutoCAD LT không chạy được** vì LT không có COM. |
| Python | 3.10 trở lên, tải tại [python.org](https://www.python.org/downloads/windows/) |
| Internet | Cần khi cài lần đầu để tải gói `mcp` và `pywin32`. Máy không có mạng: xem mục 6. |
| Claude | Claude Code (CLI, ứng dụng desktop, tab Code) hoặc Claude Desktop |

## 2. Các tool có sẵn

| Tool | Chức năng | Ghi vào bản vẽ? |
|---|---|---|
| `drawing_info` | Tên, đường dẫn, layer hiện hành, số đối tượng | Không |
| `list_layers` | Danh sách layer với màu, linetype, trạng thái on/freeze/lock | Không |
| `count_blocks` | Đếm block theo tên. Lọc được theo layer, đếm cả block lồng, bỏ qua block trên layer đang tắt/freeze. MINSERT tính đủ số hàng × cột; ARRAY liên kết (block ẩn danh `*U…`) được đếm theo các block bên trong, không đếm chính nó | Không |
| `create_layer` | Tạo layer hoặc đổi màu layer | **Có** |
| `draw_line`, `draw_circle`, `add_text` | Vẽ line, circle, text trong ModelSpace. Layer chỉ định phải có sẵn; nếu chưa có thì báo lỗi và không vẽ gì | **Có** |

Các tool ghi sẽ vẽ thẳng vào bản vẽ đang mở. Với bản vẽ gốc của CĐT hoặc NCC, hãy mở ở chế độ **Read-only**, hoặc lưu bản sao trước khi cho Claude vẽ.

## 3. Cài đặt nhanh

**Bước 1 — Cài Python.** Chạy file cài từ python.org và tick **"Add python.exe to PATH"**. Kiểm tra bằng PowerShell:

```powershell
py --version
```

**Bước 2 — Giải nén và chạy script cài đặt.** Tải `MCP_Cad_NghiaMax-0.2.1-windows.zip` ở mục **Releases** của repo GitHub, giải nén vào một thư mục bất kỳ, mở PowerShell trong thư mục đó rồi chạy:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

Script sẽ:
- tạo môi trường riêng trong `%LOCALAPPDATA%\autocad-mcp\.venv`, không ảnh hưởng Python khác trên máy;
- cài gói và lệnh `autocad-mcp.exe`;
- tạo file cấu hình mẫu `%LOCALAPPDATA%\autocad-mcp\mcp-config.json`.

Cuối cùng script in ra đường dẫn đầy đủ của `autocad-mcp.exe`. Ghi lại đường dẫn này để dùng ở bước 3.

Có thể đổi chỗ cài bằng tham số `-InstallDir "D:\Tools\autocad-mcp"`. Thêm `-Force` để xóa bản cũ rồi cài lại từ đầu.

**Bước 3 — Kiểm tra kết nối.** Mở AutoCAD và mở một bản vẽ bất kỳ, rồi chạy (thay `<User>` bằng tên tài khoản Windows):

```powershell
& "C:\Users\<User>\AppData\Local\autocad-mcp\.venv\Scripts\autocad-mcp.exe" --check
```

Nếu kết nối được, dòng cuối hiện `KET NOI OK` cùng tên bản vẽ đang mở.

## 4. Đăng ký với Claude

Trong các ví dụ dưới đây, `<EXE>` là đường dẫn đầy đủ tới `autocad-mcp.exe` ở bước 2. Trong file JSON, mỗi dấu `\` phải viết thành `\\`.

### 4.1. Claude Code — dùng cho mọi dự án (khuyên dùng)

```powershell
claude mcp add --scope user autocad -- "<EXE>"
```

### 4.2. Claude Code — chỉ dùng cho một dự án

Tạo hoặc bổ sung file `.mcp.json` ở thư mục gốc của dự án:

```json
{
  "mcpServers": {
    "autocad": {
      "command": "C:\\Users\\<User>\\AppData\\Local\\autocad-mcp\\.venv\\Scripts\\autocad-mcp.exe",
      "args": []
    }
  }
}
```

Mở lại phiên Claude Code trong thư mục dự án. Khi được hỏi có cho phép server `autocad` chạy không, chọn đồng ý.

### 4.3. Claude Desktop (tab Chat)

Trong Claude Desktop, vào **Settings → Developer → Edit Config** để mở đúng file cấu hình. Bản cài thường dùng `%APPDATA%\Claude\claude_desktop_config.json`. Bản cài từ Microsoft Store có thể lưu file ở chỗ khác, chưa xác minh. Thêm khối `"autocad"` ở trên vào trong `"mcpServers"`; nếu file chưa có thì tạo mới với nội dung như mục 4.2. Thoát hẳn Claude Desktop (kể cả biểu tượng ở khay hệ thống) rồi mở lại.

### 4.4. Thử

Mở một bản vẽ trong AutoCAD rồi gõ cho Claude: *"Kiểm tra bản vẽ đang mở trong AutoCAD"*. Claude sẽ gọi `drawing_info` và trả về tên bản vẽ.

## 5. Xử lý sự cố

| Hiện tượng | Nguyên nhân / cách xử lý |
|---|---|
| `Khong tim thay AutoCAD dang mo` | AutoCAD chưa mở, hoặc đang dùng LT. Nếu AutoCAD chạy bằng *Run as administrator* thì Claude cũng phải chạy cùng quyền, hoặc mở AutoCAD bình thường (không admin). |
| `AutoCAD chua mo ban ve nao` | Mở hoặc tạo một bản vẽ. Tab *Start* không tính là bản vẽ. |
| `Call was rejected by callee` | AutoCAD đang bận: đang chạy lệnh dở, đang mở hộp thoại, hoặc đang tải bản vẽ. Bấm vào vùng vẽ, nhấn **Esc** hai lần rồi thử lại. |
| Script bị chặn: *running scripts is disabled* | Chạy đúng lệnh có `-ExecutionPolicy Bypass` như ở bước 2. |
| `pip install that bai` | Máy không có Internet hoặc mạng công ty chặn PyPI. Cài offline theo mục 6. |
| Claude không thấy tool `autocad` | Kiểm tra lại đường dẫn `<EXE>` và dấu `\\` trong JSON, rồi khởi động lại Claude. Trong Claude Code, gõ `/mcp` để xem trạng thái server. |
| `count_blocks` chạy lâu | Bản vẽ lớn có nhiều block lồng. Dùng `include_nested` chỉ khi cần đếm cả block lồng. |
| Mở nhiều phiên AutoCAD cùng lúc | Server kết nối vào phiên AutoCAD được Windows đăng ký trước. Nên chỉ mở một phiên. |

## 6. Cài trên máy không có Internet

Trên một máy có Internet, cài **cùng phiên bản Python** với máy đích (ví dụ cùng 3.12 64-bit). Trong thư mục giải nén, chạy:

```powershell
py -m pip download .\autocad_mcp-0.2.1-py3-none-any.whl -d wheelhouse
```

Chép cả thư mục (gồm `wheelhouse`) sang máy đích, rồi cài:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1 -Wheelhouse .\wheelhouse
```

## 7. Cập nhật và gỡ cài đặt

- **Cập nhật:** giải nén bản mới, chạy lại `install.ps1`. Đường dẫn `<EXE>` không đổi nên không cần đăng ký lại.
- **Gỡ cài đặt:**
  1. Bỏ đăng ký: chạy `claude mcp remove --scope user autocad`, hoặc xóa khối `"autocad"` trong file JSON.
  2. Xóa thư mục `%LOCALAPPDATA%\autocad-mcp`.

## 8. Dành cho người phát triển

```
MCP_Cad_NghiaMax/
├─ autocad_mcp/        mã nguồn gói (server.py chứa các tool)
├─ server.py           lối vào cũ, giữ tương thích .mcp.json đang dùng (python server.py)
├─ pyproject.toml      khai báo gói, phụ thuộc mcp 2.x + pywin32, lệnh autocad-mcp
├─ install.ps1         script cài trên máy khác
├─ build_dist.ps1      đóng gói ra dist\MCP_Cad_NghiaMax-<version>-windows.zip
├─ scripts/            helper cho script COM viết tay
│  ├─ acadlib.py       kết nối, tự chờ khi AutoCAD bận, vẽ, block, layout, viewport
│  └─ build_khung_a3.py  dựng lại file mẫu khung tên A3
└─ templates/
   └─ RIK-KHUNG-A3.dwg khung A3 minh họa (13 attribute); dùng thực tế thì truyền khung chuẩn qua template=
```

Gói zip phát hành gồm wheel, `install.ps1`, README, `scripts/`, `templates/` và `SHA256SUMS.txt`. Thư mục `dist/` và `.venv/` không commit vào repo; file zip đưa lên mục GitHub Releases.

### Chạy từ mã nguồn (clone từ GitHub)

```powershell
git clone https://github.com/<tai-khoan>/MCP_Cad_NghiaMax.git
```

```powershell
cd MCP_Cad_NghiaMax; py -m venv .venv; .\.venv\Scripts\python.exe -m pip install -e .
```

```powershell
.\.venv\Scripts\autocad-mcp.exe --check
```

Đăng ký với Claude: dùng lệnh `<repo>\.venv\Scripts\autocad-mcp.exe` như ở mục 4. Với cài dạng `-e`, sửa mã nguồn xong chỉ cần kết nối lại MCP, không phải cài lại.

### Script COM viết tay (`scripts/acadlib.py`)

Dùng cho những việc mà tool MCP chưa làm được, ví dụ dựng layout, viewport hay vẽ hàng loạt. Không cần đặt `PYTHONIOENCODING`, vì thư viện tự chuyển console sang UTF-8.

Phải chạy bằng Python của môi trường đã cài gói, vì môi trường đó có sẵn `pywin32`:

| Cách cài | Python dùng để chạy script | Thư mục `scripts\` |
|---|---|---|
| Từ zip bằng `install.ps1` | `%LOCALAPPDATA%\autocad-mcp\.venv\Scripts\python.exe`, hoặc `<InstallDir>\.venv\...` nếu đã đổi chỗ cài | trong thư mục giải nén zip |
| Từ mã nguồn (`pip install -e .`) | `<repo>\.venv\Scripts\python.exe` | `<repo>\scripts` |

```powershell
& "$env:LOCALAPPDATA\autocad-mcp\.venv\Scripts\python.exe" my_script.py
```

```python
import sys; sys.path.insert(0, r"<thu-muc>\scripts")  # thư mục repo hoặc thư mục giải nén zip
import acadlib as A
app, doc = A.connect("ten-ban-ve.dwg")      # sai tên bản vẽ thì báo lỗi, không vẽ nhầm
with A.undo_group(doc):                       # gõ U một lần là hoàn tác cả đợt
    info = A.setup_sheet_a3(app, doc, "Layout1", "DN-01",
                            {"DU_AN": "...", "TEN_BAN_VE": "...", "TY_LE": "1:100", "SO_BAN_VE": "DN-01"})
print(info)                                   # đọc lại kết quả để kiểm tra, không xuất file
```

`setup_sheet_a3` thực hiện lần lượt:
- nạp block `RIK-KHUNG-A3` từ file mẫu;
- đặt page setup: DWG To PDF, A3 full bleed, tỷ lệ 1:1, `monochrome.ctb`;
- chèn khung tên tại 0,0;
- tạo viewport cắt chừa góc khung tên. Viewport khóa và nằm trên layer `A-VIEWPORT` (không in).

Gọi lại lần nữa thì hàm chỉ cập nhật attribute, không chèn thêm khung hay viewport. Thông số `scale` và `model_center` dùng khi Model vẽ theo tỷ lệ thật, ví dụ 1:100 thì `scale=0.01`.

Những lỗi COM thư viện đã xử lý sẵn:
- **AutoCAD bận** (lỗi *Call was rejected by callee* hoặc `AttributeError`): một message filter cho mọi lời gọi tự thử lại.
- **Sysvar kiểu short** phải truyền `VT_I2`: dùng `set_var`.
- **`SendCommand` chạy bất đồng bộ**: dùng `send`, hàm này chờ `CMDNAMES` rỗng.
- **Tỷ lệ viewport**: thư viện dùng `CustomScale`, vì `StandardScale` của viewport theo một bảng mã khác với của layout.

**Sửa mẫu khung tên:** sửa hàm `draw()` trong `scripts/build_khung_a3.py`, rồi chạy lại:

```powershell
& ".\.venv\Scripts\python.exe" scripts\build_khung_a3.py
```

Script mở một bản vẽ trắng tạm thời để vẽ và lưu đè `templates/RIK-KHUNG-A3.dwg`, rồi đóng bản vẽ đó. Bản vẽ đang mở của người dùng không bị đụng tới. Bản vẽ nào đã nạp block cũ thì phải xóa định nghĩa block cũ trước khi nạp lại.

- **Thêm tool:** viết hàm trong `autocad_mcp/server.py` và gắn decorator `@acad_tool`.
- **Báo lỗi cho Claude:** chỉ `ToolError` được chuyển nguyên văn tới Claude.
- **COM:** mỗi lần gọi COM phải qua `_doc()`, vì hàm này gọi `CoInitialize` cho luồng xử lý (worker thread).
- **Phát hành bản mới:** tăng `__version__` trong `autocad_mcp/__init__.py` và `version` trong `pyproject.toml`, rồi chạy:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_dist.ps1
```

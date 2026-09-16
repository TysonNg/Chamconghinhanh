# Kế hoạch sửa cache, ngày ảnh và định danh nhân viên

Ngày soạn: 16/09/2026. Trạng thái: đã triển khai ba sửa đổi chính và kiểm thử tự động; xem ghi nhận bàn giao cuối tài liệu.

> Khi triển khai: dùng superpowers:executing-plans theo từng hạng mục, viết kiểm thử tái hiện lỗi trước khi sửa. Mỗi hạng mục phải có kết quả kiểm thử và phần thay đổi riêng để rà soát.

**Mục tiêu:** Quét khuôn mặt không tự khóa; chọn ảnh đúng dự án và đầy đủ ngày-tháng-năm; ghép chân dung theo định danh đã xác nhận.

**Thiết kế:** Giữ bộ nhận diện hiện tại. Tách quy tắc ngày và sổ định danh thành hai module nhỏ, truyền ngữ cảnh dự án/nhân viên xuyên suốt API → bộ phân tích → xuất báo cáo → FaceMatcher. Dữ liệu thiếu hoặc mâu thuẫn được trả về trạng thái cần kiểm tra.

**Công nghệ:** Python hiện có của dự án, threading, pathlib, datetime, sqlite3, pytest; JavaScript và node:test cho Zalo/giao diện. Không thêm dịch vụ hay thư viện runtime cho ba thay đổi này.

**Đặc tả đầu vào:** Ba mục người dùng yêu cầu ngày 16/09/2026 trong cuộc hội thoại này. Các quy tắc bên dưới là phương án đề xuất cụ thể cho việc triển khai.

## Các quyết định chung

- Sửa cache độc lập trước. Sau đó chuẩn hóa ngày, rồi triển khai định danh.
- Không đổi model nhận diện, ngưỡng so sánh hay thuật toán OCR trong đợt này.
- Chỉ dùng dữ liệu thử trong thư mục tạm; mock DeepFace và nhà cung cấp AI. Không gọi API trả phí hoặc dùng phiên Zalo thật để chạy unit test.
- Giữ nguyên ảnh gốc. Không chỉnh watermark/EXIF để làm khớp ngày hồ sơ.
- Không suy tháng/năm từ ngày chạy, kỳ báo cáo đang mở hoặc ngày người dùng muốn bổ sung.
- Không suy danh tính từ tên gần giống. Tên hiển thị và đường dẫn không phải mã định danh.
- Không đổi tên/di chuyển hàng loạt thư mục cũ khi khởi động. Chuyển đổi dữ liệu là bước riêng, có bản xem trước, bản sao lưu và manifest.
- Kho mã đang có nhiều thay đổi chưa commit. Khi triển khai phải kiểm tra lại diff, giữ thay đổi của người dùng, chỉ đưa các phần thuộc kế hoạch vào commit.
- Các đường dẫn dữ liệu cần hoạt động cả khi chạy Python lẫn bản EXE; xác định qua base directory hiện có, không hard-code đường dẫn máy.
- Ba phần được nghiệm thu độc lập nhưng chỉ bật toàn bộ luồng mới sau khi cập nhật cả bên ghi lẫn bên đọc.

## Phạm vi phát hiện thêm cần bao phủ

1. Cache: _CACHE_LOCK là threading.Lock; _get_embedding giữ khóa rồi gọi load_disk_cache/save_disk_cache, hai hàm cũng lấy khóa.
2. Ngày: ngoài ExcelToWordExporter, các API xem/upload/xóa ảnh, tạo dự án, hai bộ tải Zalo và apply_to_attendance đều có hành vi dùng DD.
3. Định danh: Excel và PDF đã trích xuất person['id'], nhưng chân dung quản lý bằng tên. Excel còn gộp hồ sơ theo tên; PDF có bước sửa tên qua tìm kiếm chân dung toàn cục.
4. Hai API phân tích Excel/PDF có fallback về thư mục gốc khi dự án không tồn tại.
5. Tên file báo cáo cá nhân hiện chỉ gồm tên-tháng-năm, có thể ghi đè khi hai người trùng tên.
6. Bộ xuất báo cáo tổng hợp dùng matched_image_path, nên cần bảo đảm chỉ nhận ảnh đủ điều kiện liên kết từ luồng mới.

## Hạng mục 1 — Loại bỏ khóa lồng nhau trong cache

**Sửa:** src/face_matcher.py.
**Tạo kiểm thử:** tests/test_face_matcher_cache.py.

### Phương án

Giữ threading.Lock và chia rõ hàm tự quản lý khóa với helper chỉ được gọi khi đã giữ khóa. Đây là phương án chính; đổi sang RLock là bản vá nhỏ hơn nhưng vẫn giữ cấu trúc gọi lồng khó kiểm soát, nên không chọn làm thiết kế cuối.

Giao diện nội bộ:

~~~python
def _load_disk_cache_locked() -> dict:
    """Caller must hold _CACHE_LOCK."""

def _save_disk_cache_locked(force: bool = False) -> None:
    """Caller must hold _CACHE_LOCK."""

def load_disk_cache() -> dict:
    with _CACHE_LOCK:
        return _load_disk_cache_locked()

def save_disk_cache(force: bool = False) -> None:
    with _CACHE_LOCK:
        _save_disk_cache_locked(force=force)
~~~

- _get_embedding gọi helper tương ứng khi đã giữ khóa; không gọi wrapper lấy khóa lần nữa.
- Nạp cache một lần trong vùng khóa. Việc tính embedding/DeepFace nằm ngoài vùng khóa.
- Kiểm tra cache lần nữa khi ghi sau khi tính embedding, để hai worker cùng ảnh không ghi đếm thay đổi hai lần.
- Khi đến ngưỡng ghi hiện tại là 20 thay đổi, gọi helper lưu trong vùng khóa.
- Ghi file tạm cùng thư mục, đóng file rồi os.replace; chỉ reset bộ đếm dirty khi thay thế thành công.
- Lỗi ghi giữ nguyên cache RAM và trạng thái dirty để lần sau thử lại; lỗi nạp được ghi log rồi khởi tạo cache rỗng.
- Giữ khóa trong bước ghi đĩa ở bản sửa này để đơn giản hóa tính nhất quán; chưa thêm writer thread.

### Các bước và kiểm thử

- [x] Viết test cold start: _DISK_CACHE=None, mock DeepFace trả embedding hợp lệ, gọi _get_embedding.
- [x] Viết test chạm ngưỡng: cache đã nạp, dirty_count=19, thêm embedding thứ 20.
- [x] Chạy hai test trong subprocess với timeout 5 giây; bản cũ phải tái hiện timeout. Không dùng ThreadPoolExecutor context manager để chờ một thread có thể deadlock vô hạn.
- [x] Cài đặt helper và đổi các điểm gọi.
- [x] Test cache hit không gọi DeepFace; nhiều worker truy cập cache; cache file hỏng; lỗi os.replace vẫn giữ file cũ và dirty_count.
- [x] Với test đồng thời dùng barrier/event và embedding giả, không dùng sleep để tạo lịch chạy.
- [ ] Chạy riêng module kiểm thử, rồi các test hiện có liên quan FaceMatcher; rà soát và commit riêng phần cache.

**Đạt khi:** cả cold start và lần ghi thứ 20 hoàn tất trong timeout; không có callback DeepFace trong vùng khóa; file cache vẫn đọc được sau các lần ghi đồng thời.

## Hạng mục 2 — Quy tắc ngày đầy đủ và tương thích dữ liệu cũ

**Tạo:** src/attendance_dates.py; tests/test_attendance_dates.py.
**Sửa:** src/excel_extractor.py.

### Giao diện và quy tắc

~~~python
@dataclass(frozen=True)
class DayResolution:
    status: str  # found, missing, legacy_unmapped, ambiguous, invalid_date
    path: Path | None
    reason: str

def parse_attendance_date(value: str) -> date:
    """Chấp nhận chính xác YYYY-MM-DD hoặc DD/MM/YYYY; sai thì ValueError."""

def resolve_day_folder(
    project_dir: Path, attendance_date: date
) -> DayResolution:
    """Chỉ tìm trong project_dir; không tìm lên thư mục cha."""

def canonical_day_path(project_dir: Path, attendance_date: date) -> Path:
    """Trả project_dir / attendance_date.isoformat()."""

def compare_capture_date(
    attendance_date: date, observed_date: date | None
) -> str:
    """Trả consistent, mismatch hoặc unknown; không chứng thực ảnh thật."""
~~~

Quy tắc đọc:

- YYYY-MM-DD là tên thư mục chuẩn cho dữ liệu mới.
- DD-MM-YYYY được hỗ trợ đọc bằng parse ngày đầy đủ.
- Nếu có nhiều thư mục đại diện cùng ngày, trả ambiguous, không chọn theo thứ tự os.listdir.
- DD chỉ được đọc khi có mapping được xác nhận cho chính thư mục đó trong attendance-period.json đặt tại thư mục dự án.
- Manifest v1 lưu mappings theo tên thư mục, ngày ISO đầy đủ, người xác nhận và thời điểm xác nhận. Ví dụ mappings['01'].date = '2026-09-01'.
- Một thư mục DD có ảnh từ nhiều tháng không được gán cả thư mục cho một tháng. Bản xem trước chuyển đổi phải đánh dấu mixed/unknown và yêu cầu phân loại từng ảnh.
- Không có mapping: legacy_unmapped. Ngày không hợp lệ: invalid_date. Không có ảnh: missing.
- Adapter _find_day_folder dùng date_raw để parse ngày; không tiếp tục dùng day_str làm nguồn quyết định.

Kiểm thử tối thiểu:

~~~python
def test_does_not_select_same_day_in_another_month(tmp_path):
    (tmp_path / "2026-08-01").mkdir()
    result = resolve_day_folder(tmp_path, date(2026, 9, 1))
    assert result.status == "missing"
    assert result.path is None

def test_legacy_day_requires_explicit_mapping(tmp_path):
    (tmp_path / "01").mkdir()
    result = resolve_day_folder(tmp_path, date(2026, 9, 1))
    assert result.status == "legacy_unmapped"
    assert result.path is None
~~~

### Các bước

- [x] Viết và chạy test sai tháng, sai năm, giao năm, 29/02 hợp lệ/không hợp lệ.
- [x] Test DD-MM-YYYY đúng; DD chưa mapping; DD mapping sang tháng khác; nhiều thư mục cùng ngày; không fallback lên thư mục cha.
- [x] Cài đặt module ngày và thay adapter trong exporter.
- [x] Trả lý do thiếu/mơ hồ lên record; báo cáo hiển thị “Cần kiểm tra ngày ảnh” thay vì coi là không khớp mặt.
- [x] Khi có ngày quan sát trên ảnh gốc, đối chiếu với ngày hồ sơ; mismatch không được chọn làm ảnh xác nhận của ngày đó.
- [x] Không có OCR/EXIF gốc: unknown. Ngày gửi Zalo, tên file, tên thư mục và target_date không được gắn nhãn là ngày chụp đã xác minh.
- [x] Test mismatch/unknown/consistent bằng metadata giả, không gọi AI.

**Đạt khi:** ảnh tháng 8 không được gắn hồ sơ tháng 9 dù cùng ngày 01; mọi fallback DD phải có mapping rõ ràng.

## Hạng mục 3 — Đồng bộ các bên ghi/đọc ngày

**Sửa Python:** src/app.py, src/fake_photo_service.py, src/supplement_batches.py.
**Sửa JavaScript:** zalo-service/image-downloader.js, zalo-service/zalo-browser-downloader.js, zalo-service/server.js, static/js/app.js, static/js/supplement-batches.js.
**Sửa giao diện:** templates/index.html nếu có bộ chọn định dạng/ngày; templates/supplement_batches.html nếu cần trường ngày.
**Tạo test:** tests/test_daily_photo_dates.py, tests/test_supplement_attendance_dates.py, tests/test_daily_photo_dates.cjs.
**Cập nhật test:** zalo-service/tests/download.test.js và các test gọi API ngày cũ.

### Các bước

- [x] API upload/xem/xóa nhận date=YYYY-MM-DD; thống kê nhận period=YYYY-MM. Đầu vào chỉ có DD trả 400 nếu không có ngữ cảnh tháng/năm rõ ràng.
- [x] Giao diện luôn gửi ngày đầy đủ, hiển thị tháng/năm đang chọn; thao tác xóa giữ nguyên định danh ngày và đường dẫn ảnh, không xóa các tháng khác cùng DD.
- [x] ensure_project_structure chỉ tạo thư mục dự án, ngừng tạo mặc định 01..31.
- [x] Cả hai Zalo downloader ghi YYYY-MM-DD; cập nhật mặc định server/giao diện. Yêu cầu cũ chọn DD/day bị từ chối rõ ràng thay vì âm thầm ghi DD.
- [x] Ngày gửi Zalo dùng để phân thư mục nhưng giữ date_source='message'; không đổi nghĩa thành ngày chụp.
- [x] apply_to_attendance dùng canonical_day_path(target_date); bỏ startswith/endswith và không chọn thư mục bằng ngày cuối chuỗi.
- [x] Đối chiếu ngày từ ảnh gốc trước khi đưa vào luồng tự động; mismatch trả trạng thái cần kiểm tra, unknown không được báo đã xác minh.
- [ ] Lưu và hiển thị ngày đề nghị/ngày quan sát/trạng thái riêng, không dùng ngày sau chỉnh sửa để xác minh ảnh gốc.
- [x] Test tạo/xem/xóa hai thư mục 2026-08-01 và 2026-09-01: thao tác ngày nào chỉ tác động ngày đó.
- [x] Test tải Zalo qua ranh giới tháng/năm và ảnh gửi lại ngày khác; không suy ngày khi metadata thiếu.
- [x] Test staging đưa đúng thư mục ISO, không rơi vào thư mục DD cũ, mismatch không tự áp dụng.
- [x] Cập nhật assertion đường dẫn trong test download hiện có sang YYYY-MM-DD; không bỏ các kiểm thử ảnh gốc, lỗi tải, ảnh thiếu ngày.

**Đạt khi:** mọi đường ghi mới đều dùng ISO; phía xem, xóa, nhận diện và xuất báo cáo sử dụng cùng quy tắc ngày.

## Hạng mục 4 — Sổ định danh dự án/nhân viên và dữ liệu cũ

**Tạo:** src/identity_registry.py, tests/test_identity_registry.py.
**Dữ liệu runtime đề xuất:** data/identity.sqlite3; bổ sung ignore chính xác cho file DB và các sidecar.
**Sửa:** src/app.py, static/js/app.js, templates/index.html, .gitignore.

### Mô hình

- projects: project_id UUID bất biến, display_name, storage_dir duy nhất, active.
- employees: employee_id UUID bất biến, display_name, active.
- memberships: project_id, employee_id, payroll_code dạng chuỗi, valid_from, valid_to.
- portrait_bindings: project_id, employee_id, relative_path, confirmed_at, confirmed_by.
- Không tạo ID bằng tên chuẩn hóa hoặc đường dẫn. Đổi tên không đổi ID.
- Mã lấy từ person['id'] là payroll_code của nguồn chấm công, không ghi đè thành UUID.
- Payroll code phân biệt trong từng dự án tại một ngày hiệu lực; ứng dụng kiểm tra khoảng hiệu lực không chồng lấn trong transaction. Giữ nguyên số 0 đầu nếu nguồn có cung cấp; nguồn thiếu thông tin thì cần đối chiếu.
- Hồ sơ chưa có mã vẫn có UUID nội bộ, nhưng chỉ được tự ghép sau khi đã có liên kết nguồn được xác nhận.
- Không gộp nhân viên giữa hai dự án chỉ vì tên giống nhau.
- Mối liên kết chân dung phải nằm trong storage_dir của dự án đã resolve; đường dẫn client gửi vào không tự tạo quyền truy cập dự án.

Giao diện dùng ở hạng mục kế tiếp:

~~~python
@dataclass(frozen=True)
class IdentityResolution:
    status: str  # resolved, missing, ambiguous, needs_confirmation
    project_id: str
    employee_id: str | None
    reason: str

class IdentityRegistry:
    def resolve_employee(
        self, project_id: str, payroll_code: str, on_date: date
    ) -> IdentityResolution: ...

    def portrait_paths(
        self, project_id: str, employee_id: str, on_date: date
    ) -> list[Path]: ...
~~~

### Các bước

- [x] Viết test tạo ID ổn định qua restart/đổi tên; cùng tên khác mã; cùng mã ở dự án khác; mã có số 0 đầu; membership theo ngày.
- [x] Cài đặt schema và transaction; foreign keys bật ở từng connection.
- [x] Thêm API/giao diện quản lý mã chấm công và liên kết chân dung trong màn hình nhân viên hiện tại.
- [ ] Import thư mục cũ thành các bản ghi pending; tạo bản xem trước đối chiếu với mã từ Excel/PDF. Chỉ lưu binding chính thức sau khi người dùng xác nhận.
- [x] Chạy import lặp không sinh bản ghi mới cho cùng nguồn thư mục đã đăng ký.
- [x] API tạo/upload/chuyển nhân viên dùng project_id và employee_id; tên giữ để hiển thị.
- [x] Thuyên chuyển giữ employee_id, đóng/mở membership theo ngày hiệu lực, giữ lịch sử. Chưa biết ngày hiệu lực cũ thì không tự giả định.
- [x] Dữ liệu chân dung cũ được giữ; không gọi luồng cũ gộp/xóa thư mục đích theo tên. Liên kết lịch sử phải tiếp tục dùng được cho báo cáo kỳ cũ.
- [x] Cập nhật đổi tên/xóa dự án và nhân viên để bảo toàn ID, kiểm tra tham chiếu và không để binding mồ côi.

**Đạt khi:** hai người trùng tên có hai danh tính độc lập; không dùng fuzzy match để tạo binding; chuyển dự án không làm đổi danh tính hoặc mất khả năng tra cứu lịch sử đã xác nhận.

## Hạng mục 5 — Truyền định danh xuyên suốt nhận diện và báo cáo

**Sửa:** src/face_matcher.py, src/app.py, src/excel_extractor.py, src/excel_face_analyzer.py, src/pdf_face_analyzer.py, src/aggregate_report_exporter.py.
**Tạo test:** tests/test_face_matcher_identity.py, tests/test_analyzer_identity.py.
**Cập nhật test:** nhóm tests/test_aggregate_report_*.py có sử dụng dữ liệu nhận diện.

### Giao diện

Thêm API nhận diện nghiêm ngặt; giữ API theo tên như adapter trả cần xác nhận nếu chưa có mapping, không âm thầm quay lại cách tìm cũ:

~~~python
@dataclass(frozen=True)
class MatchResult:
    status: str  # matched, no_match, identity_unresolved, no_portrait, error
    image_path: str | None
    distance: float | None
    project_id: str
    employee_id: str | None
    reason: str

def match_employee_in_images(
    self, *, project_id: str, employee_id: str, attendance_date: date,
    camera_images: list[str], distance_threshold: float | None = None,
    fast_mode: bool = True, log_detail: bool = False
) -> MatchResult: ...
~~~

### Các bước

- [x] Test cùng tên ở hai dự án; cùng tên khác mã trong một dự án; người chưa mapping; dự án không tồn tại; tên gần giống.
- [x] API phân tích resolve project_id trước khi tạo tác vụ; bỏ fallback PORTRAIT_DIR/INPUT_IMAGES_DIR toàn cục. Dự án không hợp lệ trả lỗi rõ ràng.
- [x] Truyền project_id và registry vào cả ExcelFaceAnalyzer, PDFFaceAnalyzer, ExcelToWordExporter.
- [x] Resolve payroll_code và membership theo ngày record; không dựa vào tên báo cáo tùy chỉnh để chọn dự án dữ liệu.
- [x] FaceMatcher lấy đúng tập portrait_bindings đã xác nhận, không duyệt các chân dung toàn cục.
- [x] Fuzzy/substring chỉ còn ở chức năng gợi ý trong giao diện; không sử dụng trong quyết định nhận diện hoặc tự sửa person['name'].
- [x] Thay _resolve_person_name của PDF bằng tên hiển thị từ danh tính đã resolve; giữ raw_name để đối chiếu.
- [x] Cả Excel và PDF gộp hồ sơ theo project_id + employee_id + kỳ. Hồ sơ chưa resolve giữ riêng theo nguồn; không gộp tất cả mã rỗng.
- [x] Sửa cả bước _dedupe_persons_by_name trong ExcelChamCongExtractor trước khi tách file: giữ riêng các mã khác nhau dù trùng tên; khi chưa có mã giữ nguồn riêng để không mất hồ sơ trước khi đến analyzer. Kiểm tra tên file trung gian cũng không ghi đè giữa hai người trùng tên.
- [x] Nếu cùng ID/kỳ có nhiều nguồn: hợp nhất record không xung đột theo ngày, bản ghi giống hệt được loại trùng; bản ghi cùng ngày khác giờ/giá trị phải được đánh dấu conflict thay vì chọn nguồn nhiều dòng hơn.
- [x] Tên file Word có employee_id và kỳ để hai người trùng tên không ghi đè.
- [x] Không chỉ thay ảnh camera: ảnh chân dung tiêu đề báo cáo cũng lấy từ binding định danh, bỏ tra cứu tên ở _find_portrait.
- [x] Record lưu identity_status, date_status, face_match_status và reason. Chỉ điền matched_image_path cho ảnh đủ điều kiện; ảnh ứng viên chưa xác nhận lưu ở candidate_image_path.
- [x] Báo cáo cá nhân/tổng hợp hiển thị nguyên nhân cần kiểm tra; không tự đổi unresolved/error thành vắng mặt hoặc không khớp.
- [x] Test mọi điểm gọi cũ: tìm bằng rg các hàm find_portrait/find_portraits/match_face_in_images/match_all_faces và constructor exporter/analyzer; adapter không được giữ đường vòng bỏ ràng buộc.

**Đạt khi:** không có ghép chéo dự án/tự chọn người theo tên; không mất hồ sơ trùng tên; kết quả Excel và PDF nhất quán.

## Hạng mục 6 — Kiểm thử tích hợp, chuyển đổi và bàn giao

**Tạo:** tests/test_attendance_matching_integration.py.
**Tạo khi triển khai:** tools/preview_attendance_migration.py và docs/attendance-data-migration.md.
**Cập nhật:** README.md với định dạng ngày, yêu cầu mã định danh và cách xử lý hồ sơ cần kiểm tra.

### Bộ dữ liệu thử bắt buộc

- Hai dự án A/B; mỗi dự án có người trùng tên.
- Hai người khác mã nhưng cùng tên trong dự án A.
- Ảnh ngày 01 của hai tháng, giao năm, thư mục DD chưa gán kỳ.
- Nhân viên chuyển dự án với ngày hiệu lực xác nhận.
- Một ảnh có ngày quan sát khác hồ sơ, một ảnh chưa đọc được ngày.
- Một worker nạp cache lần đầu và đợt xử lý làm cache chạm ngưỡng ghi.
- Dùng embedding giả khác nhau để kiểm soát đúng/sai; không đánh giá độ chính xác model bằng bộ fixture này.

### Lệnh kiểm chứng dự kiến

Dùng interpreter đã cài đủ dependency; py -3.12 dưới đây là lệnh dự kiến theo launcher, chưa được xác nhận chạy được trong môi trường công cụ hiện tại.

~~~text
py -3.12 -m pytest -q tests/test_face_matcher_cache.py tests/test_attendance_dates.py
py -3.12 -m pytest -q tests/test_daily_photo_dates.py tests/test_supplement_attendance_dates.py
py -3.12 -m pytest -q tests/test_identity_registry.py tests/test_face_matcher_identity.py tests/test_analyzer_identity.py tests/test_attendance_matching_integration.py
py -3.12 -m pytest -q tests/test_aggregate_report_api.py tests/test_aggregate_report_exporter.py tests/test_aggregate_report_ui.py tests/test_attendance_under_3h.py
node --test zalo-service/tests/metadata.test.js zalo-service/tests/download.test.js zalo-service/tests/progress.test.js tests/test_daily_photo_dates.cjs tests/test_launcher_config.cjs tests/test_supplement_picker.cjs
~~~

Không chạy mù toàn bộ tests/: một số script hiện có dùng dữ liệu thật hoặc gọi mạng. Rà soát tác dụng phụ trước khi mở rộng phạm vi chạy.

### Các bước bàn giao

- [x] Ghi nhận baseline trước sửa. Đã quan sát ở lần rà soát trước: nhóm 9 test JS đạt 8, lỗi còn lại liên quan target_time tự sinh; đây không phải bằng chứng lỗi do ba thay đổi mới.
- [x] Test tích hợp đi từ dữ liệu hồ sơ đến ảnh xuất báo cáo, kiểm tra đúng employee_id, project_id, ngày và lý do chưa xác nhận.
- [ ] Kiểm tra giao diện xem/upload/xóa đúng tháng; việc chưa mapping phải có thao tác giải quyết được, không chỉ thông báo lỗi.
- [x] Công cụ chuyển đổi mặc định chỉ xem trước: liệt kê thư mục DD, danh tính chưa mapping, xung đột và đích ISO dự kiến; không gọi AI hoặc sửa ảnh.
- [x] Khi áp dụng mapping đã xác nhận: sao lưu DB/manifest, copy ảnh sang đích mới, kiểm tra SHA-256, giữ nguyên nguồn; manifest lưu nguồn-đích/hash để chạy lại không tạo bản sao trùng hoặc ghi đè.
- [x] Chưa xóa nguồn trong lần chuyển đổi này. Không chọn toàn bộ DD là một tháng nếu có dữ liệu lẫn kỳ.
- [ ] Kiểm tra đóng/mở lại ứng dụng và bản đóng gói: registry/cache/manifest nằm đúng nơi, không phụ thuộc thư mục làm việc hiện tại.
- [x] Nếu cần rollback cache: cache có thể tái tạo. Nếu rollback dữ liệu: dùng snapshot và manifest; không chạy mã cũ trên dữ liệu mới đã chuẩn hóa nếu chưa kiểm tra tương thích.
- [ ] Rà soát diff và commit theo từng hạng mục, chỉ sau khi các kiểm thử tương ứng đạt.

## Điều kiện nghiệm thu toàn bộ

1. Cold start và lần ghi cache thứ 20 không treo; test có timeout bảo vệ runner.
2. Không có đường ghi mới tạo thư mục chỉ DD.
3. Không chọn/xóa/ghép ảnh chỉ vì trùng ngày trong tháng hoặc vì tìm thấy ở dự án khác.
4. Ngày quan sát bị lệch hoặc chưa xác định được thể hiện riêng, không biến ngày đề nghị thành ngày chụp.
5. Tên gần giống không thể tự quyết định danh tính; dữ liệu cũ chưa mapping vẫn được giữ và có hướng xử lý.
6. Hai người trùng tên không bị gộp hồ sơ hoặc ghi đè file báo cáo.
7. Luồng Excel/PDF, ảnh tiêu đề, ảnh camera và báo cáo tổng hợp cùng dùng định danh đã resolve.
8. Không sửa nội dung ảnh gốc khi chuyển đổi; có manifest, kiểm tra hash và khả năng khôi phục.

## Ghi nhận triển khai và bàn giao

- Nhánh: fix/attendance-reliability-20260916. Đã giữ bản sao mã nguồn trước sửa tại .cache/attendance-reliability-baseline; không sao chép hoặc chuyển đổi dữ liệu ảnh thật.
- Đã sửa khóa cache; bộ phân giải ngày dùng đủ ngày/tháng/năm; luồng ghi và đọc ảnh dùng ISO; registry định danh bền vững và membership theo thời gian; Excel/PDF, báo cáo và giao diện dùng ID đã xác nhận.
- Chân dung cũ được đăng ký pending và người quản lý nhập mã/ngày hiệu lực để xác nhận. Chưa làm màn hình gợi ý mapping tự động từ Excel/PDF; đây là phần giao diện phụ chưa triển khai, không ảnh hưởng quy tắc cấm tự ghép theo tên.
- Giao diện bổ sung hiển thị lý do từ chối áp dụng ảnh gốc; chưa có bảng riêng trình bày đồng thời ngày đề nghị/ngày quan sát cho mọi nguồn ảnh.
- Kiểm thử hồi quy Python: 82 bài đạt. Có kiểm thử timeout nạp/ghi cache, API ảnh theo tháng, trùng tên, lịch sử chuyển dự án, tích hợp ảnh/ngày/ID và công cụ chuyển đổi.
- Kiểm thử JS bao phủ ngày và ID, hai downloader và các bài hồi quy metadata/download/progress/launcher. Lỗi có sẵn ở test_supplement_picker.cjs liên quan target_time tự sinh vẫn được giữ và không sửa trong phạm vi này.
- Đã kiểm tra API và logic giao diện bằng test; chưa kiểm tra trực quan trên trình duyệt, đóng/mở bản EXE hoặc build bản phát hành. Không chạy Zalo thật, dịch vụ AI hay kiểm tra độ chính xác model.
- Công cụ chuyển đổi mặc định chỉ xem trước, có sao lưu/copy/hash/manifest; chỉ chạy trên fixture tạm trong lần triển khai này. Xem docs/attendance-data-migration.md.
- Không tạo commit/merge vì workspace chứa thay đổi có sẵn của người dùng xen kẽ trong các file. Các mục checklist gộp cả commit hoặc kiểm tra EXE vẫn để mở.

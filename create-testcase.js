const ExcelJS = require('exceljs');
const path = require('path');

async function createTestcase() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Testcase - Đặt lệnh Cơ sở');

    // Header styling
    const headerStyle = {
        font: { bold: true, size: 11 },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } },
        font: { bold: true, color: { argb: 'FFFFFFFF' } }
    };

    // Set columns
    worksheet.columns = [
        { header: 'STT', key: 'stt', width: 8 },
        { header: 'Module', key: 'module', width: 20 },
        { header: 'Test Scenario', key: 'scenario', width: 35 },
        { header: 'Test Case', key: 'testcase', width: 40 },
        { header: 'Precondition', key: 'precondition', width: 30 },
        { header: 'Input', key: 'input', width: 35 },
        { header: 'Expected Result', key: 'expected', width: 40 },
        { header: 'Priority', key: 'priority', width: 10 },
        { header: 'Status', key: 'status', width: 12 }
    ];

    // Apply header style
    worksheet.getRow(1).eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });

    // Test cases data - Section 1: Mô tả Lệnh thường
    const testcases = [
        // 1. Giao diện tab Giao dịch
        {
            stt: 1,
            module: 'Giao diện tab Giao dịch',
            scenario: 'Kiểm tra hiển thị Header',
            testcase: 'Kiểm tra Header hiển thị đúng theo design',
            precondition: 'Đăng nhập thành công vào Mobile App',
            input: 'Vào màn hình Giao dịch',
            expected: 'Header hiển thị đúng giao diện theo design mới',
            priority: 'High',
            status: ''
        },
        {
            stt: 2,
            module: 'Giao diện tab Giao dịch',
            scenario: 'Kiểm tra điều hướng tab',
            testcase: 'Kiểm tra chức năng chuyển qua lại giữa 2 tab',
            precondition: 'Đăng nhập thành công vào Mobile App',
            input: 'Chọn tab Giao dịch, thực hiện chuyển qua lại giữa các tab',
            expected: 'Cho phép chuyển qua lại giữa 2 tab, chức năng điều hướng không thay đổi',
            priority: 'High',
            status: ''
        },
        // 2. Tìm mã CK
        {
            stt: 3,
            module: 'Tìm mã CK',
            scenario: 'Hiển thị popup tìm kiếm khi chưa nhập mã',
            testcase: 'Kiểm tra hiển thị popup tìm kiếm mã khi chưa nhập mã CK',
            precondition: 'Đăng nhập thành công, vào màn hình Giao dịch',
            input: 'Chưa nhập mã CK, chọn vào ô tìm kiếm',
            expected: 'Hiển thị popup tìm kiếm mã, hiển thị danh sách all mã CK cơ sở',
            priority: 'High',
            status: ''
        },
        {
            stt: 4,
            module: 'Tìm mã CK',
            scenario: 'Hiển thị lịch sử tìm kiếm',
            testcase: 'Kiểm tra hiển thị lịch sử tìm kiếm khi đã có lịch sử',
            precondition: 'Đã tìm kiếm mã CK trước đó (lưu local ở app)',
            input: 'Chọn vào ô tìm kiếm mã CK',
            expected: 'Hiển thị danh sách Tìm kiếm gần đây, hiển thị tối đa 5 mã (mã tìm kiếm gần nhất hiển thị lên trên)',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 5,
            module: 'Tìm mã CK',
            scenario: 'Tìm kiếm mã CK theo từ khóa',
            testcase: 'Kiểm tra hiển thị đúng mã CK trùng khớp với text nhập',
            precondition: 'Đăng nhập thành công, vào popup tìm kiếm mã',
            input: 'Nhập từ khóa tìm kiếm mã CK (VD: VNM)',
            expected: 'Hiển thị đúng mã CK trùng khớp với text nhập',
            priority: 'High',
            status: ''
        },
        {
            stt: 6,
            module: 'Tìm mã CK',
            scenario: 'Chọn mã CK từ popup',
            testcase: 'Kiểm tra đóng popup và fill thông tin mã khi chọn mã',
            precondition: 'Đang ở popup tìm kiếm mã CK',
            input: 'Chọn một mã CK từ danh sách tìm kiếm',
            expected: 'Đóng popup, fill thông tin mã tại màn đặt lệnh (Tên mã CK, Sàn, Phiên, Giá thị trường, Thông tin mã, Deal khớp)',
            priority: 'High',
            status: ''
        },
        {
            stt: 7,
            module: 'Tìm mã CK',
            scenario: 'Hiển thị thông tin mã CK sau khi chọn',
            testcase: 'Kiểm tra hiển thị đầy đủ thông tin mã CK',
            precondition: 'Đã chọn mã CK từ popup tìm kiếm',
            input: 'Xem thông tin mã CK hiển thị',
            expected: 'Hiển thị đúng: Tên mã CK, Sàn (HNX), Phiên (Liên tục), Giá thị trường (+0.1)/(+0.47%), Thông tin mã (KL giao dịch, các giá Trần/TC/Sàn, icon đồ thị kỹ thuật), Deal khớp (3 deal khớp gần nhất và thanh tỷ lệ)',
            priority: 'High',
            status: ''
        },
        {
            stt: 8,
            module: 'Tìm mã CK',
            scenario: 'Chọn vào giá Thị trường/Trần/TC/Sàn',
            testcase: 'Kiểm tra fill giá khi chọn vào các giá',
            precondition: 'Đã chọn mã CK, hiển thị thông tin mã',
            input: 'Chọn vào giá Thị trường, Trần, TC, Sàn',
            expected: 'Fill giá tương ứng vào ô giá',
            priority: 'High',
            status: ''
        },
        {
            stt: 9,
            module: 'Tìm mã CK',
            scenario: 'Chọn icon đồ thị kỹ thuật',
            testcase: 'Kiểm tra điều hướng khi chọn icon đồ thị kỹ thuật',
            precondition: 'Đã chọn mã CK, hiển thị icon đồ thị kỹ thuật',
            input: 'Chọn vào icon đồ thị kỹ thuật',
            expected: 'Điều hướng sang màn hình Chi tiết mã',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 10,
            module: 'Tìm mã CK',
            scenario: 'Chọn Deal khớp',
            testcase: 'Kiểm tra fill giá khi chọn Deal khớp',
            precondition: 'Đã chọn mã CK, hiển thị Deal khớp',
            input: 'Chọn vào Deal khớp gần nhất',
            expected: 'Fill giá Deal khớp vào ô giá',
            priority: 'Medium',
            status: ''
        },
        // 3. Form lệnh
        {
            stt: 11,
            module: 'Form lệnh',
            scenario: 'Hiển thị form lệnh mặc định',
            testcase: 'Kiểm tra hiển thị form lệnh khi chưa nhập thông tin',
            precondition: 'Đăng nhập thành công, vào màn hình Giao dịch',
            input: 'Xem form lệnh khi chưa nhập thông tin',
            expected: 'Giao diện hiển thị đúng như design. TK: fill theo TK giao dịch mặc định. Trạng thái lệnh = TT mặc định. Loại lệnh = Lệnh cơ bản. Giá: trống, hiển thị giá gợi ý theo phiên (PLO, ATC,...). KL: trống',
            priority: 'High',
            status: ''
        },
        {
            stt: 12,
            module: 'Form lệnh',
            scenario: 'Chọn tài khoản giao dịch',
            testcase: 'Kiểm tra hiển thị popup chọn TK khi chọn vào ô TK',
            precondition: 'Đang ở form lệnh',
            input: 'Chọn vào ô Tài khoản',
            expected: 'Hiển thị popup chọn TK bao gồm: CC05091 (TK thường), CC05097 (TK MBLink), CC05098 (TK Margin)',
            priority: 'High',
            status: ''
        },
        {
            stt: 13,
            module: 'Form lệnh',
            scenario: 'Chọn Trạng thái lệnh',
            testcase: 'Kiểm tra hiển thị popup trạng thái lệnh',
            precondition: 'Đang ở form lệnh',
            input: 'Chọn vào ô Trạng thái lệnh',
            expected: 'Hiển thị popup các Trạng thái lệnh: Mặc định, Thường, OutRoom, TT68',
            priority: 'High',
            status: ''
        },
        {
            stt: 14,
            module: 'Form lệnh',
            scenario: 'Chọn Loại lệnh',
            testcase: 'Kiểm tra hiển thị popup loại lệnh',
            precondition: 'Đang ở form lệnh',
            input: 'Chọn vào ô Loại lệnh',
            expected: 'Hiển thị popup Chọn loại lệnh bao gồm: Lệnh cơ bản, Các lệnh điều kiện',
            priority: 'High',
            status: ''
        },
        {
            stt: 15,
            module: 'Form lệnh',
            scenario: 'Nhập giá từ bàn phím',
            testcase: 'Kiểm tra cho phép nhập giá từ bàn phím',
            precondition: 'Đang ở form lệnh, đã chọn mã CK',
            input: 'Nhập giá trực tiếp từ bàn phím vào ô giá',
            expected: 'Cho phép nhập giá từ bàn phím, hiển thị đúng giá đã nhập',
            priority: 'High',
            status: ''
        },
        {
            stt: 16,
            module: 'Form lệnh',
            scenario: 'Tăng/giảm giá bằng nút +/-',
            testcase: 'Kiểm tra cho phép chọn +- để tăng giảm giá',
            precondition: 'Đang ở form lệnh, đã nhập giá',
            input: 'Chọn nút + hoặc - để tăng/giảm giá',
            expected: 'Giá tăng/giảm đúng theo bước giá quy định',
            priority: 'High',
            status: ''
        },
        {
            stt: 17,
            module: 'Form lệnh',
            scenario: 'Nhập khối lượng',
            testcase: 'Kiểm tra cho phép nhập KL từ bàn phím',
            precondition: 'Đang ở form lệnh, đã chọn mã CK',
            input: 'Nhập khối lượng từ bàn phím vào ô KL',
            expected: 'Cho phép nhập KL, hiển thị đúng KL đã nhập',
            priority: 'High',
            status: ''
        },
        // 4. Sức mua
        {
            stt: 18,
            module: 'Sức mua',
            scenario: 'Hiển thị sức mua mặc định',
            testcase: 'Kiểm tra hiển thị đúng sức mua theo tài khoản mặc định',
            precondition: 'Đăng nhập thành công, vào màn hình Giao dịch',
            input: 'Xem thông tin Sức mua',
            expected: 'Hiển thị đúng: Sức mua, Mua tối đa, Bán tối đa, Tỷ lệ cho vay, Điểm HT, Gộp sức mua. Dữ liệu hiển thị theo đúng tài khoản đã chọn',
            priority: 'High',
            status: ''
        },
        {
            stt: 19,
            module: 'Sức mua',
            scenario: 'Thay đổi tài khoản xem sức mua',
            testcase: 'Kiểm tra cập nhật sức mua khi thay đổi tài khoản',
            precondition: 'Đang ở màn hình Giao dịch',
            input: 'Thay đổi tài khoản giao dịch',
            expected: 'Sức mua, Mua tối đa, Bán tối đa, Tỷ lệ cho vay, Điểm HT, Gộp sức mua cập nhật theo tài khoản mới',
            priority: 'High',
            status: ''
        },
        {
            stt: 20,
            module: 'Sức mua',
            scenario: 'Button + hiển thị popup',
            testcase: 'Kiểm tra hiển thị popup Phương thức tăng sức mua khi chọn button +',
            precondition: 'Đang ở màn hình Giao dịch',
            input: 'Chọn button +',
            expected: 'Hiển thị popup Phương thức tăng sức mua',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 21,
            module: 'Sức mua',
            scenario: 'Chọn Nộp tiền từ popup',
            testcase: 'Kiểm tra điều hướng khi chọn Nộp tiền',
            precondition: 'Đang hiển thị popup Phương thức tăng sức mua',
            input: 'Chọn Nộp tiền',
            expected: 'Điều hướng sang luồng Nộp tiền',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 22,
            module: 'Sức mua',
            scenario: 'Button Gộp sức mua hiển thị',
            testcase: 'Kiểm tra hiển thị button Gộp sức mua với TK đủ điều kiện',
            precondition: 'Đang ở màn hình Giao dịch, tài khoản đủ điều kiện',
            input: 'Xem button Gộp sức mua',
            expected: 'Hiển thị button Gộp sức mua với tài khoản đủ điều kiện',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 23,
            module: 'Sức mua',
            scenario: 'Chọn button Gộp sức mua',
            testcase: 'Kiểm tra điều hướng khi chọn button Gộp sức mua',
            precondition: 'Đang ở màn hình Giao dịch, button Gộp sức mua hiển thị',
            input: 'Chọn vào button Gộp sức mua',
            expected: 'Điều hướng sang màn hình Gộp sức mua',
            priority: 'Medium',
            status: ''
        },
        // 5. Sổ lệnh gần nhất
        {
            stt: 24,
            module: 'Sổ lệnh gần nhất',
            scenario: 'Hiển thị sổ lệnh có lệnh',
            testcase: 'Kiểm tra hiển thị 5 lệnh gần nhất',
            precondition: 'Đã đặt lệnh trước đó',
            input: 'Xem Sổ lệnh gần nhất',
            expected: 'Hiển thị 5 lệnh gần nhất với thông tin: Mã/TK, Loại lệnh, Loại GD, Khớp/Tổng, Giá khớp/Giá, TG đặt, TG khớp, TG hủy, SHL/Kênh',
            priority: 'High',
            status: ''
        },
        {
            stt: 25,
            module: 'Sổ lệnh gần nhất',
            scenario: 'Hiển thị sổ lệnh trống',
            testcase: 'Kiểm tra hiển thị khi chưa có lệnh nào',
            precondition: 'Chưa đặt lệnh nào',
            input: 'Xem Sổ lệnh gần nhất',
            expected: 'Hiển thị "Hiện tại chưa có lệnh nào"',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 26,
            module: 'Sổ lệnh gần nhất',
            scenario: 'Chọn vào Sổ lệnh',
            testcase: 'Kiểm tra điều hướng khi chọn vào Sổ lệnh',
            precondition: 'Đang ở màn hình Giao dịch',
            input: 'Chọn vào Sổ lệnh',
            expected: 'Điều hướng sang màn hình Sổ lệnh',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 27,
            module: 'Sổ lệnh gần nhất',
            scenario: 'Chọn vào từng lệnh',
            testcase: 'Kiểm tra điều hướng khi chọn vào lệnh cụ thể',
            precondition: 'Sổ lệnh gần nhất có hiển thị lệnh',
            input: 'Chọn vào một lệnh bất kỳ trong Sổ lệnh gần nhất',
            expected: 'Điều hướng tới màn hình Chi tiết lệnh',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 28,
            module: 'Sổ lệnh gần nhất',
            scenario: 'Scroll ngang xem thông tin lệnh',
            testcase: 'Kiểm tra scroll ngang để view thông tin lệnh',
            precondition: 'Sổ lệnh gần nhất có hiển thị lệnh',
            input: 'Scroll ngang trên Sổ lệnh gần nhất',
            expected: 'Cho phép scroll ngang để view đầy đủ thông tin lệnh',
            priority: 'Low',
            status: ''
        },
        // 6. Đặt Mua/Đặt Bán
        {
            stt: 29,
            module: 'Đặt Mua/Đặt Bán',
            scenario: 'Hiển thị button đặt lệnh',
            testcase: 'Kiểm tra hiển thị tách riêng 2 button ĐẶT MUA/ ĐẶT BÁN',
            precondition: 'Đăng nhập thành công, vào màn hình Giao dịch',
            input: 'Xem button đặt lệnh',
            expected: 'Tách riêng 2 button: ĐẶT MUA/ ĐẶT BÁN. Giá trị tạm tính giữ nguyên như hiện tại',
            priority: 'High',
            status: ''
        },
        {
            stt: 30,
            module: 'Đặt Mua/Đặt Bán',
            scenario: 'Hiển thị giá trị tạm tính',
            testcase: 'Kiểm tra hiển thị đúng Giá trị tạm tính = Giá * Khối lượng',
            precondition: 'Đã nhập Giá và Khối lượng',
            input: 'Xem Giá trị tạm tính',
            expected: 'Hiển thị đúng Giá trị tạm tính = Giá * Khối lượng (đã nhập ở trên)',
            priority: 'High',
            status: ''
        },
        {
            stt: 31,
            module: 'Đặt Mua/Đặt Bán',
            scenario: 'Đặt lệnh Mua thành công',
            testcase: 'Kiểm tra hiển thị popup Xác nhận lệnh khi chọn button MUA',
            precondition: 'Đã nhập đầy đủ thông tin hợp lệ, TK đủ số dư tiền',
            input: 'Chọn button ĐẶT MUA',
            expected: 'Hiển thị popup Xác nhận lệnh MUA',
            priority: 'High',
            status: ''
        },
        {
            stt: 32,
            module: 'Đặt Mua/Đặt Bán',
            scenario: 'Đặt lệnh Bán thành công',
            testcase: 'Kiểm tra hiển thị popup Xác nhận lệnh khi chọn button BÁN',
            precondition: 'Đã nhập đầy đủ thông tin hợp lệ, TK đủ số dư CK',
            input: 'Chọn button ĐẶT BÁN',
            expected: 'Hiển thị popup Xác nhận lệnh BÁN',
            priority: 'High',
            status: ''
        },
        // 7. Popup Xác nhận lệnh
        {
            stt: 33,
            module: 'Popup Xác nhận lệnh',
            scenario: 'Hiển thị popup xác nhận lệnh Mua',
            testcase: 'Kiểm tra hiển thị đúng thông tin trên popup Xác nhận lệnh Mua',
            precondition: 'Đã chọn button ĐẶT MUA',
            input: 'Xem popup Xác nhận lệnh Mua',
            expected: 'Hiển thị đúng các thông tin: Tài khoản, Loại lệnh, Mã CK, Giá mua, Khối lượng mua, Giá trị đặt lệnh',
            priority: 'High',
            status: ''
        },
        {
            stt: 34,
            module: 'Popup Xác nhận lệnh',
            scenario: 'Hiển thị popup xác nhận lệnh Bán',
            testcase: 'Kiểm tra hiển thị đúng thông tin trên popup Xác nhận lệnh Bán',
            precondition: 'Đã chọn button ĐẶT BÁN',
            input: 'Xem popup Xác nhận lệnh Bán',
            expected: 'Hiển thị đúng các thông tin: Tài khoản, Loại lệnh, Mã CK, Giá bán, Khối lượng bán, Giá trị đặt lệnh',
            priority: 'High',
            status: ''
        },
        {
            stt: 35,
            module: 'Popup Xác nhận lệnh',
            scenario: 'Xác nhận lệnh Mua thành công',
            testcase: 'Kiểm tra thông báo khi xác nhận lệnh Mua thành công',
            precondition: 'Đang ở popup Xác nhận lệnh Mua, TK đủ số dư tiền',
            input: 'Chọn button Xác nhận Mua',
            expected: 'Hiển thị thông báo "Đặt lệnh mua thành công"',
            priority: 'High',
            status: ''
        },
        {
            stt: 36,
            module: 'Popup Xác nhận lệnh',
            scenario: 'Xác nhận lệnh Bán thành công',
            testcase: 'Kiểm tra thông báo khi xác nhận lệnh Bán thành công',
            precondition: 'Đang ở popup Xác nhận lệnh Bán, TK đủ số dư CK',
            input: 'Chọn button Xác nhận Bán',
            expected: 'Hiển thị thông báo "Đặt lệnh bán thành công"',
            priority: 'High',
            status: ''
        },
        // 1.2 Nguyên tắc đặt lệnh - Lệnh MUA
        {
            stt: 37,
            module: 'Nguyên tắc đặt lệnh - MUA',
            scenario: 'Mua với TT Mặc định/Thường',
            testcase: 'Kiểm tra đặt lệnh Mua thành công với TT Mặc định/Thường',
            precondition: 'Đăng nhập với TK1, TK7, TK8',
            input: 'Đặt lệnh Mua với Trạng thái = Mặc định hoặc Thường',
            expected: 'Mua thành công với cả 3 tiểu khoản 1, 7, 8. Ghi nhận lệnh Mua thường',
            priority: 'High',
            status: ''
        },
        {
            stt: 38,
            module: 'Nguyên tắc đặt lệnh - MUA',
            scenario: 'Mua với TT OutRoom trên TK1',
            testcase: 'Kiểm tra lỗi khi đặt lệnh Mua OutRoom trên TK1',
            precondition: 'Đăng nhập với TK1',
            input: 'Đặt lệnh Mua với Trạng thái = OutRoom',
            expected: 'Thông báo lỗi: "Lệnh OutR không áp dụng cho TK1"',
            priority: 'High',
            status: ''
        },
        {
            stt: 39,
            module: 'Nguyên tắc đặt lệnh - MUA',
            scenario: 'Mua với TT OutRoom trên TK7, TK8',
            testcase: 'Kiểm tra đặt lệnh Mua OutRoom thành công trên TK7, TK8',
            precondition: 'Đăng nhập với TK7 hoặc TK8',
            input: 'Đặt lệnh Mua với Trạng thái = OutRoom',
            expected: 'Mua thành công',
            priority: 'High',
            status: ''
        },
        {
            stt: 40,
            module: 'Nguyên tắc đặt lệnh - MUA',
            scenario: 'Mua với TT TT68',
            testcase: 'Kiểm tra lỗi khi đặt lệnh Mua với TT68',
            precondition: 'Đăng nhập với bất kỳ TK',
            input: 'Đặt lệnh Mua với Trạng thái = TT68',
            expected: 'Đặt lệnh không thành công với tất cả các TK. Thông báo lỗi: "Lệnh TT68 chỉ áp dụng cho lệnh BÁN"',
            priority: 'High',
            status: ''
        },
        // 1.2 Nguyên tắc đặt lệnh - Lệnh BÁN
        {
            stt: 41,
            module: 'Nguyên tắc đặt lệnh - BÁN',
            scenario: 'Bán với TT Mặc định',
            testcase: 'Kiểm tra đặt lệnh Bán thành công với TT Mặc định',
            precondition: 'Đăng nhập với bất kỳ tiểu khoản',
            input: 'Đặt lệnh Bán với Trạng thái = Mặc định',
            expected: 'Đặt lệnh thành công với tất cả các tiểu khoản. Ghi nhận lệnh Bán gộp Trạng thái',
            priority: 'High',
            status: ''
        },
        {
            stt: 42,
            module: 'Nguyên tắc đặt lệnh - BÁN',
            scenario: 'Bán với TT Thường/OutRoom/TT68',
            testcase: 'Kiểm tra đặt lệnh Bán thành công với TT Thường/OutRoom/TT68',
            precondition: 'Đăng nhập với bất kỳ tiểu khoản',
            input: 'Đặt lệnh Bán với Trạng thái = Thường hoặc OutRoom hoặc TT68',
            expected: 'Đặt lệnh bán thành công. Ghi nhận lệnh bán theo trạng thái chọn tương ứng',
            priority: 'High',
            status: ''
        },
        // 1.3 Đặt lệnh từ các tính năng khác
        {
            stt: 43,
            module: 'Đặt lệnh từ DM đầu tư',
            scenario: 'Đặt lệnh từ Danh mục tổng',
            testcase: 'Kiểm tra fill thông tin khi đặt lệnh từ Danh mục tổng',
            precondition: 'Vào Danh mục đầu tư, chọn tab Danh mục tổng',
            input: 'Chọn button Đặt lệnh từ DM tổng',
            expected: 'Điều hướng tới màn hình Đặt lệnh giao diện mới. Fill Tài khoản: TK8 → TK7 → TK1. Trạng thái lệnh: fill theo logic đặt lệnh từ DM của các Tiểu khoản',
            priority: 'High',
            status: ''
        },
        {
            stt: 44,
            module: 'Đặt lệnh từ DM đầu tư',
            scenario: 'Đặt lệnh từ DM đuôi 1',
            testcase: 'Kiểm tra fill TT khi đặt lệnh từ DM đuôi 1',
            precondition: 'Vào Danh mục đầu tư, chọn DM đuôi 1',
            input: 'Chọn Đặt lệnh từ DM đuôi 1',
            expected: 'Fill TT = Mặc định với tất cả các trạng thái',
            priority: 'High',
            status: ''
        },
        {
            stt: 45,
            module: 'Đặt lệnh từ DM đầu tư',
            scenario: 'Đặt lệnh từ DM đuôi 7, 8',
            testcase: 'Kiểm tra fill TT khi đặt lệnh từ DM đuôi 7, 8',
            precondition: 'Vào Danh mục đầu tư, chọn DM đuôi 7 hoặc 8',
            input: 'Chọn đặt lệnh từ DM đuôi 7 hoặc 8',
            expected: 'TT = Thường/OutR/TT68 → Fill TT tương ứng. Các TT khác (07, 08,... - không được phép GD) → Fill TT = Mặc định',
            priority: 'High',
            status: ''
        },
        {
            stt: 46,
            module: 'Đặt lệnh từ Chi tiết mã',
            scenario: 'Đặt lệnh từ Chi tiết mã',
            testcase: 'Kiểm tra fill TT khi đặt lệnh từ Chi tiết mã',
            precondition: 'Vào Chi tiết mã CK',
            input: 'Chọn button Đặt lệnh từ Chi tiết mã',
            expected: 'Điều hướng màn hình Giao diện mới. Mặc định fill Trạng thái = Mặc định',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 47,
            module: 'Đặt lệnh từ DM Alpha',
            scenario: 'Đặt lệnh từ DM Alpha',
            testcase: 'Kiểm tra fill TT khi đặt lệnh từ DM Alpha',
            precondition: 'Vào Danh mục Alpha',
            input: 'Đặt lệnh từ Chi tiết khuyến nghị hoặc Icon đặt lệnh',
            expected: 'Điều hướng màn hình Giao diện mới. Mặc định fill Trạng thái = Mặc định',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 48,
            module: 'Đặt lệnh từ Sổ lệnh',
            scenario: 'Đặt lệnh từ Sổ lệnh gần nhất',
            testcase: 'Kiểm tra fill thông tin khi đặt lệnh từ Sổ lệnh',
            precondition: 'Vào Sổ lệnh/ Sổ lệnh gần nhất',
            input: 'Chọn Mua+/ Bán+ từ chi tiết lệnh',
            expected: 'Fill đúng thông tin lệnh trước đó',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 49,
            module: 'Đặt lệnh từ Pro-Advice',
            scenario: 'Đặt lệnh từ Pro-Advice',
            testcase: 'Kiểm tra fill TT khi đặt lệnh từ Pro-Advice',
            precondition: 'Vào Sản phẩm Pro-Advice',
            input: 'Đặt lệnh từ Chi tiết khuyến nghị/ Đặt lệnh hoặc Button Mua',
            expected: 'Điều hướng màn hình Giao diện mới. Mặc định fill Trạng thái = Mặc định',
            priority: 'Medium',
            status: ''
        },
        {
            stt: 50,
            module: 'Đặt lệnh từ Nộp tiền',
            scenario: 'Đặt lệnh từ popup Nộp tiền thành công',
            testcase: 'Kiểm tra fill TT khi đặt lệnh từ popup Nộp tiền thành công',
            precondition: 'Vừa nộp tiền thành công',
            input: 'Từ popup Nộp tiền thành công, chọn Đặt lệnh',
            expected: 'Điều hướng tới màn hình Giao diện mới. Mặc định fill Trạng thái = Mặc định',
            priority: 'Medium',
            status: ''
        }
    ];

    // Add data rows
    testcases.forEach((tc) => {
        const row = worksheet.addRow({
            stt: tc.stt,
            module: tc.module,
            scenario: tc.scenario,
            testcase: tc.testcase,
            precondition: tc.precondition,
            input: tc.input,
            expected: tc.expected,
            priority: tc.priority,
            status: tc.status
        });

        // Apply alignment to all cells in the row
        row.eachCell((cell) => {
            cell.alignment = { vertical: 'top', wrapText: true };
        });
    });

    // Set row height for header
    worksheet.getRow(1).height = 30;

    // Save file
    const outputPath = path.join(__dirname, 'Output', 'Testcase-Dat-lenh-co-so.xlsx');
    await workbook.xlsx.writeFile(outputPath);
    console.log('File created successfully at:', outputPath);
}

createTestcase().catch(console.error);

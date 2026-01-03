(async () => {

/* ========== 1. TẢI THƯ VIỆN XLSX ========== */
if (!window.XLSX) {
    console.log("⏳ Đang tải thư viện Excel...");
    await new Promise(r => {
        const s = document.createElement("script");
        s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
        s.onload = r;
        document.head.appendChild(s);
    });
}

/* ========== 2. HÀM ĐỊNH DẠNG NGÀY THÁNG (FIXED) ========== */
const getFormattedDate = (it) => {
    const info = it.info_card;
    const firstCard = info?.order_list_cards?.[0];
    
    // Tìm kiếm timestamp ở tất cả các vị trí có thể có trong JSON mới nhất của Shopee
    let timestamp = info?.create_time 
                 || firstCard?.product_info?.order_create_time 
                 || firstCard?.ctime 
                 || it?.create_time;

    if (!timestamp || timestamp <= 0) return "Không rõ ngày";

    // Shopee API trả về giây (10 chữ số), JS cần mili giây (13 chữ số)
    const date = new Date(timestamp * 1000);
    if (isNaN(date.getTime())) return "Lỗi định dạng";

    const d = date.getDate().toString().padStart(2, '0');
    const m = (date.getMonth() + 1).toString().padStart(2, '0');
    const y = date.getFullYear();
    const h = date.getHours().toString().padStart(2, '0');
    const min = date.getMinutes().toString().padStart(2, '0');

    return `${d}/${m}/${y} ${h}:${min}`;
};

/* ========== 3. CÀO DỮ LIỆU ĐƠN HÀNG ========== */
async function crawlOrders() {
    let offset = 0;
    const LIMIT = 20;
    const allOrders = [];
    const seenKeys = new Set();

    while (true) {
        console.log(`📥 Đang lấy dữ liệu đơn hàng (vị trí: ${offset})...`);
        try {
            const resp = await fetch(`https://shopee.vn/api/v4/order/get_all_order_and_checkout_list?limit=${LIMIT}&offset=${offset}`);
            const json = await resp.json();
            const list = json?.data?.order_data?.details_list ?? [];

            if (list.length === 0) break;
            offset += LIMIT;

            for (const it of list) {
                const info = it.info_card;
                if (!info) continue;

                const firstCard = info.order_list_cards?.[0];
                const shopName = firstCard?.shop_info?.shop_name ?? "Shopee";
                const dateStr = getFormattedDate(it);
                const finalAmount = (info.final_total ?? 0) / 1e5;
                const statusNum = it.list_type; // 3: Hoàn thành, 4: Đã hủy, 7,8: Vận chuyển/Đang giao
                
                const statusMap = {
                    3: "Hoàn thành", 4: "Đã hủy", 7: "Vận chuyển", 
                    8: "Đang giao", 9: "Chờ thanh toán", 12: "Trả hàng"
                };
                const statusText = statusMap[statusNum] ?? "Khác";

                // Tiền hàng gốc
                let rawSum = 0;
                const products = [];
                firstCard?.product_info?.item_groups?.forEach(g => {
                    g.items?.forEach(p => {
                        const pPrice = (p.order_price ?? 0) / 1e5;
                        rawSum += pPrice;
                        products.push({ name: p.name, qty: p.amount, price: pPrice });
                    });
                });

                const key = `${dateStr}-${shopName}-${finalAmount}`;
                if (seenKeys.has(key)) continue;
                seenKeys.add(key);

                allOrders.push({
                    date: dateStr,
                    shop: shopName,
                    total: finalAmount,
                    status: statusText,
                    isPaid: [3, 7, 8].includes(statusNum), // Đơn thực tế đã chi tiền
                    isCancelled: statusNum === 4,
                    items: products,
                    shipping: (info.shipping_fee ?? 0) / 1e5,
                    voucher: Math.max(0, (rawSum + (info.shipping_fee ?? 0) / 1e5) - finalAmount)
                });
            }
        } catch (e) {
            console.error("Lỗi fetch:", e);
            break;
        }
        await new Promise(r => setTimeout(r, 400));
    }
    return allOrders;
}

/* ========== 4. XUẤT EXCEL & TỔNG KẾT ========== */
function exportToExcel(orders) {
    const rows = [];
    let sumPaid = 0;
    let sumCancelled = 0;
    let countSuccess = 0;

    orders.forEach((o, i) => {
        if (o.isPaid) {
            sumPaid += o.total;
            countSuccess++;
        }
        if (o.isCancelled) sumCancelled += o.total;

        rows.push({
            "STT": i + 1,
            "Ngày đặt": o.date,
            "Cửa hàng": o.shop,
            "Nội dung": "ĐƠN HÀNG",
            "Số lượng": o.items.reduce((a, b) => a + b.qty, 0),
            "Thanh toán (VNĐ)": o.total,
            "Trạng thái": o.status
        });

        o.items.forEach(item => {
            rows.push({
                "Nội dung": "↳ " + item.name,
                "Số lượng": item.qty,
                "Thanh toán (VNĐ)": item.price
            });
        });
    });

    // Thêm dòng tổng kết vào cuối file Excel
    rows.push({});
    rows.push({ "Cửa hàng": "--- TỔNG KẾT CHI TIÊU ---" });
    rows.push({ "Cửa hàng": "Tổng số đơn hàng:", "Thanh toán (VNĐ)": orders.length });
    rows.push({ "Cửa hàng": "Đơn thành công:", "Thanh toán (VNĐ)": countSuccess });
    rows.push({ "Cửa hàng": "TỔNG TIỀN ĐÃ THANH TOÁN:", "Thanh toán (VNĐ)": sumPaid });
    rows.push({ "Cửa hàng": "Tổng tiền đơn đã hủy:", "Thanh toán (VNĐ)": sumCancelled });

    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 5 }, { wch: 20 }, { wch: 25 }, { wch: 50 }, { wch: 10 }, { wch: 15 }, { wch: 15 }];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Shopee_Orders");
    XLSX.writeFile(wb, `Shopee_Bao_Cao_Chi_Tieu.xlsx`);
}

/* ========== 5. GIAO DIỆN WEB ========== */
function displayOnWeb(orders) {
    const paid = orders.filter(o => o.isPaid).reduce((a, b) => a + b.total, 0);
    
    document.body.innerHTML = `
        <div style="font-family: Arial; padding: 25px; background: #f4f4f4; min-height: 100vh;">
            <div style="max-width: 900px; margin: auto; background: white; padding: 30px; border-radius: 15px; box-shadow: 0 5px 20px rgba(0,0,0,0.1);">
                <h1 style="color: #ee4d2d; text-align: center;">📊 BÁO CÁO CHI TIÊU SHOPEE</h1>
                
                <div style="display: flex; justify-content: space-around; background: #fff5f2; padding: 20px; border-radius: 10px; margin: 20px 0;">
                    <div style="text-align: center;">
                        <p style="margin: 0; color: #666;">Tổng đơn hàng</p>
                        <b style="font-size: 24px;">${orders.length}</b>
                    </div>
                    <div style="text-align: center;">
                        <p style="margin: 0; color: #666;">Đơn thành công</p>
                        <b style="font-size: 24px; color: #26aa99;">${orders.filter(o => o.isPaid).length}</b>
                    </div>
                    <div style="text-align: center;">
                        <p style="margin: 0; color: #666;">Tổng tiền đã chi</p>
                        <b style="font-size: 24px; color: #ee4d2d;">${paid.toLocaleString()} VNĐ</b>
                    </div>
                </div>

                <button id="btnDL" style="width: 100%; padding: 15px; background: #ee4d2d; color: white; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 16px;">⬇️ TẢI FILE EXCEL CHI TIẾT</button>
                
                <h3 style="margin-top: 30px; border-bottom: 2px solid #eee; padding-bottom: 10px;">Lịch sử đơn hàng:</h3>
                <div id="orderList">
                    ${orders.map(o => `
                        <div style="padding: 12px; border-bottom: 1px solid #eee; display: flex; justify-content: space-between; font-size: 14px;">
                            <span><b>${o.date}</b> - ${o.shop}</span>
                            <span style="color: ${o.isPaid ? '#26aa99' : '#ee4d2d'}">${o.total.toLocaleString()}đ [${o.status}]</span>
                        </div>
                    `).join('')}
                </div>
            </div>
        </div>
    `;
    document.getElementById("btnDL").onclick = () => exportToExcel(orders);
}

/* ========== RUN ========== */
(async () => {
    console.log("🚀 Bắt đầu quét dữ liệu... Vui lòng không đóng tab.");
    const results = await crawlOrders();
    if (results.length > 0) {
        displayOnWeb(results);
        console.log("✅ Hoàn thành!");
    } else {
        alert("Không lấy được dữ liệu. Hãy đảm bảo bạn đang ở trang Shopee và đã đăng nhập.");
    }
})();

})();

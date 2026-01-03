(async () => {

/* ========== 1. LOAD THƯ VIỆN XLSX ========== */
if (!window.XLSX) {
    console.log("⏳ Đang tải thư viện Excel...");
    await new Promise(r => {
        const s = document.createElement("script");
        s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
        s.onload = r;
        document.head.appendChild(s);
    });
}

/* ========== 2. HÀM ĐỊNH DẠNG NGÀY (TRƯỜNG DỮ LIỆU GỐC CHUẨN) ========== */
const parseShopeeDate = (it) => {
    // Trường dữ liệu chuẩn nhất của Shopee v4: it.info_card.create_time
    // Fallback nếu không có: it.info_card.order_list_cards[0].product_info.order_create_time
    let ts = it?.info_card?.create_time 
          || it?.info_card?.order_list_cards?.[0]?.product_info?.order_create_time 
          || it?.create_time;

    if (!ts) return "Không rõ ngày";

    // Shopee dùng giây (10 số), JS dùng mili giây (13 số). Ví dụ: 1735900000 -> 1735900000000
    const date = new Date(ts < 1e11 ? ts * 1000 : ts);
    
    if (isNaN(date.getTime())) return "Ngày lỗi";

    const p = (n) => n.toString().padStart(2, '0');
    return `${p(date.getDate())}/${p(date.getMonth() + 1)}/${date.getFullYear()} ${p(date.getHours())}:${p(date.getMinutes())}`;
};

/* ========== 3. QUY TRÌNH CÀO DỮ LIỆU (CRAWL) ========== */
async function crawlAll() {
    let offset = 0;
    const LIMIT = 20;
    const allOrders = [];
    const seenKeys = new Set();

    // Hiển thị trạng thái đang quét lên màn hình
    document.body.innerHTML = `<div id="status" style="font-family:Arial;padding:50px;text-align:center;">
        <h2 style="color:#ee4d2d">🚀 Đang quét toàn bộ đơn hàng...</h2>
        <p id="count">Đã tìm thấy: 0 đơn hàng</p>
        <p style="color:#666">Vui lòng không đóng trình duyệt.</p>
    </div>`;

    while (true) {
        try {
            const r = await fetch(`https://shopee.vn/api/v4/order/get_all_order_and_checkout_list?limit=${LIMIT}&offset=${offset}`);
            const j = await r.json();
            const list = j?.data?.order_data?.details_list ?? [];

            if (list.length === 0) break; // Hết đơn hàng

            for (const it of list) {
                const info = it.info_card;
                if (!info) continue;

                const dateStr = parseShopeeDate(it);
                const card = info.order_list_cards?.[0];
                const shop = card?.shop_info?.shop_name ?? "Shopee";
                const final = (info.final_total ?? 0) / 1e5;

                // Chống trùng lặp đơn hàng
                const key = `${dateStr}_${shop}_${final}`;
                if (seenKeys.has(key)) continue;
                seenKeys.add(key);

                const statusMap = {3:"Hoàn thành",4:"Đã hủy",7:"Vận chuyển",8:"Đang giao",9:"Chờ thanh toán",12:"Trả hàng"};
                
                let itemSum = 0;
                const items = [];
                card?.product_info?.item_groups?.forEach(g => {
                    g.items?.forEach(p => {
                        const price = (p.order_price ?? 0) / 1e5;
                        itemSum += price;
                        items.push({ name: p.name, qty: p.amount, total: price });
                    });
                });

                allOrders.push({
                    date: dateStr,
                    shop,
                    status: statusMap[it.list_type] || `Khác (${it.list_type})`,
                    total: final,
                    itemSum: itemSum,
                    ship: (info.shipping_fee ?? 0) / 1e5,
                    isSuccess: [3, 7, 8].includes(it.list_type),
                    items: items
                });
            }

            offset += LIMIT;
            document.getElementById("count").innerText = `Đã tìm thấy: ${allOrders.length} đơn hàng`;
            await new Promise(res => setTimeout(res, 400)); // Nghỉ để tránh block IP

        } catch (e) {
            console.error(e);
            break;
        }
    }
    return allOrders;
}

/* ========== 4. GIAO DIỆN WEB (NHƯ BẢN CŨ) ========== */
function renderWeb(orders) {
    const totalPaid = orders.filter(o => o.isSuccess).reduce((s, o) => s + o.total, 0);
    
    document.body.innerHTML = `
        <div style="font-family:Segoe UI,Arial; padding:20px; background:#f4f4f4; color:#333;">
            <div style="max-width:900px; margin:auto; background:#fff; padding:30px; border-radius:12px; box-shadow:0 4px 20px rgba(0,0,0,0.1);">
                <h2 style="color:#ee4d2d; margin-top:0;">📊 TỔNG KẾT CHI TIÊU SHOPEE</h2>
                
                <div style="display:flex; gap:20px; margin-bottom:25px;">
                    <div style="flex:1; background:#fff5f2; border:1px solid #ffdbd0; padding:20px; border-radius:8px;">
                        <span style="font-size:14px; color:#666;">Tổng tiền đã thanh toán</span><br>
                        <b style="font-size:24px; color:#ee4d2d;">${totalPaid.toLocaleString()}đ</b>
                    </div>
                    <div style="flex:1; background:#f6f6f6; border:1px solid #ddd; padding:20px; border-radius:8px;">
                        <span style="font-size:14px; color:#666;">Tổng số đơn hàng</span><br>
                        <b style="font-size:24px;">${orders.length} đơn</b>
                    </div>
                </div>

                <button id="dlBtn" style="width:100%; padding:15px; background:#ee4d2d; color:#fff; border:none; border-radius:6px; font-weight:bold; cursor:pointer; font-size:16px; margin-bottom:30px;">⬇️ TẢI BÁO CÁO EXCEL (.XLSX)</button>

                <h3 style="border-bottom:2px solid #eee; padding-bottom:10px;">Lịch sử chi tiết:</h3>
                <div id="listContainer"></div>
            </div>
        </div>
    `;

    const container = document.getElementById("listContainer");
    orders.forEach((o, i) => {
        const item = document.createElement("details");
        item.style.cssText = "margin-bottom:10px; border:1px solid #eee; border-radius:5px; padding:10px;";
        if (!o.isSuccess) item.style.background = "#fafafa";

        item.innerHTML = `
            <summary style="cursor:pointer; font-weight:bold; display:flex; justify-content:space-between;">
                <span>#${i+1}. [${o.date}] - ${o.shop}</span>
                <span style="color:${o.isSuccess ? '#26aa99' : '#999'}">${o.total.toLocaleString()}đ</span>
            </summary>
            <div style="font-size:13px; color:#666; padding-top:10px; border-top:1px solid #f9f9f9; margin-top:10px;">
                <p>Trạng thái: <b>${o.status}</b></p>
                <ul style="padding-left:15px;">
                    ${o.items.map(p => `<li>${p.name} (x${p.qty}) - ${p.total.toLocaleString()}đ</li>`).join('')}
                </ul>
            </div>
        `;
        container.appendChild(item);
    });

    document.getElementById("dlBtn").onclick = () => exportExcel(orders);
}

/* ========== 5. XUẤT EXCEL ========== */
function exportExcel(orders) {
    const data = [];
    orders.forEach((o, i) => {
        data.push({
            "STT": i + 1,
            "Ngày đặt": o.date,
            "Shop": o.shop,
            "Nội dung": "--- TỔNG ĐƠN ---",
            "Thực trả": o.total,
            "Trạng thái": o.status
        });
        o.items.forEach(it => {
            data.push({ "Nội dung": "↳ " + it.name, "Thực trả": it.total });
        });
    });

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Shopee");
    XLSX.writeFile(wb, `Shopee_Report_${new Date().getTime()}.xlsx`);
}

/* ========== RUN ========== */
const results = await crawlAll();
renderWeb(results);

})();

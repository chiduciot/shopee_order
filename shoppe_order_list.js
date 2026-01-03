(async () => {
/* ========== 1. TẢI THƯ VIỆN XLSX ========== */
if (!window.XLSX) {
    await new Promise(r => {
        const s = document.createElement("script");
        s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
        s.onload = r;
        document.head.appendChild(s);
    });
}

/* ========== 2. HÀM TÌM NGÀY THÁNG (CƠ CHẾ X-RAY) ========== */
function extractDate(it) {
    // Thử lấy từ các trường phổ biến nhất của Shopee 2026
    let timestamp = it?.info_card?.create_time 
                 || it?.create_time 
                 || it?.info_card?.order_list_cards?.[0]?.product_info?.order_create_time
                 || it?.info_card?.order_list_cards?.[0]?.ctime;

    if (!timestamp) return "Không rõ ngày";

    // Chuyển đổi timestamp (10 số sang 13 số nếu cần)
    let dateObj = new Date(timestamp < 1e12 ? timestamp * 1000 : timestamp);
    
    if (isNaN(dateObj.getTime())) return "Lỗi ngày";

    const d = dateObj.getDate().toString().padStart(2, '0');
    const m = (dateObj.getMonth() + 1).toString().padStart(2, '0');
    const y = dateObj.getFullYear();
    const h = dateObj.getHours().toString().padStart(2, '0');
    const min = dateObj.getMinutes().toString().padStart(2, '0');

    return `${d}/${m}/${y} ${h}:${min}`;
}

/* ========== 3. QUY TRÌNH QUÉT TOÀN BỘ ĐƠN HÀNG ========== */
async function crawl() {
    let offset = 0;
    const LIMIT = 20;
    const orders = [];
    const seen = new Set();

    // Hiển thị màn hình chờ hiện đại
    document.body.innerHTML = `
        <div id="loader" style="font-family:Arial; text-align:center; padding-top:100px; background:#fff; position:fixed; top:0; left:0; width:100%; height:100%; z-index:9999;">
            <div style="border: 8px solid #f3f3f3; border-top: 8px solid #ee4d2d; border-radius: 50%; width: 60px; height: 60px; animation: spin 1s linear infinite; margin:auto;"></div>
            <h2 style="color:#ee4d2d; margin-top:20px;">🚀 ĐANG QUÉT DỮ LIỆU ĐƠN HÀNG...</h2>
            <p id="progress" style="font-size:18px; color:#555;">Khởi động...</p>
            <style>@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }</style>
        </div>
    `;

    while (true) {
        try {
            const r = await fetch(`https://shopee.vn/api/v4/order/get_all_order_and_checkout_list?limit=${LIMIT}&offset=${offset}`);
            const j = await r.json();
            const list = j?.data?.order_data?.details_list ?? [];

            if (list.length === 0) break;

            for (const entry of list) {
                const info = entry.info_card;
                if (!info) continue;

                const dateStr = extractDate(entry);
                const cards = info.order_list_cards || [];

                cards.forEach(card => {
                    const shop = card.shop_info?.shop_name || "Shopee";
                    const total = (info.final_total ?? 0) / 1e5;
                    const statusText = {3:"Hoàn thành", 4:"Đã hủy", 7:"Vận chuyển", 8:"Đang giao", 12:"Trả hàng"}[entry.list_type] || "Khác";

                    const products = [];
                    card.product_info?.item_groups?.forEach(g => {
                        g.items?.forEach(p => {
                            products.push({
                                name: p.name,
                                qty: p.amount,
                                price: (p.order_price ?? 0) / 1e5
                            });
                        });
                    });

                    // Chống trùng (Unique Key)
                    const key = `${dateStr}_${shop}_${total}`;
                    if (!seen.has(key)) {
                        seen.add(key);
                        orders.push({ date: dateStr, shop, total, status: statusText, products, isSuccess: [3,7,8].includes(entry.list_type) });
                    }
                });
            }
            offset += LIMIT;
            document.getElementById("progress").innerText = `Đã tìm thấy ${orders.length} đơn hàng...`;
            await new Promise(res => setTimeout(res, 400));
        } catch (e) { break; }
    }
    return orders;
}

/* ========== 4. GIAO DIỆN WEB (DETAILS & TỔNG KẾT) ========== */
function render(orders) {
    const totalSpent = orders.filter(o => o.isSuccess).reduce((s, o) => s + o.total, 0);

    document.body.innerHTML = `
        <div style="font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding:40px; background:#f0f2f5; min-height:100vh;">
            <div style="max-width:1000px; margin:auto; background:#fff; padding:40px; border-radius:20px; box-shadow:0 10px 40px rgba(0,0,0,0.1);">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:30px;">
                    <h1 style="color:#ee4d2d; margin:0;">🛍️ THỐNG KÊ CHI TIÊU SHOPEE</h1>
                    <button id="excelBtn" style="background:#26aa99; color:#fff; border:none; padding:12px 25px; border-radius:10px; cursor:pointer; font-weight:bold; font-size:16px;">⬇️ XUẤT FILE EXCEL</button>
                </div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:20px; margin-bottom:40px;">
                    <div style="background:#fff5f2; border:1px solid #ffdbd0; padding:25px; border-radius:15px;">
                        <span style="color:#666; font-size:14px; text-transform:uppercase; letter-spacing:1px;">Tổng tiền đã thanh toán</span><br>
                        <b style="font-size:32px; color:#ee4d2d;">${totalSpent.toLocaleString()} VNĐ</b>
                    </div>
                    <div style="background:#f0f7f6; border:1px solid #d1e7e4; padding:25px; border-radius:15px;">
                        <span style="color:#666; font-size:14px; text-transform:uppercase; letter-spacing:1px;">Tổng đơn hàng đã mua</span><br>
                        <b style="font-size:32px; color:#26aa99;">${orders.length} đơn</b>
                    </div>
                </div>

                <div id="order-list">
                    ${orders.map((o, i) => `
                        <div style="border: 1px solid #eee; border-radius: 12px; margin-bottom: 15px; overflow: hidden; background:#fff;">
                            <div onclick="this.nextElementSibling.style.display = this.nextElementSibling.style.display === 'none' ? 'block' : 'none'" 
                                 style="padding: 18px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; transition: background 0.3s;"
                                 onmouseover="this.style.background='#fafafa'" onmouseout="this.style.background='#fff'">
                                <div>
                                    <span style="color:#888; font-size:12px;">#${i + 1}</span>
                                    <strong style="display:block; font-size:15px;">${o.date} | ${o.shop}</strong>
                                </div>
                                <div style="text-align:right;">
                                    <strong style="color:#ee4d2d; font-size:16px;">${o.total.toLocaleString()}đ</strong>
                                    <span style="display:block; font-size:11px; color:#999;">${o.status} ⌵</span>
                                </div>
                            </div>
                            <div style="display:none; padding: 20px; background: #fcfcfc; border-top: 1px solid #eee;">
                                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                                    <thead>
                                        <tr style="text-align: left; color: #777; border-bottom: 1px solid #eee;">
                                            <th style="padding: 8px 0;">Sản phẩm</th>
                                            <th style="padding: 8px 0; width: 60px; text-align:center;">SL</th>
                                            <th style="padding: 8px 0; width: 100px; text-align:right;">Giá</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${o.products.map(p => `
                                            <tr style="border-bottom: 1px solid #f1f1f1;">
                                                <td style="padding: 10px 0;">${p.name}</td>
                                                <td style="padding: 10px 0; text-align:center;">${p.qty}</td>
                                                <td style="padding: 10px 0; text-align:right;">${p.price.toLocaleString()}đ</td>
                                            </tr>
                                        `).join('')}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    `).join('')}
                </div>
            </div>
        </div>
    `;

    document.getElementById("excelBtn").onclick = () => {
        const rows = [];
        orders.forEach((o, idx) => {
            rows.push({ "STT": idx + 1, "Ngày": o.date, "Shop": o.shop, "Nội dung": "TỔNG ĐƠN", "Tiền": o.total, "Trạng thái": o.status });
            o.products.forEach(p => rows.push({ "Nội dung": "↳ " + p.name, "Tiền": p.price + " x " + p.qty }));
        });
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Shopee");
        XLSX.writeFile(wb, `Shopee_Report_2026.xlsx`);
    };
}

/* ========== KHỞI CHẠY ========== */
const results = await crawl();
render(results);

})();

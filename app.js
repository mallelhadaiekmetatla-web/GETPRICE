const EXCEL_PATH = "./data/products.xlsx";
let products = new Map();
let scanner = null;
let scanning = false;
const $ = (id) => document.getElementById(id);
const normalizeBarcode = (value) => String(value ?? "").trim().replace(/\.0$/, "");

async function loadProducts() {
  const status = $("database-status");
  try {
    const response = await fetch(EXCEL_PATH + "?v=" + Date.now());
    if (!response.ok) throw new Error("تعذر تحميل ملف Excel.");
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellText: true, cellDates: false });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    products.clear();
    for (const row of rows) {
      const barcode = normalizeBarcode(row["CODE BARRE"] ?? row["CODE_BARRE"] ?? row["barcode"] ?? row["Barcode"]);
      const designation = String(row["DESIGNATION"] ?? row["designation"] ?? row["Designation"] ?? "").trim();
      const priceRaw = row["PRIX DE VENTE"] ?? row["PRIX_DE_VENTE"] ?? row["price"] ?? row["Price"] ?? "";
      const price = Number(String(priceRaw).replace(/\s/g, "").replace(",", "."));
      if (barcode && designation && Number.isFinite(price)) products.set(barcode, { barcode, designation, price });
    }
    status.textContent = `قاعدة البيانات جاهزة • ${products.size} منتج`;
    status.classList.add("loaded");
  } catch (error) {
    status.textContent = "تعذر تحميل Excel. تأكد من وجود الملف داخل data/products.xlsx.";
    showNotice("لم يتم تحميل قاعدة المنتجات. راجع ملف Excel واتصال الإنترنت.");
    console.error(error);
  }
}

function formatPrice(price) {
  return new Intl.NumberFormat("fr-DZ", { maximumFractionDigits: 2 }).format(price);
}
function showNotice(message, good = false) {
  $("notice").textContent = message;
  $("notice").style.color = good ? "#16896d" : "#d34d69";
}
function searchProduct(value) {
  const barcode = normalizeBarcode(value);
  if (!barcode) return;
  $("barcode").value = barcode;
  const product = products.get(barcode);
  if (!product) {
    $("product-name").textContent = "المنتج غير موجود";
    $("product-price").textContent = "—";
    $("product-barcode").textContent = `الكود: ${barcode}`;
    showNotice("لم نعثر على هذا الكود في قاعدة البيانات. تحقق من الرقم أو حدّث ملف Excel.");
    return;
  }
  $("product-name").textContent = product.designation;
  $("product-price").textContent = formatPrice(product.price);
  $("product-barcode").textContent = `CODE BARRE : ${product.barcode}`;
  showNotice("تم العثور على المنتج بنجاح", true);
  $("result").scrollIntoView({ behavior: "smooth", block: "nearest" });
}
$("search-form").addEventListener("submit", (event) => {
  event.preventDefault();
  searchProduct($("barcode").value);
});
async function startScanner() {
  if (scanning) return;
  if (!window.Html5Qrcode) {
    showNotice("مكتبة الكاميرا لم تُحمّل. تحقق من اتصال الإنترنت.");
    return;
  }
  try {
    scanner = new Html5Qrcode("reader");
    const cameras = await Html5Qrcode.getCameras();
    if (!cameras || !cameras.length) throw new Error("لم يتم العثور على كاميرا.");
    const preferred = cameras.find(c => /back|rear|environment|arrière|خلف/i.test(c.label)) || cameras[cameras.length - 1];
    $("reader").style.minHeight = "220px";
    await scanner.start(
      preferred.id,
      { fps: 10, qrbox: (w, h) => ({ width: Math.min(300, Math.floor(w * 0.88)), height: Math.min(140, Math.floor(h * 0.55)) }), aspectRatio: 1.6 },
      (decodedText) => {
        searchProduct(decodedText);
        stopScanner();
      },
      () => {}
    );
    scanning = true;
    $("start-scan").disabled = true;
    $("stop-scan").disabled = false;
    showNotice("الكاميرا تعمل. وجّهها نحو الكود بارك.", true);
  } catch (error) {
    console.error(error);
    showNotice("تعذر فتح الكاميرا. اسمح للمتصفح باستخدامها وافتح الموقع عبر HTTPS.");
    if (scanner) { try { await scanner.clear(); } catch (_) {} scanner = null; }
  }
}
async function stopScanner() {
  if (scanner && scanning) {
    try { await scanner.stop(); } catch (_) {}
    try { await scanner.clear(); } catch (_) {}
  } else if (scanner) {
    try { await scanner.clear(); } catch (_) {}
  }
  scanner = null;
  scanning = false;
  $("start-scan").disabled = false;
  $("stop-scan").disabled = true;
  $("reader").style.minHeight = "0";
}
$("start-scan").addEventListener("click", startScanner);
$("stop-scan").addEventListener("click", stopScanner);
loadProducts();

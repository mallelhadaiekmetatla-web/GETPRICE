
const EXCEL_PATH = "./data/products.xlsx";

let products = new Map();
let scanner = null;
let scanning = false;

const $ = (id) => document.getElementById(id);

// الحفاظ على الباركود كنص، بما في ذلك الأصفار في البداية
function normalizeBarcode(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/\.0$/, "");
}

// توحيد أسماء الأعمدة لتجاوز اختلاف المسافات والشرطات والحروف
function normalizeHeader(value) {
  return String(value ?? "")
    .replace(/^\uFEFF/, "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ");
}

// البحث عن العمود حتى لو اختلفت طريقة كتابته
function getColumn(row, possibleNames) {
  const keys = Object.keys(row);
  const normalizedKeys = new Map(
    keys.map((key) => [normalizeHeader(key), key])
  );

  for (const name of possibleNames) {
    const actualKey = normalizedKeys.get(normalizeHeader(name));
    if (actualKey !== undefined) {
      return row[actualKey];
    }
  }

  return "";
}

// تحويل السعر إلى رقم مع دعم الفاصلة العشرية
function parsePrice(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : NaN;
  }

  let price = String(value ?? "")
    .trim()
    .replace(/\s/g, "")
    .replace(/[^\d.,\-]/g, "");

  if (!price) return NaN;

  const comma = price.lastIndexOf(",");
  const dot = price.lastIndexOf(".");

  // إذا وُجدت الفاصلة والنقطة معًا، نعتبر آخر واحدة فاصلة عشرية
  if (comma !== -1 && dot !== -1) {
    if (comma > dot) {
      price = price.replace(/\./g, "").replace(",", ".");
    } else {
      price = price.replace(/,/g, "");
    }
  } else if (comma !== -1) {
    price = price.replace(",", ".");
  }

  const result = Number(price);
  return Number.isFinite(result) ? result : NaN;
}

async function loadProducts() {
  const status = $("database-status");

  try {
    if (!status) {
      throw new Error('لم يتم العثور على العنصر "database-status" في HTML.');
    }

    status.textContent = "جاري تحميل قاعدة المنتجات...";

    const response = await fetch(
      `${EXCEL_PATH}?v=${Date.now()}`,
      { cache: "no-store" }
    );

    if (!response.ok) {
      throw new Error(
        `تعذر تحميل Excel. HTTP ${response.status}: ${response.statusText}`
      );
    }

    const buffer = await response.arrayBuffer();

    if (buffer.byteLength === 0) {
      throw new Error("ملف Excel فارغ.");
    }

    if (typeof XLSX === "undefined") {
      throw new Error("مكتبة قراءة Excel لم تُحمّل.");
    }

    const workbook = XLSX.read(buffer, {
      type: "array",
      cellText: true,
      cellDates: false
    });

    if (!workbook.SheetNames.length) {
      throw new Error("لا توجد أوراق داخل ملف Excel.");
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: false
    });

    products.clear();

    let invalidRows = 0;

    for (const row of rows) {
      const barcode = normalizeBarcode(
        getColumn(row, [
          "Code Barre",
          "CODE BARRES",
          "CODE_BARRE",
          "BARCODE",
          "BAR CODE",
          "EAN",
          "EAN13"
        ])
      );

      const designation = String(
        getColumn(row, [
          "DESIGNATION",
          "DESIGNATION PRODUIT",
          "Produit",
          "PRODUCT NAME"
        ]) ?? ""
      ).trim();

      const price = parsePrice(
        getColumn(row, [
          "Prix de Vente H.T",
          "PRIX_DE_VENTE",
          "PRIX VENTE",
          "PRICE",
          "SELLING PRICE"
        ])
      );

      // تجاهل الصفوف الفارغة
      if (!barcode && !designation && !Number.isFinite(price)) {
        continue;
      }

      if (!barcode || !designation || !Number.isFinite(price)) {
        invalidRows++;
        continue;
      }

      products.set(barcode, {
        barcode,
        designation,
        price
      });
    }

    if (products.size === 0) {
      status.textContent = "لم يتم العثور على منتجات صالحة";
      status.classList.remove("loaded");

      showNotice(
        "تم فتح ملف Excel، لكن لم تُقرأ أي منتجات. تحقق من أسماء الأعمدة والبيانات."
      );

      console.warn("Excel headers:", Object.keys(rows[0] ?? {}));
      console.warn("Excel rows:", rows);

      return;
    }

    status.textContent =
      `قاعدة البيانات جاهزة • ${products.size} منتج`;

    status.classList.add("loaded");

    if (invalidRows > 0) {
      showNotice(
        `تم تحميل ${products.size} منتج. تم تجاهل ${invalidRows} صف غير مكتمل.`,
        true
      );
    } else {
      showNotice("تم تحميل قاعدة المنتجات بنجاح.", true);
    }

    console.log(`تم تحميل ${products.size} منتج من Excel.`);
  } catch (error) {
    console.error("خطأ تحميل قاعدة المنتجات:", error);

    if (status) {
      status.textContent = "تعذر تحميل قاعدة المنتجات";
      status.classList.remove("loaded");
    }

    showNotice(
      "تعذر تحميل قاعدة المنتجات. تحقق من رابط Excel وأسماء الأعمدة واتصال الإنترنت."
    );
  }
}

function formatPrice(price) {
  return new Intl.NumberFormat("fr-DZ", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  }).format(price);
}

function showNotice(message, good = false) {
  const notice = $("notice");

  if (!notice) {
    console.warn(message);
    return;
  }

  notice.textContent = message;
  notice.style.color = good ? "#16896d" : "#d34d69";
}

function searchProduct(value) {
  const barcode = normalizeBarcode(value);

  if (!barcode) {
    showNotice("أدخل رقم الباركود أولاً.");
    return;
  }

  $("barcode").value = barcode;

  const product = products.get(barcode);

  if (!product) {
    $("product-name").textContent = "المنتج غير موجود";
    $("product-price").textContent = "—";
    $("product-barcode").textContent = `الكود: ${barcode}`;

    showNotice(
      products.size === 0
        ? "قاعدة المنتجات غير محمّلة. تحقق من حالة قاعدة البيانات."
        : "لم نعثر على هذا الكود. تحقق من الرقم أو حدّث ملف Excel."
    );

    return;
  }

  $("product-name").textContent = product.designation;
  $("product-price").textContent = formatPrice(product.price);
  $("product-barcode").textContent = `CODE BARRE : ${product.barcode}`;

  showNotice("تم العثور على المنتج بنجاح.", true);

  $("result").scrollIntoView({
    behavior: "smooth",
    block: "nearest"
  });
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

    if (!cameras || cameras.length === 0) {
      throw new Error("لم يتم العثور على كاميرا.");
    }

    const preferred =
      cameras.find((camera) =>
        /back|rear|environment|arrière|خلف/i.test(camera.label)
      ) || cameras[cameras.length - 1];

    $("reader").style.minHeight = "220px";

    await scanner.start(
      preferred.id,
      {
        fps: 10,
        qrbox: (width, height) => ({
          width: Math.min(300, Math.floor(width * 0.88)),
          height: Math.min(140, Math.floor(height * 0.55))
        }),
        aspectRatio: 1.6
      },
      (decodedText) => {
        searchProduct(decodedText);
        stopScanner();
      },
      () => {}
    );

    scanning = true;

    $("start-scan").disabled = true;
    $("stop-scan").disabled = false;

    showNotice("الكاميرا تعمل. وجّهها نحو الباركود.", true);
  } catch (error) {
    console.error("خطأ تشغيل الكاميرا:", error);

    showNotice(
      "تعذر فتح الكاميرا. اسمح للمتصفح باستخدامها وافتح الموقع عبر HTTPS."
    );

    if (scanner) {
      try {
        await scanner.clear();
      } catch (_) {}

      scanner = null;
    }

    scanning = false;

    $("start-scan").disabled = false;
    $("stop-scan").disabled = true;
  }
}

async function stopScanner() {
  if (scanner) {
    try {
      if (scanning) {
        await scanner.stop();
      }
    } catch (error) {
      console.warn("تعذر إيقاف الكاميرا بشكل طبيعي:", error);
    }

    try {
      await scanner.clear();
    } catch (_) {}
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

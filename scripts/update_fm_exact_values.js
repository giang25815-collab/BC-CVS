const fs = require('fs');

const FM_EXACT_VALUES = {
  "201, Nguyễn Khuyến, , P.Trảng Dài, Đồng Nai": 1816100,
  "06, Hoàng Hoa Thám, , P.Vũng Tàu, Hồ Chí Minh": 1930350,
  "148, Hoàng Hoa Thám, , P.Vũng Tàu, Hồ Chí Minh": 2628150,
  "171, Nam Kỳ Khởi Nghĩa, P.6, TP.Vũng Tàu, Bà Rịa Vũng Tàu": 2735700,
  "26D, Lê Hồng Phong, , P.Tam Thắng, Hồ Chí Minh": 3307200,
  "48, Đồ Chiểu, , P.Vũng Tàu, Hồ Chí Minh": 692400,
  "58, Trương Công Định, , P.Vũng Tàu, Hồ Chí Minh": 2907600,
  "Số 816, Trần Phú, , P.Vũng Tàu, Hồ Chí Minh": 1909500,
  "102 Tầng 1+2 Tòa nhà Phú Hòa,, Số 117, Đ.D1, Khu 7, , P.Phú Lợi, Hồ Chí Minh": 2679150,
  "111, Hoàng Văn Thụ, , P.Thủ Dầu Một, Hồ Chí Minh": 855750,
  "126, Hoàng Hoa Thám, , P.Phú Lợi, Hồ Chí Minh": 4786500,
  "166, Trần Văn Ơn, , P.Phú Lợi, Hồ Chí Minh": 2925450,
  "19, Đường TC3-XC2, , P.Bến Cát, Hồ Chí Minh": 2459250,
  "239, Hoàng Văn Thụ, , P.Thủ Dầu Một, Hồ Chí Minh": 3793500,
  "26/13, 1, KP Hòa Long, , P.Bình Hòa, Hồ Chí Minh": 1248600,
  "29-33, Trần Hưng Đạo, , P.Thủ Dầu Một, Hồ Chí Minh": 1423050,
  "30/3, Nguyễn Văn Tiết, , P.Lái Thiêu, Hồ Chí Minh": 825600,
  "7, Phú Lợi, , P.Phú Lợi, Hồ Chí Minh": 1681350,
  "B2,, Đường Hùng Vương, , P.Bình Dương, Hồ Chí Minh": 1363800,
  "G05, Tầng trệt, Block CT2, ĐL Bình Dương, KP Hưng Lộc, , P.Thuận An, Hồ Chí Minh": 2112750,
  "Khu đô thị vườn Tokyu BD, Lô H8,, Đường Tạo Lực 5, H. Phú, , P.Bình Dương, Hồ Chí Minh": 2163150,
  "L6 A11, A12, A16, Lý Thái Tổ, , P.Bình Dương, Hồ Chí Minh": 2432100,
  "Lô C18, Đại lộ Hùng Vương, , P.Bình Dương, Hồ Chí Minh": 8824300,
  "Số 203, Lê Hồng Phong, KP8,, , P.Phú Lợi, Hồ Chí Minh": 1297050,
  "Số 356,, Đường 30/4,, , P.Thủ Dầu Một, Hồ Chí Minh": 2552100,
  "Số 490, ĐL Bình Dương, , P.Phú Lợi, Hồ Chí Minh": 919950,
  "Số 62/2, Đường 745, KP. Thạnh Lợi, , P.Thuận An, Hồ Chí Minh": 981450,
  "Tầng trệt, Block D, KDC Hiệp Thành III, Tổ 105, KP.7, , P.Phú Lợi, Hồ Chí Minh": 5727750,
  "Ô 1 và ô 2, lô DC37, Khu dân cư Vietsing, , P.An Phú, Hồ Chí Minh": 3770850,
  "Ô 18 Khu Đô Thị Becamex, 30/04, , P.Phú Lợi, Hồ Chí Minh": 1691700,
  "Ô R4, JF1A Khu đô thị mới, , P.Bình Dương, Hồ Chí Minh": 2768850,
  "198-198A, Phan Đình Phùng, , P.Trấn Biên, Đồng Nai": 1705800,
  "2, Nguyễn An Ninh, , P.Dĩ An, Hồ Chí Minh": 2663700,
  "A2, Trần Quốc Toản, , P.Tam Hiệp, Đồng Nai": 2597550,
  "Căn 03 Marina Retails, Phân khu 8 thuộc dự án Hoa Sen Đại Phước, , X.Đại Phước, Đồng Nai": 0
};

let total = 0;
for (let k in FM_EXACT_VALUES) total += FM_EXACT_VALUES[k];
console.log('Total FamilyMart exact sum:', total);

function updateFileContent(filepath) {
  if (!fs.existsSync(filepath)) return;
  let content = fs.readFileSync(filepath, 'utf8');
  let updatedStores = 0;
  let updatedFm = 0;

  for (let addr in FM_EXACT_VALUES) {
    const val = FM_EXACT_VALUES[addr];
    // Pattern 1: in DEFAULT_FM_STORES: "addr": "...", ... "actual": 0.0
    const escapedAddr = addr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    
    // In DEFAULT_FM_STORES
    const pat1 = new RegExp(`("addr":\\s*"${escapedAddr}"[\\s\\S]*?"actual":\\s*)[0-9.]+`, 'g');
    if (pat1.test(content)) {
      content = content.replace(pat1, `$1${val}`);
      updatedFm++;
    }
    
    // In MASTER_DATA.stores
    const pat2 = new RegExp(`("store_address":\\s*"${escapedAddr}"[\\s\\S]*?"actual":\\s*)[0-9.]+`, 'g');
    if (pat2.test(content)) {
      content = content.replace(pat2, `$1${val}`);
      updatedStores++;
    }
  }
  
  fs.writeFileSync(filepath, content, 'utf8');
  console.log(`Updated ${filepath}: DEFAULT_FM_STORES = ${updatedFm}, MASTER_DATA.stores = ${updatedStores}`);
}

['master_data.js', 'index.html', 'MasterData.gs', 'CRypto.html'].forEach(updateFileContent);

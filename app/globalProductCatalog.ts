export type GlobalProductCategory = {
  id: string;
  name: string;
  globalName: string;
  keywords: string[];
};

const TOBACCO_BRANDS = [
  "中华", "芙蓉王", "利群", "玉溪", "黄鹤楼", "云烟", "南京", "双喜", "红塔山", "白沙", "娇子", "黄金叶", "苏烟", "泰山", "七匹狼", "中南海", "牡丹", "贵烟", "真龙", "钻石", "煊赫门", "万宝路", "黄山", "长白山", "延安", "兰州", "宽窄", "荷花", "天子", "红河", "红金龙", "金圣", "人民大会堂", "雨花石", "红旗渠", "工字牌", "五叶神", "滕王阁", "好猫", "金猴子", "金丝猴", "黄果树", "都宝", "大前门", "大重九", "阿诗玛", "石林", "将军", "熊猫", "大熊猫", "上海", "恒大", "哈德门", "红双喜", "天下秀", "壹枝笔", "小熊猫", "白将", "红将", "紫云", "云龙", "茶花", "恭贺新禧", "福牌", "三五", "555", "骆驼", "camel", "marlboro", "万宝路", "kent", "健牌", "esse", "爱喜", "mevius", "七星", "sevenstars", "winston", "温斯顿", "davidoff", "大卫杜夫", "parliament", "百乐门", "luckystrike", "好彩", "dunhill", "登喜路", "pallmall", "百乐门", "rothmans", "more", "sobranie", "vogue", "chesterfield", "ld", "l&m", "lm", "bond", "blackdevil", "captainblack", "heets", "iqos", "heets"
];

const BEVERAGE_BRANDS = [
  "可乐", "可口可乐", "cocacola", "coke", "cola", "百事", "百事可乐", "pepsi", "雪碧", "sprite", "芬达", "fanta", "美年达", "mirinda", "七喜", "7up", "红牛", "redbull", "东鹏", "东鹏特饮", "脉动", "mizone", "宝矿力", "宝矿力水特", "pocari", "佳得乐", "gatorade", "外星人", "战马", "乐虎", "体质能量", "尖叫", "健力宝", "元气森林", "元气", "燃茶", "气泡水", "苏打水", "巴黎水", "perrier", "圣培露", "sanpellegrino", "农夫山泉", "怡宝", "百岁山", "景田", "冰露", "娃哈哈", "屈臣氏", "evian", "依云", "volvic", "voss", "恒大冰泉", "康师傅", "冰红茶", "绿茶", "茉莉清茶", "茉莉蜜茶", "鲜果橙", "水蜜桃", "龙井", "康师傅900", "统一", "统一冰红茶", "阿萨姆", "茶π", "茶派", "茶pai", "东方树叶", "茶里王", "小茗同学", "燃茶", "三得利", "乌龙茶", "无糖茶", "海之言", "王老吉", "加多宝", "和其正", "椰树", "椰汁", "特种兵", "露露", "六个核桃", "营养快线", "ad钙", "旺仔", "旺仔牛奶", "伊利", "蒙牛", "纯甄", "安慕希", "光明", "雀巢", "nescafe", "星巴克", "starbucks", "瑞幸", "luckin", "咖啡", "coffee", "果粒橙", "美汁源", "minute maid", "汇源", "农夫果园", "椰子水", "coconut", "monster", "魔爪", "vitaminwater", "powerade", "lipton", "立顿", "冰峰", "北冰洋", "大窑", "宏宝莱", "亚洲沙示", "sarsi", "维他", "维他柠檬茶", "vitasoy", "vita", "if", "oatly", "燕麦奶", "啤儿茶爽", "清凉茶", "凉茶", "饮料", "矿泉水", "纯净水", "汽水", "茶饮", "奶茶", "果汁", "乳酸菌", "酸奶", "牛奶", "豆奶", "豆浆", "能量饮料", "运动饮料", "water", "soda", "juice", "tea", "milk", "yogurt", "energy drink", "drink", "beverage", "напиток", "вода", "чай", "сок", "молоко", "مشروب", "ماء", "شاي", "عصير", "حليب"
];

export const GLOBAL_PRODUCT_CATEGORIES: GlobalProductCategory[] = [
  {
    id: "food-beverage",
    name: "食品饮料",
    globalName: "Food & Beverages",
    keywords: [
      ...BEVERAGE_BRANDS,
      "方便面", "泡面", "螺蛳粉", "noodle", "ramen", "零食", "snack", "饼干", "cookie", "巧克力", "chocolate", "糖", "candy", "口香糖", "gum", "面包", "bread", "饼", "坚果", "nut", "薯片", "chips"
    ]
  },
  {
    id: "tobacco",
    name: "烟草烟具",
    globalName: "Tobacco & Smoking Accessories",
    keywords: [
      "烟", "香烟", "卷烟", "烟草", "细支", "中支", "爆珠", "硬盒", "软盒", "打火机", "火机", "火柴", "点烟器", "cigarette", "tobacco", "smoke", "lighter", "зажигалка", "сигарета", "табак", "ولاعة", "سجائر", "تبغ",
      ...TOBACCO_BRANDS
    ]
  },
  { id: "daily-necessities", name: "日用百货", globalName: "Daily Necessities", keywords: ["纸巾", "纸", "抽纸", "卷纸", "湿巾", "牙刷", "牙膏", "毛巾", "洗衣", "清洁", "洗洁精", "拖鞋", "雨伞", "电池", "袋子", "垃圾袋", "杯子", "水杯", "日用品", "百货", "household", "tissue", "paper", "toothbrush", "toothpaste", "towel", "umbrella", "battery", "cleaner"] },
  { id: "pet-supplies", name: "宠物用品", globalName: "Pet Supplies", keywords: ["宠物", "狗", "狗狗", "猫咪", "宠物衣服", "宠物服装", "狗衣服", "猫衣服", "牵引绳", "项圈", "宠物窝", "猫砂", "狗粮", "猫粮", "pet", "dog", "cat", "puppy", "kitten", "leash", "collar", "pet clothes", "pet apparel"] },
  { id: "apparel", name: "服装鞋帽", globalName: "Apparel & Accessories", keywords: ["衣服", "服装", "上衣", "裤", "裙", "外套", "夹克", "羽绒", "鞋", "帽", "袜", "内衣", "tshirt", "t-shirt", "shirt", "pants", "jeans", "dress", "jacket", "coat", "shoes", "hat", "socks"] },
  { id: "jewelry", name: "珠宝饰品", globalName: "Jewelry & Accessories", keywords: ["珠宝", "饰品", "首饰", "戒指", "项链", "耳环", "耳钉", "手链", "手镯", "吊坠", "玉髓", "银", "高碳钻", "宝石", "jewelry", "jewellery", "ring", "necklace", "earring", "bracelet", "bangle", "pendant", "gem", "silver", "zircon"] },
  { id: "home-furniture", name: "家居家具", globalName: "Home & Furniture", keywords: ["家居", "家具", "床", "床垫", "枕头", "被子", "沙发", "桌", "椅", "柜", "灯", "窗帘", "mattress", "bed", "pillow", "sofa", "table", "chair", "cabinet", "lamp", "curtain"] },
  { id: "electronics", name: "电子数码", globalName: "Electronics", keywords: ["电子", "数码", "手机", "充电器", "数据线", "耳机", "音箱", "电源", "插座", "转换头", "键盘", "鼠标", "屏幕", "phone", "charger", "cable", "usb", "earphone", "headphone", "speaker", "power bank", "keyboard", "mouse", "adapter"] },
  { id: "beauty-care", name: "美妆个护", globalName: "Beauty & Personal Care", keywords: ["美妆", "护肤", "洗发", "沐浴", "香水", "口红", "面膜", "乳液", "精华", "剃须", "洗面奶", "化妆", "beauty", "skin", "shampoo", "body wash", "perfume", "lipstick", "mask", "cream", "razor", "cosmetic"] },
  { id: "mother-baby", name: "母婴用品", globalName: "Mother & Baby", keywords: ["母婴", "宝宝", "婴儿", "奶瓶", "奶嘴", "纸尿裤", "尿不湿", "玩具", "童装", "baby", "infant", "diaper", "bottle", "toy", "kids", "children"] },
  { id: "office-school", name: "办公文具", globalName: "Office & School Supplies", keywords: ["办公", "文具", "笔", "本子", "纸张", "打印", "文件夹", "胶带", "剪刀", "订书机", "office", "pen", "notebook", "paper", "printer", "folder", "tape", "scissors", "stapler"] },
  { id: "auto-parts", name: "汽车用品", globalName: "Auto Parts & Accessories", keywords: ["汽车", "车载", "轮胎", "机油", "雨刷", "脚垫", "车灯", "充气泵", "香薰", "auto", "car", "tire", "oil", "wiper", "car mat"] },
  { id: "tools-hardware", name: "五金工具", globalName: "Tools & Hardware", keywords: ["五金", "工具", "螺丝", "螺丝刀", "扳手", "锤", "钳", "电钻", "锁", "tool", "hardware", "screw", "screwdriver", "wrench", "hammer", "plier", "drill", "lock"] },
  { id: "sports-outdoor", name: "运动户外", globalName: "Sports & Outdoors", keywords: ["运动", "户外", "健身", "瑜伽", "球", "帐篷", "背包", "水壶", "sport", "fitness", "yoga", "ball", "tent", "backpack", "bottle"] }
];

export function normalizeProductText(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "");
}

const STRONG_TOBACCO_PATTERN = new RegExp(TOBACCO_BRANDS.map(normalizeProductText).filter(Boolean).join("|"));
const STRONG_BEVERAGE_PATTERN = new RegExp(BEVERAGE_BRANDS.map(normalizeProductText).filter((item) => item.length >= 2).join("|"));

export function classifyGlobalProduct(value: unknown) {
  const text = normalizeProductText(value);
  if (!text) return { category: "未识别", confidence: 0, matchedKeyword: "" };

  const strongTobacco = text.match(STRONG_TOBACCO_PATTERN);
  if (strongTobacco) return { category: "烟草烟具", confidence: 0.99, matchedKeyword: strongTobacco[0] };

  const strongBeverage = text.match(STRONG_BEVERAGE_PATTERN);
  if (strongBeverage) return { category: "食品饮料", confidence: 0.98, matchedKeyword: strongBeverage[0] };

  let best = { category: "未识别", confidence: 0, matchedKeyword: "" };
  for (const category of GLOBAL_PRODUCT_CATEGORIES) {
    for (const keyword of category.keywords) {
      const normalizedKeyword = normalizeProductText(keyword);
      if (!normalizedKeyword || normalizedKeyword.length < 2) continue;
      if (text.includes(normalizedKeyword)) {
        const confidence = Math.min(0.96, Math.max(0.55, normalizedKeyword.length / Math.max(text.length, 1) + 0.45));
        if (confidence > best.confidence) best = { category: category.name, confidence, matchedKeyword: keyword };
      }
    }
  }
  return best;
}

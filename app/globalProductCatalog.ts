export type GlobalProductCategory = {
  id: string;
  name: string;
  globalName: string;
  keywords: string[];
};

export const GLOBAL_PRODUCT_CATEGORIES: GlobalProductCategory[] = [
  {
    id: "food-beverage",
    name: "食品饮料",
    globalName: "Food & Beverages",
    keywords: ["饮料", "水", "矿泉水", "纯净水", "可乐", "cola", "coke", "coca", "pepsi", "雪碧", "sprite", "芬达", "fanta", "红牛", "redbull", "咖啡", "coffee", "茶", "tea", "果汁", "juice", "牛奶", "milk", "酸奶", "yogurt", "方便面", "泡面", "螺蛳粉", "noodle", "ramen", "零食", "snack", "饼干", "cookie", "巧克力", "chocolate", "糖", "candy", "口香糖", "gum", "面包", "bread", "饼", "坚果", "nut", "薯片", "chips", "苏打", "soda", "энергетик", "напиток", "вода", "кофе", "чай", "сок", "молоко", "مشروب", "ماء", "قهوة", "شاي", "عصير"]
  },
  {
    id: "tobacco",
    name: "烟草烟具",
    globalName: "Tobacco & Smoking Accessories",
    keywords: [
      "烟", "香烟", "卷烟", "烟草", "细支", "中支", "爆珠", "硬盒", "软盒", "硬", "软",
      "中华", "芙蓉王", "利群", "玉溪", "黄鹤楼", "云烟", "南京", "双喜", "红塔山", "白沙", "娇子", "黄金叶", "苏烟", "泰山", "七匹狼", "中南海", "牡丹", "贵烟", "真龙", "钻石", "煊赫门", "万宝路", "黄山", "长白山", "延安", "兰州", "宽窄", "荷花", "天子", "红河", "红金龙", "金圣", "人民大会堂", "雨花石",
      "红旗渠", "工字牌", "五叶神", "滕王阁", "好猫", "金猴子", "金丝猴", "黄果树", "都宝", "大前门", "大重九", "阿诗玛", "石林", "将军", "熊猫", "大熊猫", "上海", "恒大", "哈德门", "红双喜", "黄鹤楼", "天下秀", "利群", "白将", "红将", "壹枝笔", "小熊猫", "云龙", "云烟", "茶花", "紫云", "软蓝", "硬蓝", "软珍", "硬珍", "软红", "硬红", "软金", "硬金", "蓝", "金", "白", "黑", "红",
      "marlboro", "cigarette", "tobacco", "smoke", "lighter", "打火机", "火机", "火柴", "点烟器", "зажигалка", "сигарета", "табак", "ولاعة", "سجائر", "تبغ"
    ]
  },
  {
    id: "daily-necessities",
    name: "日用百货",
    globalName: "Daily Necessities",
    keywords: ["纸巾", "纸", "抽纸", "卷纸", "湿巾", "牙刷", "牙膏", "毛巾", "洗衣", "清洁", "洗洁精", "拖鞋", "雨伞", "电池", "袋子", "垃圾袋", "杯子", "水杯", "日用品", "百货", "household", "tissue", "paper", "toothbrush", "toothpaste", "towel", "umbrella", "battery", "cleaner", "полотенце", "зубная", "бумага", "батарейка", "مناديل", "فرشاة", "معجون", "منشفة"]
  },
  {
    id: "pet-supplies",
    name: "宠物用品",
    globalName: "Pet Supplies",
    keywords: ["宠物", "狗", "猫", "犬", "狗狗", "猫咪", "宠物衣服", "宠物服装", "狗衣服", "猫衣服", "雨衣", "牵引绳", "项圈", "宠物窝", "猫砂", "狗粮", "猫粮", "pet", "dog", "cat", "puppy", "kitten", "leash", "collar", "pet clothes", "pet apparel", "корм", "собак", "кош", "поводок", "ошейник", "حيوان", "كلب", "قط", "طوق"]
  },
  {
    id: "apparel",
    name: "服装鞋帽",
    globalName: "Apparel & Accessories",
    keywords: ["衣服", "服装", "上衣", "裤", "裙", "外套", "夹克", "羽绒", "鞋", "帽", "袜", "内衣", "tshirt", "t-shirt", "shirt", "pants", "jeans", "dress", "jacket", "coat", "shoes", "hat", "socks", "одежда", "рубашка", "брюки", "платье", "куртка", "обувь", "ملابس", "قميص", "بنطال", "حذاء"]
  },
  {
    id: "jewelry",
    name: "珠宝饰品",
    globalName: "Jewelry & Accessories",
    keywords: ["珠宝", "饰品", "首饰", "戒指", "项链", "耳环", "耳钉", "手链", "手镯", "吊坠", "玉髓", "银", "高碳钻", "宝石", "jewelry", "jewellery", "ring", "necklace", "earring", "bracelet", "bangle", "pendant", "gem", "silver", "zircon", "украш", "кольцо", "ожерелье", "серьги", "браслет", "مجوهرات", "خاتم", "قلادة", "سوار"]
  },
  {
    id: "home-furniture",
    name: "家居家具",
    globalName: "Home & Furniture",
    keywords: ["家居", "家具", "床", "床垫", "枕头", "被子", "沙发", "桌", "椅", "柜", "灯", "窗帘", "mattress", "bed", "pillow", "sofa", "table", "chair", "cabinet", "lamp", "curtain", "матрас", "кровать", "подушка", "диван", "стол", "стул", "أثاث", "مرتبة", "سرير", "وسادة", "أريكة"]
  },
  {
    id: "electronics",
    name: "电子数码",
    globalName: "Electronics",
    keywords: ["电子", "数码", "手机", "充电器", "数据线", "耳机", "音箱", "电源", "插座", "转换头", "键盘", "鼠标", "屏幕", "phone", "charger", "cable", "usb", "earphone", "headphone", "speaker", "power bank", "keyboard", "mouse", "adapter", "кабель", "заряд", "телефон", "наушник", "адаптер", "شاحن", "كابل", "هاتف", "سماعة"]
  },
  {
    id: "beauty-care",
    name: "美妆个护",
    globalName: "Beauty & Personal Care",
    keywords: ["美妆", "护肤", "洗发", "沐浴", "香水", "口红", "面膜", "乳液", "精华", "剃须", "洗面奶", "化妆", "beauty", "skin", "shampoo", "body wash", "perfume", "lipstick", "mask", "cream", "razor", "cosmetic", "космет", "шампунь", "духи", "крем", "бритва", "تجميل", "عطر", "شامبو", "كريم"]
  },
  {
    id: "mother-baby",
    name: "母婴用品",
    globalName: "Mother & Baby",
    keywords: ["母婴", "宝宝", "婴儿", "奶瓶", "奶嘴", "纸尿裤", "尿不湿", "玩具", "童装", "baby", "infant", "diaper", "bottle", "toy", "kids", "children", "детск", "ребен", "подгуз", "игруш", "طفل", "حفاض", "لعبة"]
  },
  {
    id: "office-school",
    name: "办公文具",
    globalName: "Office & School Supplies",
    keywords: ["办公", "文具", "笔", "本子", "纸张", "打印", "文件夹", "胶带", "剪刀", "订书机", "office", "pen", "notebook", "paper", "printer", "folder", "tape", "scissors", "stapler", "ручка", "тетрад", "бумага", "папка", "قلم", "دفتر", "ورق"]
  },
  {
    id: "auto-parts",
    name: "汽车用品",
    globalName: "Auto Parts & Accessories",
    keywords: ["汽车", "车载", "轮胎", "机油", "雨刷", "脚垫", "车灯", "充气泵", "香薰", "auto", "car", "tire", "oil", "wiper", "car mat", "авто", "машин", "шина", "масло", "سيارة", "إطار", "زيت"]
  },
  {
    id: "tools-hardware",
    name: "五金工具",
    globalName: "Tools & Hardware",
    keywords: ["五金", "工具", "螺丝", "螺丝刀", "扳手", "锤", "钳", "电钻", "锁", "tool", "hardware", "screw", "screwdriver", "wrench", "hammer", "plier", "drill", "lock", "инструмент", "винт", "молоток", "дрель", "أداة", "مطرقة", "مسمار"]
  },
  {
    id: "sports-outdoor",
    name: "运动户外",
    globalName: "Sports & Outdoors",
    keywords: ["运动", "户外", "健身", "瑜伽", "球", "帐篷", "背包", "水壶", "sport", "fitness", "yoga", "ball", "tent", "backpack", "bottle", "спорт", "фитнес", "палатка", "рюкзак", "رياضة", "خيمة", "حقيبة"]
  }
];

export function normalizeProductText(value: unknown) {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "");
}

const STRONG_TOBACCO_PATTERN = /(中华|芙蓉王|利群|玉溪|黄鹤楼|云烟|南京|双喜|红塔山|白沙|娇子|黄金叶|苏烟|泰山|七匹狼|中南海|牡丹|贵烟|真龙|钻石|煊赫门|万宝路|黄山|长白山|延安|兰州|宽窄|荷花|天子|红河|红金龙|金圣|人民大会堂|雨花石|红旗渠|工字牌|五叶神|滕王阁|好猫|金猴子|金丝猴|黄果树|都宝|大前门|大重九|阿诗玛|石林|将军|熊猫|大熊猫|上海|恒大|哈德门|红双喜|天下秀|小熊猫)/;

export function classifyGlobalProduct(value: unknown) {
  const text = normalizeProductText(value);
  if (!text) return { category: "未识别", confidence: 0, matchedKeyword: "" };

  const strongTobacco = text.match(STRONG_TOBACCO_PATTERN);
  if (strongTobacco) return { category: "烟草烟具", confidence: 0.99, matchedKeyword: strongTobacco[0] };

  let best = { category: "未识别", confidence: 0, matchedKeyword: "" };
  for (const category of GLOBAL_PRODUCT_CATEGORIES) {
    for (const keyword of category.keywords) {
      const normalizedKeyword = normalizeProductText(keyword);
      if (!normalizedKeyword) continue;
      if (text.includes(normalizedKeyword)) {
        const confidence = Math.min(0.98, Math.max(0.55, normalizedKeyword.length / Math.max(text.length, 1) + 0.45));
        if (confidence > best.confidence) best = { category: category.name, confidence, matchedKeyword: keyword };
      }
    }
  }
  return best;
}

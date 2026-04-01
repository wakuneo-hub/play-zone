/* ==========================================================
   Play-Zone 出品ツール  v30  (2026-03-31)
   楽天テキスト → Amazon出品ファイル(.xlsm)生成
   ========================================================== */

// ─── テンプレート列マッピング (成功ファイル amazon_B024_upload.xlsm から取得, 1-based) ───
const COL = {
  SKU:         1,    // A  - contribution_sku
  productType: 2,    // B  - product_type
  recordAction:3,    // C  - record_action
  parentage:   4,    // D  - parentage_level
  parentSku:   5,    // E  - child_parent_sku
  varTheme:    6,    // F  - variation_theme
  itemName:    7,    // G  - item_name
  brand:       8,    // H  - brand
  idType:      9,    // I  - product_id_type
  idValue:     10,   // J  - product_id_value
  browseNode1: 11,   // K  - recommended_browse_nodes
  packageUnit: 16,   // P  - package_level (単品)
  targetAudience1: 20, // T - target_audience
  modelNumber: 25,   // Y  - model_number (品番)
  modelName:   26,   // Z  - model_name
  manufacturer:27,   // AA - manufacturer
  mainImage:   30,   // AD - main_product_image
  otherImage1: 31,   // AE
  otherImage2: 32,   // AF
  otherImage3: 33,   // AG
  otherImage4: 34,   // AH
  otherImage5: 35,   // AI
  otherImage6: 36,   // AJ
  otherImage7: 37,   // AK
  description: 40,   // AN - product_description
  bullet1:     41,   // AO
  bullet2:     42,   // AP
  bullet3:     43,   // AQ
  bullet4:     44,   // AR
  bullet5:     45,   // AS
  keywords:    46,   // AT - generic_keyword
  lifestyle:   52,   // AZ - lifestyle
  style:       53,   // BA - style
  department:  54,   // BB - department
  gender:      55,   // BC - target_gender
  ageRange:    56,   // BD - age_range_description
  material1:   57,   // BE - material
  fabricType:  60,   // BH - fabric_type
  specialSize: 64,   // BL - special_size_type
  colorMap:    65,   // BM - color (standardized)
  color:       66,   // BN - color (value)
  size:        67,   // BO - size
  occasion1:   73,   // BU - occasion_type (ハロウィン)
  occasion2:   74,   // BV - occasion_type (カーニバル)
  partNumber:  78,   // BZ - part_number (メーカー型番)
  materialComp:82,   // CD - material_composition
  careInst:    83,   // CE - care_instructions
  importType:  84,   // CF - distribution_designation
  isExclusive: 86,   // CH - is_exclusive_product
  setName:     141,  // EK - set_name
  condition:   152,  // EV - condition_type
  listPrice:   154,  // EX - list_price
  fulfillment: 178,  // FV - fulfillment_channel_code
  quantity:    179,  // FW - quantity
  price:       183,  // GA - our_price (販売価格)
  shipping:    208,  // GZ - merchant_shipping_group
  countryOfOrigin: 218, // HJ - country_of_origin
  battery:     219,  // HK - batteries_required
  hazmat:      235,  // IA - supplier_declared_dg_hz_regulation
};

// ─── 固定値 ───
const FIXED = {
  recordAction: '作成または置換 (完全更新)',
  productType:  'COSTUME_OUTFIT',
  brand:        'Play-Zone(プレイゾーン)',
  browseNode:   '10345693051',
  conditionVal: '新品',
  manufacturer: 'play-zone',
  packageUnit:  '単品',
  modelName:    'イベント衣装',
  occasion1:    'ハロウィン',
  occasion2:    'カーニバル',
  careInst:     'ドライクリーニングのみ',
  countryOfOrigin: '中国',
  battery:      'いいえ',
  targetAud:    'ユニセックス(大人)',
  lifestyle:    'カジュアル',
  style:        'コンテンポラリー',
  department:   'ユニセックス',
  gender:       '女性',
  ageRange:     '大人',
  specialSize:  '標準',
  importType:   '正規品',
  isExclusive:  'いいえ',
  fulfillment:  'DEFAULT',
  shipping:     'プライム配送パターン',
  hazmat:       '該当なし',
};

// ─── Toast ───
function showToast(msg, type) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = 'toast ' + (type || 'success');
  clearTimeout(el._tid);
  el._tid = setTimeout(() => el.classList.add('hidden'), 3000);
}

// ─── 日付フォーマット ───
function formatDate() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}`;
}

// ─── セル値クリーン ───
function cleanDataCell(val) {
  return String(val ?? '')
    .replace(/<[^>]*>/g, '')
    .replace(/[\r\n]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/* ==========================================================
   楽天テキスト解析
   実際の楽天ページからコピーされるテキスト形式:
   - 商品名（最初の長いテキスト行）
   - 商品番号：XXXX
   - X,XXX円 送料無料
   - 商品情報（テーブル: ラベル\t値 or ラベル 値）
     セット内容 / カラー / サイズ / 素材
   - その他商品説明（■で始まる箇条書き）
   - 注意書き（※で始まる）
   - 梱包について
   ========================================================== */
function parseRakutenText(text) {
  const raw = (text || '')
    .replace(/<br\s*\/?>/gi, '\n')   // <br>タグを改行に変換
    .replace(/<[^>]*>/g, '')          // その他HTMLタグを除去
    .replace(/&nbsp;/gi, ' ')
    .replace(/\r\n/g, '\n').replace(/\r/g, '\n')
    .replace(/　/g, ' ')
    .trim();
  if (!raw) return null;

  const lines = raw.split('\n').map(l => l.trim()).filter(Boolean);

  // デバッグ: 最初の30行を表示して構造を確認
  console.log('[DEBUG] === 楽天テキスト解析 ===');
  console.log('[DEBUG] 全行数:', lines.length);
  lines.slice(0, 30).forEach((l, i) => console.log(`[DEBUG] Line ${i}: ${l.slice(0, 80)}`));

  const result = {
    parentSku: '',
    itemName: '',
    price: '',
    listPrice: '',
    colors: [],
    sizes: [],
    material: '',
    fabricType: '',
    description: '',
    bullets: [],
    keywords: '',
    setContents: '',
  };

  // ─── 商品番号（SKU）抽出 ───
  for (const line of lines) {
    const skuMatch = line.match(/商品番号\s*[:：]\s*(\S+)/);
    if (skuMatch) {
      result.parentSku = skuMatch[1].trim();
      break;
    }
  }

  // ─── 価格抽出 ───
  // 楽天ページからコピーすると「2,180」と「円」と「送料無料」が別々の行になる
  // まずlinesを結合して隣接行をまとめてからパターンマッチ
  {
    let found = false;

    // 方法1: 「数字行」→次の行が「円」→その次が「送料無料」のパターン
    for (let i = 0; i < lines.length - 1; i++) {
      const numMatch = lines[i].match(/^(\d[\d,]+)$/);  // 数字だけの行
      if (numMatch && /^円/.test(lines[i + 1])) {
        // 次の行が「円」で始まる → この数字が価格
        result.price = numMatch[1].replace(/,/g, '');
        found = true;
        console.log('[DEBUG] 方法1(数字行+円行)で価格取得:', result.price);
        break;
      }
    }

    // 方法2: 同じ行に「X,XXX円」がある場合（送料無料も同じ行にある場合を優先）
    if (!found) {
      for (const line of lines) {
        const m = line.match(/(\d[\d,]+)\s*円/);
        if (m && /(送料無料|送料込)/.test(line)) {
          result.price = m[1].replace(/,/g, '');
          found = true;
          console.log('[DEBUG] 方法2(同一行に円+送料無料)で価格取得:', result.price);
          break;
        }
      }
    }

    // 方法3: 「送料無料」の直前数行にある数字を価格とする
    if (!found) {
      const freeShipIdx = lines.findIndex(l => /^送料無料|^送料込/.test(l));
      if (freeShipIdx > 0) {
        for (let i = freeShipIdx - 1; i >= Math.max(0, freeShipIdx - 3); i--) {
          // 数字だけの行、または「X,XXX円」の行
          const m1 = lines[i].match(/^(\d[\d,]+)$/);
          const m2 = lines[i].match(/(\d[\d,]+)\s*円/);
          if (m1) {
            result.price = m1[1].replace(/,/g, '');
            found = true;
            console.log('[DEBUG] 方法3(送料無料の前)で価格取得:', result.price);
            break;
          } else if (m2) {
            result.price = m2[1].replace(/,/g, '');
            found = true;
            console.log('[DEBUG] 方法3(送料無料の前+円)で価格取得:', result.price);
            break;
          }
        }
      }
    }

    // 方法4: 「商品情報」より前にある最後の「X,XXX円」
    if (!found) {
      const infoIdx = lines.findIndex(l => /商品情報/.test(l));
      const searchEnd = infoIdx > 0 ? infoIdx : lines.length;
      for (let i = 0; i < searchEnd; i++) {
        const m = lines[i].match(/(\d[\d,]+)\s*円/);
        if (m) {
          result.price = m[1].replace(/,/g, '');
          found = true;
        }
      }
      if (found) console.log('[DEBUG] 方法4(商品情報前)で価格取得:', result.price);
    }

    // 方法5: 数字のみの行で3桁以上（価格っぽいもの）
    if (!found) {
      for (const line of lines) {
        const m = line.match(/^(\d[\d,]{2,})$/);
        if (m) {
          const num = parseInt(m[1].replace(/,/g, ''));
          if (num >= 100 && num <= 999999) {
            result.price = String(num);
            found = true;
            console.log('[DEBUG] 方法5(数字のみ行)で価格取得:', result.price);
            break;
          }
        }
      }
    }

    console.log('[DEBUG] 最終価格:', result.price || '(見つからず)');
  }

  // ─── 商品名抽出 ───
  // 方法1: SKU番号（A805B等）を含む行を商品名とする（最も確実）
  if (result.parentSku) {
    for (const line of lines) {
      if (line.includes(result.parentSku) && line.length > 20) {
        result.itemName = 'play-zone ' + line;
        break;
      }
    }
  }
  // 方法2: SKUが見つからなかった場合、商品情報セクションの前の長い行を使う
  if (!result.itemName) {
    const infoIdx = lines.findIndex(l => /^商品情報/.test(l));
    const priceIdx = lines.findIndex(l => /^\d[\d,]+\s*円/.test(l));
    const searchEnd = infoIdx > 0 ? infoIdx : (priceIdx > 0 ? priceIdx : lines.length);

    for (let i = 0; i < searchEnd; i++) {
      const line = lines[i];
      if (line.length < 20) continue;
      if (/^(カテゴリ|新商品|YEARS|ANNIVERSARY|楽天|Rakuten|お気に入り|ログイン|検索|ジャンル|ヘルプ)/i.test(line)) continue;
      if (/^\d[\d,]*\s*円/.test(line)) continue;
      if (/^商品番号/.test(line)) continue;
      if (/^(最強翌日|39ショップ|\d+ポイント)/.test(line)) continue;
      result.itemName = 'play-zone ' + line;
      break;
    }
  }
  // 商品名からNewを削除（大文字小文字問わず）
  if (result.itemName) {
    result.itemName = result.itemName.replace(/\s*new\s*/gi, ' ').replace(/\s+/g, ' ').trim();
  }

  // ─── キャッチコピー抽出 ───
  // キャッチコピーの特徴: スペース区切りのキーワードが多数並ぶ（3個以上）
  // 商品名の直前にある。普通の文章やUI要素はスペースが少ない。
  let catchCopy = '';
  if (result.parentSku) {
    // 商品名の最後の出現位置を探す（ページ下部の本体部分）
    // SKU全体で見つからない場合は数字部分だけで検索（A114V → A114）
    const skuVariants = [result.parentSku];
    const baseNum = result.parentSku.match(/^([A-Za-z]*\d+)/);
    if (baseNum && baseNum[1] !== result.parentSku) {
      skuVariants.push(baseNum[1]);
    }
    let nameIdx = -1;
    for (const skuTry of skuVariants) {
      for (let i = lines.length - 1; i >= 0; i--) {
        if (lines[i].includes(skuTry) && lines[i].length > 20) {
          nameIdx = i;
          break;
        }
      }
      if (nameIdx !== -1) break;
    }
    console.log('[DEBUG] 商品名行index(最後):', nameIdx);
    if (nameIdx > 0) {
      // 商品名から上に遡り、スペース区切りキーワードが多い行を探す
      for (let i = nameIdx - 1; i >= Math.max(0, nameIdx - 30); i--) {
        const l = lines[i];
        if (l.includes(result.parentSku)) break;
        // スペースで分割して、2文字以上の単語が4個以上あるか？
        const words = l.split(/[\s　]+/).filter(w => w.length >= 2);
        console.log(`[DEBUG] Line ${i}: words=${words.length} "${l.slice(0, 60)}"`);
        if (words.length >= 4 && l.length >= 20) {
          catchCopy = l;
          console.log('[DEBUG] キャッチコピー発見:', l.slice(0, 100));
          break;
        }
      }
    }
  }
  console.log('[DEBUG] 最終キャッチコピー:', catchCopy || '(見つからず)');

  // ─── 商品情報テーブルから各項目を抽出 ───
  // 楽天のテーブルをコピーすると「ラベル\t値」or「ラベル 値」になる
  const tableLabels = {
    'セット内容': 'setContents',
    'カラー': 'colorRaw',
    '色': 'colorRaw',
    'サイズ': 'sizeRaw',
    '素材': 'materialRaw',
  };

  const tableData = {};
  for (const line of lines) {
    for (const [label, key] of Object.entries(tableLabels)) {
      // "セット内容\tワンピースドレス、Tバック" or "セット内容 ワンピースドレス、Tバック"
      const re = new RegExp('^' + label + '[\\s\\t:：]+(.+)', '');
      const m = line.match(re);
      if (m && !tableData[key]) {
        tableData[key] = m[1].trim();
      }
    }
  }

  // セット内容
  result.setContents = tableData.setContents || '';

  // カラー
  if (tableData.colorRaw) {
    const colorText = tableData.colorRaw;
    const colors = colorText.split(/[、,\/]+/).map(c => c.trim())
      .filter(c => c && c !== '未選択' && c !== '選択してください');
    result.colors = colors;
  }

  // カラーが見つからない場合、楽天の「商品詳細を選択」セクションから取得
  // パターン: 「カラー：未選択」or「カラー 未選択」の後に「ホワイト\n1,980円\nピンク\n1,980円」
  if (result.colors.length === 0) {
    // 「カラー」を含む行を探す（商品詳細選択セクション内）
    let colorSelIdx = -1;
    for (let i = 0; i < lines.length; i++) {
      if (/^(カラー|色)\s*[:：]?\s*(未選択|選択)/.test(lines[i]) ||
          (lines[i] === 'カラー' || lines[i] === '色')) {
        colorSelIdx = i;
        console.log('[DEBUG] カラー選択セクション発見 Line ' + i + ': ' + lines[i]);
        break;
      }
    }
    if (colorSelIdx >= 0) {
      for (let i = colorSelIdx + 1; i < Math.min(colorSelIdx + 30, lines.length); i++) {
        const l = lines[i];
        // 価格行はスキップ
        if (/^\d[\d,]*\s*円?$/.test(l) || /^円$/.test(l)) continue;
        // 「未選択」「選択してください」はスキップ
        if (/未選択|選択して/.test(l)) continue;
        // 空白・短すぎはスキップ
        if (l.length < 1) continue;
        // セクション終了判定
        if (/^(数量|注文|了承|◆|商品情報|セット内容|素材|タイプ|商品詳細|商品説明|その他)/.test(l)) break;
        // 「サイズ」が出たら色セクション終了（サイズ選択セクション開始）
        if (/^サイズ\s*[:：]?\s*(未選択|選択)?/.test(l)) break;
        // 短い日本語テキスト（1〜20文字、数字でない）は色名候補
        if (l.length >= 1 && l.length <= 20 && !/^\d/.test(l) && !/円/.test(l) && !/了承/.test(l)) {
          result.colors.push(l.trim());
          console.log('[DEBUG] カラー選択肢取得:', l.trim());
        }
      }
    }
  }

  // サイズ
  if (tableData.sizeRaw) {
    const sizeText = tableData.sizeRaw;
    if (sizeText.match(/フリー/)) {
      result.sizes = ['Free Size'];
    } else {
      result.sizes = sizeText.split(/[、,\/\s]+/).map(s => s.trim()).filter(Boolean);
    }
  }
  if (result.sizes.length === 0) result.sizes = ['Free Size'];

  // 素材
  if (tableData.materialRaw) {
    const matText = tableData.materialRaw;
    // "ポリエステル/ポリウレタン" → material=ポリエステル, fabricType=ポリエステル、ポリウレタン
    const parts = matText.split(/[\/]+/).map(s => s.trim()).filter(Boolean);
    result.material = parts[0] || matText;
    result.fabricType = parts.join('、');
  }

  // ─── 商品説明を完成形フォーマットで構築 ───
  // 商品情報 → セット内容/カラー/サイズ/素材 → 商品説明(■) → 注意書き → 梱包
  const descParts = [];
  descParts.push('商品情報');
  if (result.setContents) descParts.push('セット内容：' + result.setContents);
  if (tableData.colorRaw) descParts.push('カラー：' + tableData.colorRaw);
  if (tableData.sizeRaw) descParts.push('サイズ：' + tableData.sizeRaw);
  if (tableData.materialRaw) descParts.push('素材：' + tableData.materialRaw);
  descParts.push('');

  // ■商品説明
  let bulletLines = [];
  let inDesc = false;
  for (const line of lines) {
    if (line.match(/^(その他商品説明|商品説明)/)) {
      inDesc = true;
      const rest = line.replace(/^(その他商品説明|商品説明)\s*/, '').trim();
      if (rest) bulletLines.push(rest);
      continue;
    }
    if (inDesc) {
      if (line.match(/^(注意書き|注意事項|梱包について|【プライバシー)/)) break;
      bulletLines.push(line);
    }
  }

  if (bulletLines.length > 0) {
    descParts.push('商品説明');
    bulletLines.forEach(l => descParts.push(l));
    descParts.push('');
  }

  // 注意書き
  let inNotice = false;
  const noticeLines = [];
  for (const line of lines) {
    if (line.match(/^(注意書き|注意事項)/)) {
      inNotice = true;
      const rest = line.replace(/^(注意書き|注意事項)\s*/, '').trim();
      if (rest) noticeLines.push(rest);
      continue;
    }
    if (inNotice) {
      if (line.match(/^(梱包について|【プライバシー)/)) break;
      noticeLines.push(line);
    }
  }
  if (noticeLines.length > 0) {
    descParts.push('注意事項');
    noticeLines.forEach(l => descParts.push(l));
    descParts.push('');
  }

  // 梱包について
  descParts.push('梱包について');
  descParts.push('【プライバシー配慮★安心安全】');
  descParts.push('中身がわからないように、品名は「衣類」と記載し、発送元は「個人名」で発送します。');

  result.description = descParts.join('<br>\n');

  // ■で始まる仕様を箇条書き(bullet_point)に
  const allBulletText = bulletLines.join(' ');
  const bulletMatches = allBulletText.match(/■[^■]+/g);
  if (bulletMatches) {
    result.bullets = bulletMatches.slice(0, 5).map(b => b.replace(/^■\s*/, '').trim());
  }

  // 箇条書きが足りなければセット内容・素材・サイズで補完
  if (result.bullets.length === 0) {
    if (result.setContents) result.bullets.push('セット内容: ' + result.setContents);
    if (result.fabricType) result.bullets.push('素材: ' + result.fabricType);
    if (result.sizes.length > 0) result.bullets.push('サイズ: ' + result.sizes.join(', '));
  }

  // ─── 検索キーワード生成 ───
  // キャッチコピーのみ（商品名は入れない）
  console.log('[DEBUG] キャッチコピー:', catchCopy || '(見つからず)');
  if (catchCopy) {
    const words = catchCopy.split(/[\s　]+/).map(w => w.trim()).filter(w => w.length >= 2);
    result.keywords = words.slice(0, 20).join(' ');
  }
  console.log('[DEBUG] 最終キーワード:', result.keywords || '(空)');

  // ─── 参考価格 = 販売価格×2＋20円 ───
  if (result.price) {
    result.listPrice = String(parseInt(result.price) * 2 + 20);
  }

  // ─── バリエーションテーマ自動判定 ───
  const hasMultiColor = result.colors.length > 1;
  const hasMultiSize = result.sizes.length > 1 || (result.sizes.length === 1 && result.sizes[0] !== 'Free Size');
  if (hasMultiColor && hasMultiSize) {
    result.varTheme = 'サイズ_名前/色_名前';
  } else if (hasMultiColor) {
    result.varTheme = '色';
  } else if (hasMultiSize) {
    result.varTheme = 'サイズ_名前';
  } else {
    result.varTheme = ''; // なし（単品）
  }
  console.log('[DEBUG] バリエーションテーマ自動判定:', result.varTheme || 'なし（単品）');

  // ─── 商品タイプ自動判定 ───
  result.productType = 'COSTUME_OUTFIT'; // デフォルト
  if (result.itemName && /(レオタード|体操服|ボディスーツ|ボディースーツ|ジャンプスーツ|水着|競泳)/.test(result.itemName)) {
    result.productType = 'LEOTARD';
  }
  console.log('[DEBUG] 商品タイプ自動判定:', result.productType);

  return result;
}

/* ==========================================================
   フォームに反映
   ========================================================== */
function applyParsedToForm(parsed) {
  if (!parsed) return;

  document.getElementById('f-parentSku').value = parsed.parentSku || '';
  document.getElementById('f-itemName').value = parsed.itemName || '';
  document.getElementById('f-price').value = parsed.price || '';
  document.getElementById('f-listPrice').value = parsed.listPrice || parsed.price || '';
  document.getElementById('f-material').value = parsed.material || '';
  document.getElementById('f-fabricType').value = parsed.fabricType || '';
  document.getElementById('f-description').value = parsed.description || '';
  document.getElementById('f-keywords').value = parsed.keywords || '';
  document.getElementById('f-variationTheme').value = parsed.varTheme ?? '';
  document.getElementById('f-productType').value = parsed.productType || 'COSTUME_OUTFIT';

  for (let i = 1; i <= 5; i++) {
    document.getElementById('f-bullet' + i).value = parsed.bullets[i-1] || '';
  }

  // 色バリエーション
  const colorList = document.getElementById('color-list');
  colorList.innerHTML = '';
  const colors = parsed.colors.length > 0 ? parsed.colors : ['ブラック'];
  colors.forEach(c => addColorRow(c));

  // サイズバリエーション
  const sizeList = document.getElementById('size-list');
  sizeList.innerHTML = '';
  parsed.sizes.forEach(s => addSizeRow(s));

  document.getElementById('form-section').style.display = '';
  document.getElementById('download-section').style.display = '';
}

/* ==========================================================
   バリエーション行の追加
   ========================================================== */
function addColorRow(colorName) {
  const list = document.getElementById('color-list');
  const row = document.createElement('div');
  row.className = 'variant-row';
  row.innerHTML = `
    <input type="text" class="color-name" value="${escHtml(colorName || '')}" placeholder="色名">
    <input type="text" class="color-suffix" value="${escHtml(colorToSuffix(colorName))}" placeholder="SKU末尾 (例: B)">
    <div class="img-inputs">
      <input type="text" class="color-img" placeholder="メイン画像URL">
    </div>
    <button class="btn-remove" onclick="this.closest('.variant-row').remove()">削除</button>
  `;
  list.appendChild(row);
}

function addSizeRow(sizeName) {
  const list = document.getElementById('size-list');
  const row = document.createElement('div');
  row.className = 'variant-row';
  row.innerHTML = `
    <input type="text" class="size-name" value="${escHtml(sizeName || '')}" placeholder="サイズ名">
    <button class="btn-remove" onclick="this.closest('.variant-row').remove()">削除</button>
  `;
  list.appendChild(row);
}

function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function colorToSuffix(color) {
  const map = {
    'ブラック':'B','黒':'B','black':'B',
    'ホワイト':'W','白':'W','white':'W',
    'レッド':'R','赤':'R','red':'R',
    'ブルー':'BL','青':'BL','blue':'BL',
    'ピンク':'P','pink':'P',
    'パープル':'PP','紫':'PP','purple':'PP',
    'グリーン':'G','緑':'G','green':'G',
    'イエロー':'Y','黄':'Y','yellow':'Y',
    'オレンジ':'O','orange':'O',
    'ネイビー':'N','navy':'N',
    'グレー':'GR','灰':'GR','gray':'GR','grey':'GR',
    'ベージュ':'BE','beige':'BE',
    'ワインレッド':'WR','ワイン':'WR',
    'ローズ':'RS','rose':'RS',
  };
  const key = (color || '').trim().toLowerCase();
  for (const [k, v] of Object.entries(map)) {
    if (k.toLowerCase() === key) return v;
  }
  return (color || 'X').charAt(0).toUpperCase();
}

/* ==========================================================
   フォームからデータ収集
   ========================================================== */
function generateSheetData() {
  const data = {
    parentSku:  document.getElementById('f-parentSku').value.trim(),
    itemName:   document.getElementById('f-itemName').value.trim(),
    price:      document.getElementById('f-price').value.trim(),
    listPrice:  document.getElementById('f-listPrice').value.trim(),
    material:   document.getElementById('f-material').value.trim(),
    fabricType: document.getElementById('f-fabricType').value.trim(),
    description:document.getElementById('f-description').value.trim(),
    keywords:   document.getElementById('f-keywords').value.trim(),
    varTheme:   document.getElementById('f-variationTheme').value,
    productType: document.getElementById('f-productType').value,
    bullets: [],
    colors: [],
    sizes: [],
  };

  for (let i = 1; i <= 5; i++) {
    data.bullets.push(document.getElementById('f-bullet' + i).value.trim());
  }

  document.querySelectorAll('#color-list .variant-row').forEach(row => {
    const name = row.querySelector('.color-name').value.trim();
    const suffix = row.querySelector('.color-suffix').value.trim();
    const img = row.querySelector('.color-img').value.trim();
    if (name) data.colors.push({ name, suffix, img });
  });

  document.querySelectorAll('#size-list .variant-row').forEach(row => {
    const name = row.querySelector('.size-name').value.trim();
    if (name) data.sizes.push(name);
  });

  if (!data.parentSku) {
    showToast('親SKUを入力してください', 'error');
    return null;
  }
  if (!data.itemName) {
    showToast('商品名を入力してください', 'error');
    return null;
  }

  return data;
}

/* ==========================================================
   行データ構築
   ========================================================== */
function buildRow(sku, data, opts) {
  // opts: { isParent, isSingle, color, colorSuffix, size, imageUrl }
  const row = new Array(300).fill('');

  // === 識別子 ===
  row[COL.SKU - 1]          = sku;
  row[COL.productType - 1]  = data.productType || FIXED.productType;
  row[COL.recordAction - 1] = FIXED.recordAction;

  if (opts.isSingle) {
    row[COL.parentage - 1]  = '';
    row[COL.parentSku - 1]  = '';
    row[COL.varTheme - 1]   = '';
  } else {
    row[COL.parentage - 1]  = opts.isParent ? '親' : '子供';
    row[COL.parentSku - 1]  = opts.isParent ? '' : data.parentSku;
    row[COL.varTheme - 1]   = data.varTheme;
  }

  // === 基本情報 ===
  row[COL.itemName - 1]     = data.itemName;
  row[COL.brand - 1]        = FIXED.brand;
  row[COL.idType - 1]       = 'GTIN免除';
  row[COL.browseNode1 - 1]  = FIXED.browseNode;
  row[COL.packageUnit - 1]  = FIXED.packageUnit;
  row[COL.targetAudience1 - 1] = FIXED.targetAud;
  row[COL.modelNumber - 1]  = sku;                 // Y: 品番
  row[COL.modelName - 1]    = FIXED.modelName;      // Z: イベント衣装
  row[COL.manufacturer - 1] = FIXED.manufacturer;   // AA: play-zone
  row[COL.partNumber - 1]   = sku;                  // BZ: メーカー型番
  row[COL.setName - 1]      = sku;                  // EK: セット名

  // === 画像 ===
  if (opts.imageUrl) {
    row[COL.mainImage - 1] = opts.imageUrl;
  }

  // === 説明・仕様・キーワード ===
  row[COL.description - 1]  = data.description;
  for (let i = 0; i < 5; i++) {
    row[COL.bullet1 + i - 1] = data.bullets[i] || '';
  }
  row[COL.keywords - 1]     = data.keywords;

  // === 属性 ===
  row[COL.lifestyle - 1]    = FIXED.lifestyle;
  row[COL.style - 1]        = FIXED.style;
  row[COL.department - 1]   = FIXED.department;
  row[COL.gender - 1]       = FIXED.gender;
  row[COL.ageRange - 1]     = FIXED.ageRange;
  row[COL.material1 - 1]    = data.material;
  row[COL.fabricType - 1]   = data.fabricType;
  row[COL.materialComp - 1] = data.fabricType;       // CD: 材料の組成
  row[COL.careInst - 1]     = FIXED.careInst;        // CE: お手入れ方法
  row[COL.specialSize - 1]  = FIXED.specialSize;
  row[COL.occasion1 - 1]    = FIXED.occasion1;       // BU: ハロウィン
  row[COL.occasion2 - 1]    = FIXED.occasion2;       // BV: カーニバル

  // === 色・サイズ ===
  row[COL.colorMap - 1]     = opts.color || '';
  row[COL.color - 1]        = opts.color || '';
  row[COL.size - 1]         = opts.size || '';

  // === 輸出入・限定 ===
  row[COL.importType - 1]   = FIXED.importType;
  row[COL.isExclusive - 1]  = FIXED.isExclusive;

  // === オファー情報 ===
  row[COL.condition - 1]    = FIXED.conditionVal;
  row[COL.listPrice - 1]    = data.listPrice || data.price;
  row[COL.fulfillment - 1]  = FIXED.fulfillment;
  row[COL.quantity - 1]     = '10';
  row[COL.price - 1]        = data.price;
  row[COL.shipping - 1]     = FIXED.shipping;

  // === 安全・原産国 ===
  row[COL.countryOfOrigin - 1] = FIXED.countryOfOrigin; // HJ: 中国
  row[COL.battery - 1]      = FIXED.battery;            // HK: いいえ
  row[COL.hazmat - 1]       = FIXED.hazmat;             // IA: 該当なし

  return row;
}

function buildSheetRows(data) {
  const rows = [];

  // バリエーションなし（単品）→ 親+子1行のみ（色バリエーション無視）
  const isSingle = !data.varTheme;
  if (isSingle) {
    data.varTheme = '色';
    console.log('[DEBUG] 単品モード: 親+子1行出力（色バリエーション無視）');
  }

  // 親行
  const parentRow = buildRow(data.parentSku, data, {
    isParent: true,
    color: data.colors[0]?.name || '',
    size: data.sizes[0] || 'Free Size',
  });
  rows.push(parentRow);

  if (isSingle) {
    // 単品: 子1行のみ（SKU=親SKU、色サフィックスなし）
    const childRow = buildRow(data.parentSku, data, {
      isParent: false,
      color: data.colors[0]?.name || '',
      size: data.sizes[0] || 'Free Size',
      imageUrl: data.colors[0]?.img || '',
    });
    rows.push(childRow);
  } else {
    // バリエーションあり: 色 × サイズの組み合わせ
    const colors = data.colors.length > 0 ? data.colors : [{ name: '', suffix: '', img: '' }];
    const sizes = data.sizes.length > 0 ? data.sizes : ['Free Size'];

    for (const color of colors) {
      for (const size of sizes) {
        let childSku = data.parentSku + (color.suffix || '');
        if (sizes.length > 1) {
          const sizeCode = size.replace(/\s+/g, '').toUpperCase().slice(0, 3);
          childSku += sizeCode;
        }

        const childRow = buildRow(childSku, data, {
          isParent: false,
          color: color.name,
          colorSuffix: color.suffix,
          size: size,
          imageUrl: color.img || '',
        });
        rows.push(childRow);
      }
    }
  }

  console.log('[DEBUG] ' + (isSingle ? '単品' : 'バリエーション') + 'モード: ' + rows.length + '行出力 (親1+子' + (rows.length - 1) + ')');
  return rows;
}

/* ==========================================================
   テンプレートベースのXLSMダウンロード
   ========================================================== */
let _templateArrayBuffer = null;
const TEMPLATE_DB_NAME = 'PlayZoneTemplateDB';
const TEMPLATE_STORE = 'templates';
const TEMPLATE_KEY = 'amazon_template';

// IndexedDBにテンプレートを保存（ページを閉じても残る）
function saveTemplateToIDB(arrayBuffer) {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(TEMPLATE_DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(TEMPLATE_STORE);
    req.onsuccess = () => {
      const tx = req.result.transaction(TEMPLATE_STORE, 'readwrite');
      tx.objectStore(TEMPLATE_STORE).put(arrayBuffer, TEMPLATE_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    };
    req.onerror = () => reject(req.error);
  });
}

// IndexedDBからテンプレートを読み込み
function loadTemplateFromIDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(TEMPLATE_DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(TEMPLATE_STORE);
    req.onsuccess = () => {
      const tx = req.result.transaction(TEMPLATE_STORE, 'readonly');
      const getReq = tx.objectStore(TEMPLATE_STORE).get(TEMPLATE_KEY);
      getReq.onsuccess = () => resolve(getReq.result || null);
      getReq.onerror = () => reject(getReq.error);
    };
    req.onerror = () => reject(req.error);
  });
}

// 列番号(0-based) → Excelの列文字 (0=A, 1=B, ..., 25=Z, 26=AA, ...)
function colToLetter(c) {
  let s = '';
  let n = c;
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

// XMLエスケープ
function xmlEsc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// description列のクリーン（<br>タグは保持、それ以外のHTMLタグは除去）
function cleanDescriptionCell(val) {
  return String(val ?? '')
    .replace(/[\r\n]+/g, '')    // 改行は除去（<br>で改行するので）
    .replace(/\s+/g, ' ')
    .trim();
}

// 1行分のXML文字列を生成（inlineStr形式、実績ファイルと同じ）
function buildRowXml(rowNum, rowData) {
  let cells = '';
  const descColIdx = COL.description - 1; // description列のインデックス(0-based)
  rowData.forEach((val, ci) => {
    // description列は<br>タグを保持する特別処理
    const cellVal = (ci === descColIdx)
      ? cleanDescriptionCell(String(val ?? ''))
      : cleanDataCell(String(val ?? ''));
    if (!cellVal) return;
    const col = colToLetter(ci);
    const ref = col + rowNum;
    cells += `<c r="${ref}" t="inlineStr" s="62"><is><t>${xmlEsc(cellVal)}</t></is></c>`;
  });
  return `<row r="${rowNum}" spans="1:297">${cells}</row>`;
}

async function downloadXlsx(data) {
  // クリーンテンプレート（成功実績ファイルベース、sharedStringsなし）を読み込む
  let buf = _templateArrayBuffer;

  if (!buf) {
    buf = await loadTemplateFromIDB();
    if (buf) _templateArrayBuffer = buf;
  }

  if (!buf) {
    showToast('先にテンプレート(.xlsm)を選択してください', 'error');
    return;
  }

  // JSZipでテンプレートを開く
  const zip = await JSZip.loadAsync(buf);
  const sheetPath = 'xl/worksheets/sheet5.xml';
  let xml = await zip.file(sheetPath).async('string');

  // ヘッダー行(Row1〜5) + テンプレート行(Row6=ABC123, Row7)を抽出
  // ※ Row6-7はAmazonがスキップする固定行。データはRow8から書く。
  const keepRows = [];
  for (let r = 1; r <= 7; r++) {
    const re = new RegExp('<row r="' + r + '"[^>]*(?:\\/>|>[\\s\\S]*?<\\/row>)');
    const m = xml.match(re);
    if (m) keepRows.push(m[0]);
  }
  console.log('[DEBUG] 保持行: ' + keepRows.length + '行 (Row1-7)');

  // データ行XMLを構築（Row8から開始）
  const DATA_START_ROW = 8;
  const allRows = buildSheetRows(data);
  const dataRowsXml = allRows.map((rowData, ri) => buildRowXml(DATA_START_ROW + ri, rowData));
  console.log('[DEBUG] 生成行数: ' + allRows.length + ' (親1+子' + (allRows.length - 1) + ')');

  // 末尾の空マーカー行
  const lastDataRow = DATA_START_ROW + allRows.length - 1;
  const emptyRowNum = lastDataRow + 1;

  // sheetData全体を再構築
  const newSheetData = '<sheetData>'
    + keepRows.join('')
    + dataRowsXml.join('')
    + '<row r="' + emptyRowNum + '" ht="12.75" />'
    + '</sheetData>';

  const sdStart = xml.indexOf('<sheetData>');
  const sdEnd = xml.indexOf('</sheetData>');
  xml = xml.slice(0, sdStart) + newSheetData + xml.slice(sdEnd + '</sheetData>'.length);

  // dimensionを更新
  xml = xml.replace(/<dimension ref="[^"]*"\/>/,
    '<dimension ref="A1:KK' + emptyRowNum + '"/>');

  // sheet5.xmlを書き戻す
  zip.file(sheetPath, xml);
  console.log('[DEBUG] sheetData再構築完了: 保持7行 + データ' + allRows.length + '行 (Row' + DATA_START_ROW + '〜' + lastDataRow + ')');

  // xlsmとしてダウンロード
  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12'
  });

  const fileName = `Amazon_${data.parentSku}_${formatDate()}.xlsm`;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(a.href);
  showToast('ダウンロードしました: ' + fileName, 'success');
}

/* ==========================================================
   イベントバインド
   ========================================================== */
document.addEventListener('DOMContentLoaded', () => {
  // テンプレートファイル選択 → IndexedDBに永続保存
  document.getElementById('template-file').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (file) {
      const buf = await file.arrayBuffer();
      _templateArrayBuffer = buf;
      await saveTemplateToIDB(buf);
      document.getElementById('template-status').textContent = '保存済み（次回から自動読み込み）';
      document.getElementById('template-status').style.color = '#10b981';
      document.getElementById('template-section').style.background = '#f0fdf4';
      showToast('テンプレート保存OK。次回から自動で読み込みます。');
    }
  });

  // ページ読み込み時にIndexedDBからテンプレートを自動復元
  loadTemplateFromIDB().then(buf => {
    if (buf) {
      _templateArrayBuffer = buf;
      document.getElementById('template-status').textContent = '保存済み（自動読み込み完了）';
      document.getElementById('template-status').style.color = '#10b981';
      document.getElementById('template-section').style.background = '#f0fdf4';
    }
  }).catch(() => {});

  // 解析ボタン
  document.getElementById('btn-parse').addEventListener('click', () => {
    const text = document.getElementById('rakuten-input').value;
    if (!text.trim()) {
      showToast('テキストを貼り付けてください', 'error');
      return;
    }
    const parsed = parseRakutenText(text);
    if (parsed) {
      applyParsedToForm(parsed);
      showToast('解析完了');
    } else {
      showToast('解析できませんでした', 'error');
    }
  });

  // クリアボタン
  document.getElementById('btn-clear').addEventListener('click', () => {
    document.getElementById('rakuten-input').value = '';
    document.getElementById('form-section').style.display = 'none';
    document.getElementById('download-section').style.display = 'none';
    showToast('クリアしました');
  });

  // 色追加
  document.getElementById('btn-add-color').addEventListener('click', () => addColorRow(''));

  // サイズ追加
  document.getElementById('btn-add-size').addEventListener('click', () => addSizeRow(''));

  // ダウンロードボタン
  document.getElementById('btn-download').addEventListener('click', async () => {
    const data = generateSheetData();
    if (!data) return;
    try {
      await downloadXlsx(data);
    } catch (e) {
      console.error(e);
      showToast('ダウンロードエラー: ' + e.message, 'error');
    }
  });
});

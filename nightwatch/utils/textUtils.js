function parseActionsInOrder(description) {
  console.log(`\n🔍 Phân tích mô tả: "${description}"`);
  console.log(description);
  const patterns = [
    {
      type: "find_title_all",
      regex: /Tìm thấy "(.*?)" theo Tiêu đề và Tất cả cụm từ/g,
      extract: (m) => ({
        action: "find_title_all",
        text: m[1],
      }),
    },
    {
      type: "find_title_exact",
      regex: /Tìm thấy "(.*?)" theo Tiêu đề và Chính xác cụm từ/g,
      extract: (m) => ({
        action: "find_title_exact",
        text: m[1],
      }),
    },
    {
      type: "find_title_any",
      regex: /Tìm thấy "(.*?)" theo Tiêu đề và Ít nhất một từ/g,
      extract: (m) => ({
        action: "find_title_any",
        text: m[1],
      }),
    },
    {
      type: "find_content_all",
      regex: /Tìm thấy "(.*?)" theo Nội dung và Tất cả cụm từ/g,
      extract: (m) => ({
        action: "find_content_all",
        text: m[1],
      }),
    },
    {
      type: "find_content_exact",
      regex: /Tìm thấy "(.*?)" theo Nội dung và Chính xác cụm từ/g,
      extract: (m) => ({
        action: "find_content_exact",
        text: m[1],
      }),
    },
    {
      type: "find_content_any",
      regex: /Tìm thấy "(.*?)" theo Nội dung và Ít nhất một từ/g,
      extract: (m) => ({
        action: "find_content_any",
        text: m[1],
      }),
    },
    {
      type: "find_keyword_all",
      regex: /Tìm thấy "(.*?)" theo Từ khóa và Tất cả cụm từ/g,
      extract: (m) => ({
        action: "find_keyword_all",
        text: m[1],
      }),
    },
    {
      type: "find_keyword_exact",
      regex: /Tìm thấy "(.*?)" theo Từ khóa và Chính xác cụm từ/g,
      extract: (m) => ({
        action: "find_keyword_exact",
        text: m[1],
      }),
    },
    {
      type: "find_keyword_any",
      regex: /Tìm thấy "(.*?)" theo Từ khóa và Ít nhất một từ/g,
      extract: (m) => ({
        action: "find_keyword_any",
        text: m[1],
      }),
    },
    {
      type: "find_all_all",
      regex: /Tìm thấy "(.*?)" theo Tất cả và Tất cả cụm từ/g,
      extract: (m) => ({
        action: "find_all_all",
        text: m[1],
      }),
    },
    {
      type: "find_all_exact",
      regex: /Tìm thấy "(.*?)" theo Tất cả và Chính xác cụm từ/g,
      extract: (m) => ({
        action: "find_all_exact",
        text: m[1],
      }),
    },
    {
      type: "find_all_any",
      regex: /Tìm thấy "(.*?)" theo Tất cả và Ít nhất một từ/g,
      extract: (m) => ({
        action: "find_all_any",
        text: m[1],
      }),
    },
    {
      type: "click_by_radio",
      regex: /Chọn radio "(.*?)"/g,
      extract: (m) => ({
        action: "click_by_radio",
        text: m[1],
      }),
    },
    {
      type: "upload_file",
      regex: /Tải tệp "(.*?)" lên tại ô có id "(.*?)"/g,
      extract: (m) => ({
        action: "upload_file",
        filename: m[1],
        elementId: m[2],
      }),
    },
    {
      type: "navigate",
      regex: /Truy cập trang "(.*?)"/g,
      extract: (m) => ({
        action: "navigate",
        url: m[1],
      }),
    },
    {
      type: "alert_check_and_accept",
      regex: /Xuất hiện alert với nội dung là "(.*?)", sau đó nhấn OK/g,
      extract: (m) => ({
        action: "alert_check_and_accept",
        expectedText: m[1],
      }),
    },
    {
      type: "click_by_id",
      regex: /Chọn thẻ có id là "?([a-zA-Z0-9_-]+)"?/g,
      extract: (m) => ({
        action: "click_by_id",
        id: m[1],
      }),
    },
    {
      type: "or",
      regex: /\s+hoặc\s+/g,
      extract: (m) => ({ action: "or" }),
    },
    {
      type: "or_statement",
      regex: /Di chuột đến "(.*?)"/g,
      extract: (m) => ({ action: "hover", targetText: m[1] }),
    },
    {
      type: "hover",
      regex: /Di chuột đến "(.*?)"/g,
      extract: (m) => ({ action: "hover", targetText: m[1] }),
    },
    {
      type: "click",
      regex: /Bấm vào "(.*?)"(?! từ)/g,
      extract: (m) => ({ action: "click", targetText: m[1] }),
    },
    {
      type: "dropdown_click",
      regex: /Chọn "(.*?)" từ "(.*?)" được xổ xuống/g,
      extract: (m) => ({
        action: "dropdown_click",
        childText: m[1],
        parentText: m[2],
      }),
    },
    {
      type: "scroll",
      regex: /Cuộn chuột (xuống|lên)/g,
      extract: (m) => ({
        action: "scroll",
        direction: m[1] === "xuống" ? "down" : "up",
      }),
    },
    {
      type: "click_input",
      regex: /Bấm vào ô "(.*?)"/g,
      extract: (m) => ({ action: "click_input_by_label", label: m[1] }),
    },
    {
      type: "type",
      regex: /Gõ "(.*?)"/g,
      extract: (m) => ({ action: "type", value: m[1] }),
    },
    {
      type: "press_key",
      regex: /Nhấn nút "(Enter|Tab|Esc)"/gi,
      extract: (m) => ({ action: "press_key", key: m[1].toUpperCase() }),
    },
    {
      type: "drag_drop",
      regex: /Kéo "(.*?)" và thả vào "(.*?)"/g,
      extract: (m) => ({
        action: "drag_drop",
        sourceText: m[1],
        targetText: m[2],
      }),
    },
    {
      type: "select_dropdown_by_id",
      regex: /Chọn "(.*?)" từ danh sách có id là "(.*?)"/g,
      extract: (m) => ({
        action: "select_dropdown_by_id",
        value: m[1],
        id: m[2],
      }),
    },
    {
      type: "wait_time",
      regex: /Chờ "?(\d+)"? giây/g,
      extract: (m) => ({ action: "wait", seconds: parseInt(m[1], 10) }),
    },
    {
      type: "check_count",
      regex: /Kiểm tra số lượng "(.*?)" là (\d+)/g,
      extract: (m) => ({
        action: "check_count",
        text: m[1],
        expectedCount: parseInt(m[2], 10),
      }),
    },
    {
      type: "check_visible",
      regex: /Kiểm tra thấy "(.*?)"/g,
      extract: (m) => ({ action: "check_visible", text: m[1] }),
    },
  ];

  const results = [];

  for (const pattern of patterns) {
    let match;
    while ((match = pattern.regex.exec(description)) !== null) {
      results.push({
        index: match.index,
        matchedText: match[0],
        ...pattern.extract(match),
      });
    }
  }

  // Sắp xếp theo thứ tự xuất hiện
  results.sort((a, b) => a.index - b.index);

  console.log(`\n🔍 Phân tích mô tả: `);
  console.log(results.map(({ index, ...rest }) => rest));
  // Loại bỏ trường 'index'
  return results.map(({ index, ...rest }) => rest);
}

// Hàm hỗ trợ: Tạo XPath so sánh text có dấu (tiếng Việt)
function createTextMatchXpath(inputText) {
  const upperChars =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZĐÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ";
  const lowerChars =
    "abcdefghijklmnopqrstuvwxyzđáàảãạâấầẩẫậăắằẳẵặéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ";
  const lowerText = inputText.toLowerCase();
  return `//*[translate(normalize-space(text()), '${upperChars}', '${lowerChars}') = "${lowerText}"]`;
}

// Hàm hỗ trợ: Tạo XPath so sánh text có dấu (tiếng Việt)
function createTextMatchXpathContain(inputText) {
  const upperChars =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZĐÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ";
  const lowerChars =
    "abcdefghijklmnopqrstuvwxyzđáàảãạâấầẩẫậăắằẳẵặéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ";
  const lowerText = inputText.toLowerCase();
  return `//*[contains(translate(normalize-space(text()), '${upperChars}', '${lowerChars}'), "${lowerText}")]`;
}

// Hàm hỗ trợ: Tạo XPath so sánh text có dấu (tiếng Việt)
function createTextMatchXpathContain(inputText) {
  const upperChars =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZĐÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ";
  const lowerChars =
    "abcdefghijklmnopqrstuvwxyzđáàảãạâấầẩẫậăắằẳẵặéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ";
  const lowerText = inputText.toLowerCase();
  return `//*[contains(translate(normalize-space(text()), '${upperChars}', '${lowerChars}'), "${lowerText}")]`;
}

const accented =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZĐÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴđáàảãạâấầẩẫậăắằẳẵặéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ";
const unaccented =
  "abcdefghijklmnopqrstuvwxyzdaaaaaaaaaaaaaaaaaeeeeeeeeeeeiiiiiooooooooooooooooouuuuuuuuuuuyyyyydaaaaaaaaaaaaaaaaaeeeeeeeeeeeiiiiiooooooooooooooooouuuuuuuuuuuyyyyy";

function getXpathForTittle(text) {
  //(tiêu đề)
  const xpath = `//p[contains(@class, 'title_book') and contains(
    translate(string(.), '${accented}', '${unaccented}'),
      '${removeVietnameseAccents(text)}'
  )]`;
  return xpath;
}

function getXpathForContent(text) {
  //nội dung
  const xpath = `(//div[contains(@class, 'decs_book')]//p)[1][contains(
      translate(normalize-space(string(.)), '${accented}', '${unaccented}'),
      '${removeVietnameseAccents(text)}'
    )]`;
  return xpath;
}
function getXpathForKeyWord(text) {
  // từ khóa
  const xpath = `(//div[contains(@class, 'decs_book')]//p)[3][contains(
      translate(normalize-space(string(.)), '${accented}', '${unaccented}'),
      '${removeVietnameseAccents(text)}'
    )]`;
  return xpath;
}
function getXpathForAll(text) {
  //tất cả
  const xpath = `//*[@id="box_search_book_result"][contains(
      translate(normalize-space(string(.)), '${accented}', '${unaccented}'),
      '${removeVietnameseAccents(text)}'
    )]`;
  return xpath;
}

// Tất cả ký tự có dấu (bao gồm cả chữ hoa và thường)
function removeVietnameseAccents(str) {
  const from =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZĐÁÀẢÃẠÂẤẦẨẪẬĂẮẰẲẴẶÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸỴ" +
    "đáàảãạâấầẩẫậăắằẳẵặéèẻẽẹêếềểễệíìỉĩịóòỏõọôốồổỗộơớờởỡợúùủũụưứừửữựýỳỷỹỵ";
  const to =
    "abcdefghijklmnopqrstuvwxyzdaaaaaaaaaaaaaaaaaeeeeeeeeeeeiiiiiooooooooooooooooouuuuuuuuuuuyyyyy" +
    "daaaaaaaaaaaaaaaaaeeeeeeeeeeeiiiiiooooooooooooooooouuuuuuuuuuuyyyyy";
  return str
    .split("")
    .map((c) => {
      const index = from.indexOf(c);
      return index !== -1 ? to[index] : c;
    })
    .join("");
}

module.exports = {
  parseActionsInOrder,
  createTextMatchXpath,
  createTextMatchXpathContain,
  getXpathForTittle,
  getXpathForContent,
  getXpathForKeyWord,
  getXpathForAll,
};

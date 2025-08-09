/**
 * 配色方案生成器工具
 * 基於色彩理論計算多種配色風格
 */

/**
 * 將 HEX 顏色轉換為 HSL
 * @param {string} hex - HEX 顏色值 (#RRGGBB)
 * @returns {Object} {h, s, l} HSL 值
 */
function hexToHsl(hex) {
	// 移除 # 符號
	hex = hex.replace("#", "");

	// 轉換為 RGB
	const r = parseInt(hex.substr(0, 2), 16) / 255;
	const g = parseInt(hex.substr(2, 2), 16) / 255;
	const b = parseInt(hex.substr(4, 2), 16) / 255;

	const max = Math.max(r, g, b);
	const min = Math.min(r, g, b);
	let h,
		s,
		l = (max + min) / 2;

	if (max === min) {
		h = s = 0; // 灰色
	} else {
		const d = max - min;
		s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

		switch (max) {
			case r:
				h = (g - b) / d + (g < b ? 6 : 0);
				break;
			case g:
				h = (b - r) / d + 2;
				break;
			case b:
				h = (r - g) / d + 4;
				break;
		}
		h /= 6;
	}

	return {
		h: Math.round(h * 360),
		s: Math.round(s * 100),
		l: Math.round(l * 100),
	};
}

/**
 * 將 HSL 顏色轉換為 HEX
 * @param {number} h - 色相 (0-360)
 * @param {number} s - 飽和度 (0-100)
 * @param {number} l - 亮度 (0-100)
 * @returns {string} HEX 顏色值
 */
function hslToHex(h, s, l) {
	h = h % 360;
	s = Math.max(0, Math.min(100, s)) / 100;
	l = Math.max(0, Math.min(100, l)) / 100;

	const c = (1 - Math.abs(2 * l - 1)) * s;
	const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
	const m = l - c / 2;
	let r = 0,
		g = 0,
		b = 0;

	if (0 <= h && h < 60) {
		r = c;
		g = x;
		b = 0;
	} else if (60 <= h && h < 120) {
		r = x;
		g = c;
		b = 0;
	} else if (120 <= h && h < 180) {
		r = 0;
		g = c;
		b = x;
	} else if (180 <= h && h < 240) {
		r = 0;
		g = x;
		b = c;
	} else if (240 <= h && h < 300) {
		r = x;
		g = 0;
		b = c;
	} else if (300 <= h && h < 360) {
		r = c;
		g = 0;
		b = x;
	}

	r = Math.round((r + m) * 255);
	g = Math.round((g + m) * 255);
	b = Math.round((b + m) * 255);

	return (
		"#" +
		[r, g, b]
			.map((x) => {
				const hex = x.toString(16);
				return hex.length === 1 ? "0" + hex : hex;
			})
			.join("")
	);
}

/**
 * 驗證 HEX 顏色格式
 * @param {string} hex - HEX 顏色值
 * @returns {boolean} 是否為有效的 HEX 格式
 */
function isValidHex(hex) {
	const hexRegex = /^#?([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/;
	return hexRegex.test(hex);
}

/**
 * 標準化 HEX 顏色格式
 * @param {string} hex - HEX 顏色值
 * @returns {string} 標準化後的 HEX 顏色值 (#RRGGBB)
 */
function normalizeHex(hex) {
	hex = hex.replace("#", "");

	// 將 3 位 HEX 轉換為 6 位
	if (hex.length === 3) {
		hex = hex
			.split("")
			.map((char) => char + char)
			.join("");
	}

	return "#" + hex.toUpperCase();
}

/**
 * 單色系配色方案（Monochromatic）
 * 基於不同的亮度和飽和度變化
 * @param {string} baseHex - 主色 HEX 值
 * @returns {Array} 5 個顏色的 HEX 陣列
 */
function generateMonochromatic(baseHex) {
	const hsl = hexToHsl(baseHex);
	const colors = [];

	// 主色
	colors.push(baseHex);

	// 更深的變化（降低亮度）
	colors.push(hslToHex(hsl.h, hsl.s, Math.max(10, hsl.l - 30)));

	// 更淺的變化（提高亮度）
	colors.push(hslToHex(hsl.h, hsl.s, Math.min(90, hsl.l + 30)));

	// 降低飽和度
	colors.push(hslToHex(hsl.h, Math.max(10, hsl.s - 40), hsl.l));

	// 提高飽和度
	colors.push(hslToHex(hsl.h, Math.min(100, hsl.s + 20), hsl.l));

	return colors;
}

/**
 * 類似色配色方案（Analogous）
 * 使用色相環上相鄰的顏色
 * @param {string} baseHex - 主色 HEX 值
 * @returns {Array} 5 個顏色的 HEX 陣列
 */
function generateAnalogous(baseHex) {
	const hsl = hexToHsl(baseHex);
	const colors = [];

	// 主色
	colors.push(baseHex);

	// 左側相鄰色 (-30度)
	colors.push(hslToHex((hsl.h - 30 + 360) % 360, hsl.s, hsl.l));

	// 右側相鄰色 (+30度)
	colors.push(hslToHex((hsl.h + 30) % 360, hsl.s, hsl.l));

	// 更遠的左側色 (-60度)
	colors.push(hslToHex((hsl.h - 60 + 360) % 360, hsl.s, hsl.l));

	// 更遠的右側色 (+60度)
	colors.push(hslToHex((hsl.h + 60) % 360, hsl.s, hsl.l));

	return colors;
}

/**
 * 互補色配色方案（Complementary）
 * 使用色相環對面的顏色
 * @param {string} baseHex - 主色 HEX 值
 * @returns {Array} 5 個顏色的 HEX 陣列
 */
function generateComplementary(baseHex) {
	const hsl = hexToHsl(baseHex);
	const colors = [];

	// 主色
	colors.push(baseHex);

	// 互補色 (+180度)
	const complementHue = (hsl.h + 180) % 360;
	colors.push(hslToHex(complementHue, hsl.s, hsl.l));

	// 主色的淺色變化
	colors.push(hslToHex(hsl.h, hsl.s, Math.min(90, hsl.l + 25)));

	// 互補色的深色變化
	colors.push(hslToHex(complementHue, hsl.s, Math.max(10, hsl.l - 25)));

	// 中性色（降低飽和度的主色）
	colors.push(hslToHex(hsl.h, Math.max(10, hsl.s - 50), hsl.l));

	return colors;
}

/**
 * 分裂互補色配色方案（Split Complementary）
 * 使用互補色兩側的顏色
 * @param {string} baseHex - 主色 HEX 值
 * @returns {Array} 5 個顏色的 HEX 陣列
 */
function generateSplitComplementary(baseHex) {
	const hsl = hexToHsl(baseHex);
	const colors = [];

	// 主色
	colors.push(baseHex);

	// 分裂互補色 1 (+150度)
	colors.push(hslToHex((hsl.h + 150) % 360, hsl.s, hsl.l));

	// 分裂互補色 2 (+210度)
	colors.push(hslToHex((hsl.h + 210) % 360, hsl.s, hsl.l));

	// 主色的變化
	colors.push(
		hslToHex(hsl.h, Math.max(20, hsl.s - 30), Math.min(80, hsl.l + 20)),
	);

	// 分裂互補色的平均
	colors.push(hslToHex((hsl.h + 180) % 360, Math.max(20, hsl.s - 20), hsl.l));

	return colors;
}

/**
 * 三分色配色方案（Triadic）
 * 使用色相環上等距的三個顏色
 * @param {string} baseHex - 主色 HEX 值
 * @returns {Array} 5 個顏色的 HEX 陣列
 */
function generateTriadic(baseHex) {
	const hsl = hexToHsl(baseHex);
	const colors = [];

	// 主色
	colors.push(baseHex);

	// 三分色 1 (+120度)
	colors.push(hslToHex((hsl.h + 120) % 360, hsl.s, hsl.l));

	// 三分色 2 (+240度)
	colors.push(hslToHex((hsl.h + 240) % 360, hsl.s, hsl.l));

	// 主色的淺色變化
	colors.push(hslToHex(hsl.h, hsl.s, Math.min(85, hsl.l + 20)));

	// 主色的深色變化
	colors.push(hslToHex(hsl.h, hsl.s, Math.max(15, hsl.l - 20)));

	return colors;
}

/**
 * 四方色配色方案（Tetradic）
 * 使用色相環上兩對互補色
 * @param {string} baseHex - 主色 HEX 值
 * @returns {Array} 5 個顏色的 HEX 陣列
 */
function generateTetradic(baseHex) {
	const hsl = hexToHsl(baseHex);
	const colors = [];

	// 主色
	colors.push(baseHex);

	// 相鄰色 (+90度)
	colors.push(hslToHex((hsl.h + 90) % 360, hsl.s, hsl.l));

	// 互補色 (+180度)
	colors.push(hslToHex((hsl.h + 180) % 360, hsl.s, hsl.l));

	// 第四個色 (+270度)
	colors.push(hslToHex((hsl.h + 270) % 360, hsl.s, hsl.l));

	// 中性色（降低飽和度）
	colors.push(hslToHex(hsl.h, Math.max(15, hsl.s - 40), hsl.l));

	return colors;
}

/**
 * 黃金比例配色方案（Golden Ratio）
 * 使用黃金比例角度 (137.5°) 計算配色
 * @param {string} baseHex - 主色 HEX 值
 * @returns {Array} 5 個顏色的 HEX 陣列
 */
function generateGoldenRatio(baseHex) {
	const hsl = hexToHsl(baseHex);
	const colors = [];
	const goldenAngle = 137.5;

	// 主色
	colors.push(baseHex);

	// 黃金比例色彩
	for (let i = 1; i < 5; i++) {
		const newHue = (hsl.h + goldenAngle * i) % 360;
		colors.push(hslToHex(newHue, hsl.s, hsl.l));
	}

	return colors;
}

/**
 * 生成配色方案的主要函數
 * @param {string} baseHex - 主色 HEX 值
 * @param {string} scheme - 配色方案類型
 * @returns {Array} 5 個顏色的 HEX 陣列
 */
function generateColorPalette(baseHex, scheme) {
	// 驗證和標準化輸入
	if (!isValidHex(baseHex)) {
		throw new Error("無效的 HEX 顏色格式");
	}

	baseHex = normalizeHex(baseHex);

	switch (scheme.toLowerCase()) {
		case "monochromatic":
			return generateMonochromatic(baseHex);
		case "analogous":
			return generateAnalogous(baseHex);
		case "complementary":
			return generateComplementary(baseHex);
		case "splitcomplementary":
		case "split-complementary":
			return generateSplitComplementary(baseHex);
		case "triadic":
			return generateTriadic(baseHex);
		case "tetradic":
			return generateTetradic(baseHex);
		case "goldenratio":
		case "golden-ratio":
			return generateGoldenRatio(baseHex);
		default:
			throw new Error("不支援的配色方案類型");
	}
}

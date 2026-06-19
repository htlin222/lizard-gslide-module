// TODO(shared-migration): replace the typeof *_color fallback block with
//   getThemeColors() (shared/theme_colors.js); see kpi_minter.js for the pattern.
//   If/when this minter moves to batch requests, use shared/shape_requests.js builders.
/**
 * Server-side core for the Gallery Minter (圖片藝廊鑄造器) dialog.
 *
 * Takes a list of image URLs (one per line, optional "| caption" suffix) and
 * lays them out as a grid of images on the current slide. Columns are chosen
 * automatically from the image count (grid-style equal layout) unless the user
 * overrides them. Each image is fitted (aspect-preserving) into its grid cell,
 * with an optional caption text box rendered below.
 *
 * Mirrors the Callout / Grid Minter pattern:
 *  - getGalleryTemplates() feeds the dialog's template picker + live preview.
 *  - insertGalleryIntoSlide(payload) inserts onto the current slide.
 *  - Positions are computed like grid_minter (start Y ~120, margin 30, gap 12).
 *
 * Images are fetched by Google Slides itself via slide.insertImage(url), so the
 * URLs MUST be publicly accessible. Each insertImage() is wrapped in try/catch;
 * a failed URL is collected as a warning and never aborts the rest of the batch.
 */

/**
 * Single source of truth for gallery templates. Colors resolve from the global
 * palette so they track the configured theme. `border` controls whether each
 * image gets an outline; `borderColor` is the outline color when enabled.
 *
 * @return {Array<{id,name,border,borderColor,borderWidth,captionColor,swatch}>}
 */
function buildGalleryTemplates_() {
	const main = (typeof main_color !== "undefined" && main_color) || "#3D6869";
	const accent =
		(typeof accent_color !== "undefined" && accent_color) || "#f29424";
	const text =
		(typeof text_color !== "undefined" && text_color) || "#333333";
	return [
		{
			id: "plain",
			name: "Plain",
			border: false,
			borderColor: "#FFFFFF",
			borderWidth: 0,
			captionColor: text,
			swatch: "#DDDDDD",
		},
		{
			id: "framed",
			name: "Framed",
			border: true,
			borderColor: main,
			borderWidth: 2,
			captionColor: main,
			swatch: main,
		},
		{
			id: "accent",
			name: "Accent",
			border: true,
			borderColor: accent,
			borderWidth: 2,
			captionColor: accent,
			swatch: accent,
		},
		{
			id: "thin",
			name: "Thin",
			border: true,
			borderColor: "#999999",
			borderWidth: 1,
			captionColor: text,
			swatch: "#999999",
		},
	];
}

/**
 * Returns the gallery templates for client-side preview. Called from the dialog
 * through google.script.run.
 * @return {Array<Object>}
 */
function getGalleryTemplates() {
	return buildGalleryTemplates_();
}

/**
 * Suggests a column count for `n` images. Keeps galleries roughly landscape to
 * match 16:9 slides. Mirrors the client-side suggestion in the dialog scripts.
 *
 * @param {number} n - image count
 * @return {number}
 */
function suggestGalleryCols_(n) {
	if (n <= 0) return 1;
	if (n === 1) return 1;
	if (n === 2) return 2;
	if (n <= 4) return 2;
	if (n <= 6) return 3;
	if (n <= 9) return 3;
	return 4;
}

/**
 * Computes image-cell rectangles (in PT) for a rows×cols gallery grid. Columns
 * divide the usable width evenly; rows are sized evenly within the usable
 * height. Mirrors grid_minter's geometry (start Y ~120, margin 30, gap 12).
 *
 * @param {number} rows
 * @param {number} cols
 * @param {number} pageW - page width in PT
 * @param {number} pageH - page height in PT
 * @param {{margin?: number, gap?: number, top?: number, left?: number}} [opts]
 * @return {Array<{x: number, y: number, w: number, h: number}>} row-major order
 */
function computeGalleryPositions_(rows, cols, pageW, pageH, opts) {
	const o = opts || {};
	const margin = o.margin != null ? o.margin : 30;
	const gap = o.gap != null ? o.gap : 12;
	const left = o.left != null ? o.left : margin;
	const top = o.top != null ? o.top : 120;

	const usableW = pageW - left - margin;
	const cellW = (usableW - (cols - 1) * gap) / cols;

	const usableH = pageH - top - margin;
	const cellH = (usableH - (rows - 1) * gap) / rows;

	const positions = [];
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			positions.push({
				x: left + c * (cellW + gap),
				y: top + r * (cellH + gap),
				w: cellW,
				h: cellH,
			});
		}
	}
	return positions;
}

/**
 * Inserts a gallery of images onto the current slide.
 *
 * Each image is inserted via slide.insertImage(url) — which makes Google Slides
 * fetch the URL — then fitted (aspect-preserving) into its grid cell and
 * centered. When a caption is provided a text box is added below the image. Any
 * URL that fails to insert is collected and returned in `warnings`; one bad URL
 * never aborts the rest.
 *
 * @param {{items: Array<{url: string, caption?: string}>, cols?: number,
 *   templateId?: string, captions?: boolean}} payload
 * @return {{success: boolean, error?: string, warnings?: string[]}}
 */
function insertGalleryIntoSlide(payload) {
	try {
		const p = payload || {};
		const items = (p.items || []).filter(function (it) {
			return it && it.url && String(it.url).trim();
		});
		if (!items.length) {
			return { success: false, error: "No image URLs to insert." };
		}

		const templates = buildGalleryTemplates_();
		let tpl = templates[0];
		for (let i = 0; i < templates.length; i++) {
			if (templates[i].id === p.templateId) tpl = templates[i];
		}

		const showCaptions = p.captions !== false; // default to showing captions
		const font =
			(typeof main_font_family !== "undefined" && main_font_family) ||
			"Source Sans Pro";

		let cols = p.cols > 0 ? Math.floor(p.cols) : 0;
		if (!cols) cols = suggestGalleryCols_(items.length);
		const rows = Math.ceil(items.length / cols);

		const presentation = SlidesApp.getActivePresentation();
		const pageW = presentation.getPageWidth();
		const pageH = presentation.getPageHeight();

		// Resolve the target slide (fall back to the first slide).
		let slide = null;
		try {
			slide = presentation.getSelection().getCurrentPage().asSlide();
		} catch (e) {
			slide = presentation.getSlides()[0];
		}
		if (!slide) return { success: false, error: "No slide available." };

		const positions = computeGalleryPositions_(rows, cols, pageW, pageH, {});

		// Caption strip reserved at the bottom of each cell when captions are on.
		const captionH = 18;
		const captionGap = 4;

		const warnings = [];
		const cropRequests = [];

		items.forEach(function (item, i) {
			const pos = positions[i];
			const url = String(item.url).trim();
			const caption = (item.caption || "").trim();
			const hasCaption = showCaptions && caption;

			// Region available for the image (leave room for a caption below).
			const imgRegionH = hasCaption
				? Math.max(20, pos.h - captionH - captionGap)
				: pos.h;

			let image = null;
			try {
				image = slide.insertImage(url);
			} catch (imgErr) {
				warnings.push("無法載入圖片: " + url);
				return; // skip this one, keep going
			}

			try {
				// Match the dialog preview, which renders every image as a uniform
				// cell with `object-fit: cover` (fills the whole cell, cropped to the
				// cell aspect — NOT letterbox-fit). Read the intrinsic size BEFORE
				// resizing (insertImage leaves the element at its natural dimensions),
				// then fill the cell and crop the overflowing edges so the visible
				// area keeps the cell aspect. Border wraps the full cell, like preview.
				const natW = image.getWidth();
				const natH = image.getHeight();

				image.setWidth(pos.w);
				image.setHeight(imgRegionH);
				image.setLeft(pos.x);
				image.setTop(pos.y);

				// Compute a "cover" crop: trim the long axis so the remaining region
				// matches the cell aspect. CropProperties offsets are fractions of the
				// (post-crop) image dimensions chopped off each edge.
				if (natW > 0 && natH > 0) {
					const cellAspect = pos.w / imgRegionH;
					const natAspect = natW / natH;
					let leftOff = 0;
					let rightOff = 0;
					let topOff = 0;
					let bottomOff = 0;
					if (natAspect > cellAspect) {
						// Image too wide → crop left/right.
						const keep = cellAspect / natAspect; // fraction of width kept
						const trim = (1 - keep) / 2;
						leftOff = trim;
						rightOff = trim;
					} else if (natAspect < cellAspect) {
						// Image too tall → crop top/bottom.
						const keep = natAspect / cellAspect; // fraction of height kept
						const trim = (1 - keep) / 2;
						topOff = trim;
						bottomOff = trim;
					}
					if (leftOff || rightOff || topOff || bottomOff) {
						cropRequests.push({
							updateImageProperties: {
								objectId: image.getObjectId(),
								fields: "cropProperties",
								imageProperties: {
									cropProperties: {
										leftOffset: leftOff,
										rightOffset: rightOff,
										topOffset: topOff,
										bottomOffset: bottomOff,
									},
								},
							},
						});
					}
				}

				if (tpl.border && tpl.borderWidth > 0) {
					image.getBorder().getLineFill().setSolidFill(tpl.borderColor);
					image.getBorder().setWeight(tpl.borderWidth);
				}
			} catch (sizeErr) {
				// Image inserted but sizing failed — leave it as-is, not fatal.
			}

			// Caption text box below the image region.
			if (hasCaption) {
				try {
					const capBox = slide.insertShape(
						SlidesApp.ShapeType.TEXT_BOX,
						pos.x,
						pos.y + imgRegionH + captionGap,
						pos.w,
						captionH,
					);
					capBox.getText().setText(caption);
					capBox
						.getText()
						.getTextStyle()
						.setForegroundColor(tpl.captionColor)
						.setFontSize(10)
						.setFontFamily(font);
					capBox
						.getText()
						.getParagraphStyle()
						.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
					capBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
				} catch (capErr) {
					// Caption is best-effort; don't fail the whole insert.
				}
			}
		});

		// Apply all "cover" crops in one batch. Cropping is cosmetic; if it fails
		// the images simply remain stretched-to-fill rather than cropped, so we
		// degrade gracefully and still report success.
		if (cropRequests.length) {
			try {
				Slides.Presentations.batchUpdate(
					{ requests: cropRequests },
					presentation.getId(),
				);
			} catch (cropErr) {
				console.warn("Gallery crop step failed: " + cropErr.message);
			}
		}

		return { success: true, warnings: warnings };
	} catch (e) {
		console.error("Error inserting gallery: " + e.message);
		return { success: false, error: e.message };
	}
}

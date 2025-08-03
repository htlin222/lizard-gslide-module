// 🚀 CACHE MANAGEMENT MODULE - Optimized resource caching
/**
 * Centralized cache management for ultra-performance batch processing
 * - Pre-calculated colors and transforms
 * - GUID pool management
 * - Slide property caching
 */

var ULTRA_CACHE = null;

/**
 * Initialize the ultra performance cache
 */
function initializeUltraCache() {
	if (ULTRA_CACHE) return ULTRA_CACHE;
	
	ULTRA_CACHE = {
		// Pre-calculated RGB colors (avoid repeated hexToRgb calls)
		colors: {
			main: hexToRgb(main_color),
			inactive: hexToRgb('#888888'),
			gray: hexToRgb('#E0E0E0'),
			white: hexToRgb('#FFFFFF'),
			bgColor: hexToRgb('#FFFFFF')
		},
		// Pre-generated GUIDs batch (avoid expensive UUID generation)
		guids: [],
		guidIndex: 0,
		// Reusable objects to reduce memory allocation
		transforms: {
			identity: { scaleX: 1, scaleY: 1, unit: 'PT' },
			rotation90: { scaleX: 0, shearX: -1, shearY: 1, scaleY: 0, unit: 'PT' }
		}
	};
	
	// Pre-generate 1000 GUIDs
	for (let i = 0; i < 1000; i++) {
		ULTRA_CACHE.guids.push(Utilities.getUuid().replace(/-/g, '').slice(0, 8));
	}
	
	return ULTRA_CACHE;
}

/**
 * Get next GUID from the pre-generated pool
 */
function getNextGuid() {
	const cache = initializeUltraCache();
	const guid = cache.guids[cache.guidIndex];
	cache.guidIndex = (cache.guidIndex + 1) % cache.guids.length;
	return guid;
}

/**
 * Create optimized slide cache with pre-calculated properties
 */
function createUltraSlideCache(presentation, slides) {
	const width = presentation.getPageWidth();
	const height = presentation.getPageHeight();
	
	return {
		width, height, totalSlides: slides.length,
		// Pre-calculated positions
		progressBarY: height - progressBarHeight,
		rightFooterX: width,
		centerX: width / 2,
		// Pre-calculated common sizes
		sizes: {
			pageNum: { width: { magnitude: 70, unit: 'PT' }, height: { magnitude: 30, unit: 'PT' } },
			progressBar: { height: { magnitude: progressBarHeight, unit: 'PT' } },
			tabBar: { height: { magnitude: 14, unit: 'PT' } },
			sectionBox: { width: { magnitude: 600, unit: 'PT' }, height: { magnitude: 150, unit: 'PT' } }
		},
		// Slide data with minimal API calls
		slideData: slides.map((slide, idx) => ({
			id: slide.getObjectId(),
			layoutName: idx === 0 ? null : slide.getLayout().getLayoutName(), // Skip first slide
			slide: idx < 2 ? slide : null // Only keep references we actually need
		}))
	};
}

/**
 * Clear cache (for memory management)
 */
function clearUltraCache() {
	ULTRA_CACHE = null;
}
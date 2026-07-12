/**
 * Shared target-slide resolver for all minters' insert<X>IntoSlide(payload).
 *
 * Priority: payload.pageObjectId (batch callers like html2slides pass the id
 * of a freshly created slide) → the user's current selection → first slide.
 *
 * When pageObjectId IS given but cannot be resolved, this returns null and
 * the minter fails loudly — it must NOT fall back to the selection, because
 * drawing on whatever slide the user happens to have open is worse than a
 * clean per-slide error (this exact bug once dumped every minter's output
 * onto slide 1).
 *
 * @param {Presentation} presentation
 * @param {string=} pageObjectId
 * @return {Slide|null}
 */
function resolveMinterTargetSlide_(presentation, pageObjectId) {
	if (pageObjectId) {
		try {
			const byId = presentation.getSlideById(pageObjectId);
			if (byId) return byId;
		} catch (e) {
			// try a fresh handle below
		}
		// The caller's handle may be stale (e.g. after saveAndClose) — retry
		// with a freshly opened one before giving up.
		try {
			const fresh = SlidesApp.openById(presentation.getId());
			const byId2 = fresh.getSlideById(pageObjectId);
			if (byId2) return byId2;
		} catch (e) {
			// give up below
		}
		return null;
	}
	try {
		return presentation.getSelection().getCurrentPage().asSlide();
	} catch (e) {
		return presentation.getSlides()[0] || null;
	}
}

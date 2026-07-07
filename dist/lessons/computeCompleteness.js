"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.computeCompleteness = computeCompleteness;
// Returns the highest completeness tier met by the lesson, or null if none apply.
// Tiers (highest to lowest):
//   great   : ≥1 cause, ≥1 impact, every cause AND every impact has ≥1 linked action/fpc
//   good    : ≥1 cause, ≥1 impact, every cause has ≥1 linked action/fpc
//   okay    : ≥1 cause, ≥1 impact, ≥1 action/fpc linked to any cause or impact
//   minimum : ≥1 impact, ≥1 action/fpc linked to any impact
function computeCompleteness(s) {
    const hasCause = s.causeIds.length > 0;
    const hasImpact = s.impactIds.length > 0;
    // Normalise ids to strings for comparison so numeric and string ids match.
    const causeIdSet = new Set(s.causeIds.map(String));
    const impactIdSet = new Set(s.impactIds.map(String));
    // Helpers: does a given cause/impact id have at least one linked item?
    const causeHasLink = (id) => s.linkedItems.some((li) => li.causeId != null && String(li.causeId) === String(id));
    const impactHasLink = (id) => s.linkedItems.some((li) => li.impactId != null && String(li.impactId) === String(id));
    // great
    if (hasCause &&
        hasImpact &&
        [...causeIdSet].every(causeHasLink) &&
        [...impactIdSet].every(impactHasLink)) {
        return 'great';
    }
    // good
    if (hasCause && hasImpact && [...causeIdSet].every(causeHasLink)) {
        return 'good';
    }
    // okay
    if (hasCause &&
        hasImpact &&
        s.linkedItems.some((li) => li.causeId != null || li.impactId != null)) {
        return 'okay';
    }
    // minimum
    if (hasImpact && s.linkedItems.some((li) => li.impactId != null)) {
        return 'minimum';
    }
    return null;
}
//# sourceMappingURL=computeCompleteness.js.map
// ms-project-txt-parser.js
// Placeholder parser for MS Project TXT exports used in the
// Lessons Learned Metadata sub-module of the user portal.
// 
// You will provide the actual parsing logic later. For now, this file
// simply exposes an async function that reads the file text so that
// wiring and imports can be verified end-to-end.

/**
 * Parse a TXT file that was exported from MS Project.
 *
 * First responsibility: verify which major sections are present:
 * - <Resources>...</Resources>
 * - <Tasks>...</Tasks>
 * - <Assignments>...</Assignments>
 *
 * @param {File} file - The TXT file selected by the user.
 * @returns {Promise<object>} Summary info including which sections are present.
 */
export async function parseMsProjectTxt(file) {
    if (!file) {
        throw new Error('No file provided to parseMsProjectTxt.');
    }

    // Read the file contents as text.
    const text = await file.text();

    // Detect presence of each section by the opening and closing tags.
    const hasResources =
        text.includes('<Resources>') && text.includes('</Resources>');
    const hasTasks =
        text.includes('<Tasks>') && text.includes('</Tasks>');
    const hasAssignments =
        text.includes('<Assignments>') && text.includes('</Assignments>');

    const presentSections = [];
    if (hasResources) presentSections.push('Resources');
    if (hasTasks) presentSections.push('Tasks');
    if (hasAssignments) presentSections.push('Assignments');

    const presentCount = presentSections.length;

    // Log for debugging in the console.
    console.log('[ms-project-txt-parser] Section summary:', {
        fileName: file.name,
        hasResources,
        hasTasks,
        hasAssignments,
        presentSections,
        presentCount
    });

    // We deliberately do not throw if sections are missing; the caller
    // will turn this into a friendly user message.
    return {
        fileName: file.name,
        size: file.size,
        hasResources,
        hasTasks,
        hasAssignments,
        presentSections,
        presentCount
    };
}



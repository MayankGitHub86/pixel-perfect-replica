// backend/scripts/lib/update-workbook-pointer.js
//
// Updates or creates the singleton ExcelWorkbook record in PostgreSQL.
// ExcelWorkbook has a String UUID primary key (`id`) and unique `driveItemId`.
// Rules:
// 1. If exactly one record exists: updates it using its actual UUID string.
// 2. If no record exists: creates a fresh record with valid fields.
// 3. If multiple records exist: aborts with an error requiring manual investigation.
//    NEVER deletes or modifies any record when multiple exist.

'use strict';

/**
 * Updates the database pointer for ExcelWorkbook atomically.
 *
 * @param {object} tx - Prisma transaction or client instance
 * @param {object} params
 * @param {string} params.driveItemId - OneDrive DriveItem ID of the new workbook
 * @param {string} params.webUrl - OneDrive web URL of the new workbook
 * @returns {Promise<object>} The updated or created ExcelWorkbook record
 * @throws {Error} If multiple records exist, parameters are invalid, or DB fails
 */
async function updateWorkbookPointer(tx, { driveItemId, webUrl }) {
  if (!driveItemId || typeof driveItemId !== 'string') {
    throw new Error('updateWorkbookPointer requires a valid string driveItemId');
  }
  if (!webUrl || typeof webUrl !== 'string') {
    throw new Error('updateWorkbookPointer requires a valid string webUrl');
  }

  const existingRecords = await tx.excelWorkbook.findMany({
    orderBy: { createdAt: 'asc' },
  });

  // Multiple records exist: abort without modifying or deleting anything
  if (existingRecords.length > 1) {
    throw new Error(
      `Multiple ExcelWorkbook records found in database (${existingRecords.length} records). ` +
      'Aborting update to prevent ambiguity or data loss. Manual investigation is required.'
    );
  }

  // Exactly one record exists: update it using its actual UUID string
  if (existingRecords.length === 1) {
    const primaryRecord = existingRecords[0];
    return tx.excelWorkbook.update({
      where: { id: primaryRecord.id },
      data: { driveItemId, webUrl },
    });
  }

  // Zero records exist: create a fresh record
  return tx.excelWorkbook.create({
    data: { driveItemId, webUrl },
  });
}

module.exports = { updateWorkbookPointer };

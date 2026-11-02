const { PrismaClient } = require('@prisma/client');
(async () => {
  const prisma = new PrismaClient();
  const list = await prisma.excelWorkbook.findMany({ orderBy: { createdAt: 'asc' } });
  console.log('EXACT_COUNT:', list.length);
  list.forEach((wb, i) => {
    let cleanUrl = '';
    try {
      const u = new URL(wb.webUrl);
      cleanUrl = `${u.origin}${u.pathname}`;
    } catch {
      cleanUrl = '(unparseable)';
    }
    console.log(`RECORD_${i + 1}:`, JSON.stringify({
      id: wb.id,
      driveItemId: wb.driveItemId,
      webUrlClean: cleanUrl,
      createdAt: wb.createdAt,
      updatedAt: wb.updatedAt,
    }, null, 2));
  });
  await prisma.$disconnect();
})();


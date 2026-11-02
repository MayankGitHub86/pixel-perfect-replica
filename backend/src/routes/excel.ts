// backend/src/routes/excel.ts
import { Router, Request, Response, NextFunction } from 'express';
import { prisma } from '../lib/db';

export const excelRouter = Router();

// GET /api/excel-sync/status
excelRouter.get('/status', async (req: Request, res: Response, next: NextFunction) => {
  try {
    const [recentRuns, lastSuccess, failedJobsCount, totalCandidates, syncedCandidates] = await Promise.all([
      prisma.excelSyncRun.findMany({ orderBy: { createdAt: 'desc' }, take: 10 }),
      prisma.excelSyncRun.findFirst({ where: { result: 'success' }, orderBy: { createdAt: 'desc' } }),
      prisma.processingJob.count({ where: { type: 'excel_sync', status: 'failed' } }),
      prisma.candidate.count(),
      prisma.excelSyncRun.count({ where: { result: 'success' } }),
    ]);

    const pendingRows = Math.max(0, totalCandidates - syncedCandidates);

    // Real workbook info via service
    const { getExcelStatus } = await import('../lib/excel');
    const realStatus = await getExcelStatus();

    res.json({
      workbookName: realStatus.workbookName,
      state: realStatus.state,
      storageProvider: realStatus.storageProvider,
      lastSuccessfulSync: realStatus.lastSuccessfulSync,
      pendingRows,
      failedJobs: failedJobsCount,
      lastReconciliation: realStatus.lastReconciliation,
      workbookUrl: realStatus.workbookUrl,
      history: recentRuns.map((r) => ({
        id: r.id,
        at: r.createdAt.toISOString(),
        rows: r.rows,
        result: r.result as any,
        message: r.message,
      })),
    });
  } catch (err) {
    next(err);
  }
});

// POST /api/excel-sync/test
excelRouter.post('/test', async (req: Request, res: Response) => {
  try {
    const { ensureWorkbookExists, ensureWorksheetAndTable } = await import('../lib/excel');
    const wb = await ensureWorkbookExists();
    await ensureWorksheetAndTable(wb.driveItemId);
    res.json({ ok: true, message: 'Workbook, worksheet, and table are reachable.' });
  } catch (e: any) {
    res.status(500).json({ ok: false, message: e.message ?? 'Graph test failed' });
  }
});

// POST /api/excel-sync/run
excelRouter.post('/run', async (req: Request, res: Response) => {
  try {
    const { checkCandidateExcelEligibility } = await import('../lib/excel');
    // Enqueue sync only for eligible candidates that have not yet succeeded
    const unsynced = await prisma.candidate.findMany({
      where: { NOT: { excelSyncRuns: { some: { result: 'success' } } } },
      select: { id: true },
    });

    const { excelSyncQueue } = await import('../lib/queue');
    let queuedCount = 0;
    for (const c of unsynced) {
      const eligibility = await checkCandidateExcelEligibility(c.id);
      if (eligibility.eligible) {
        await excelSyncQueue.add('sync', { candidateId: c.id });
        queuedCount++;
      }
    }
    res.json({ queued: queuedCount, totalEvaluated: unsynced.length });
  } catch (e: any) {
    res.status(500).json({ error: { code: 'EXCEL_SYNC_ERROR', message: e.message ?? 'Sync failed' } });
  }
});

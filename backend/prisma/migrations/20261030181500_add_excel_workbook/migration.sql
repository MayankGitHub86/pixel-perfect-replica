-- CreateTable
CREATE TABLE "ExcelWorkbook" (
    "id" TEXT NOT NULL,
    "driveItemId" TEXT NOT NULL,
    "webUrl" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "ExcelWorkbook_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "ExcelWorkbook_driveItemId_key" ON "ExcelWorkbook"("driveItemId");


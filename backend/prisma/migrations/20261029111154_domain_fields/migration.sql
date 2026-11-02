/*
  Warnings:

  - You are about to drop the column `jobTitle` on the `Application` table. All the data in the column will be lost.
  - You are about to drop the column `firstName` on the `Candidate` table. All the data in the column will be lost.
  - You are about to drop the column `lastName` on the `Candidate` table. All the data in the column will be lost.
  - You are about to drop the column `linkedinUrl` on the `Candidate` table. All the data in the column will be lost.
  - Added the required column `role` to the `Application` table without a default value. This is not possible if the table is not empty.
  - Added the required column `related` to the `ProcessingJob` table without a default value. This is not possible if the table is not empty.
  - Added the required column `type` to the `ProcessingJob` table without a default value. This is not possible if the table is not empty.

*/
-- DropIndex
DROP INDEX "AuditLog_entity_idx";

-- DropIndex
DROP INDEX "ProcessingJob_emailId_idx";

-- AlterTable
ALTER TABLE "Application" DROP COLUMN "jobTitle",
ADD COLUMN     "recruiter" TEXT,
ADD COLUMN     "role" TEXT NOT NULL,
ADD COLUMN     "sourceEmail" TEXT NOT NULL DEFAULT '',
ALTER COLUMN "status" SET DEFAULT 'new';

-- AlterTable
ALTER TABLE "Attachment" ALTER COLUMN "emailId" DROP NOT NULL;

-- AlterTable
ALTER TABLE "AuditLog" ADD COLUMN     "actor" TEXT NOT NULL DEFAULT 'system',
ADD COLUMN     "related" TEXT NOT NULL DEFAULT '',
ADD COLUMN     "requestId" TEXT NOT NULL DEFAULT '',
ADD COLUMN     "result" TEXT NOT NULL DEFAULT 'info',
ALTER COLUMN "entity" SET DEFAULT '';

-- AlterTable
ALTER TABLE "Candidate" DROP COLUMN "firstName",
DROP COLUMN "lastName",
DROP COLUMN "linkedinUrl",
ADD COLUMN     "appliedRole" TEXT,
ADD COLUMN     "college" TEXT,
ADD COLUMN     "degree" TEXT,
ADD COLUMN     "experienceYears" DOUBLE PRECISION,
ADD COLUMN     "graduationYear" INTEGER,
ADD COLUMN     "interviewStatus" TEXT NOT NULL DEFAULT 'not_scheduled',
ADD COLUMN     "name" TEXT,
ADD COLUMN     "processingStatus" TEXT NOT NULL DEFAULT 'new',
ADD COLUMN     "receivedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
ADD COLUMN     "resumeFileName" TEXT,
ADD COLUMN     "resumeRef" TEXT,
ADD COLUMN     "reviewFields" TEXT[] DEFAULT ARRAY[]::TEXT[],
ADD COLUMN     "skills" TEXT[] DEFAULT ARRAY[]::TEXT[],
ADD COLUMN     "sourceEmailId" TEXT;

-- AlterTable
ALTER TABLE "EmailMessage" ADD COLUMN     "applicationId" TEXT,
ADD COLUMN     "candidateId" TEXT,
ADD COLUMN     "error" TEXT,
ADD COLUMN     "history" JSONB,
ALTER COLUMN "mailboxId" DROP NOT NULL,
ALTER COLUMN "receivedAt" SET DEFAULT CURRENT_TIMESTAMP;

-- AlterTable
ALTER TABLE "ExcelSyncRun" ADD COLUMN     "message" TEXT NOT NULL DEFAULT '',
ADD COLUMN     "result" TEXT NOT NULL DEFAULT 'success',
ADD COLUMN     "rows" INTEGER NOT NULL DEFAULT 1,
ALTER COLUMN "candidateId" DROP NOT NULL;

-- AlterTable
ALTER TABLE "ProcessingJob" ADD COLUMN     "completedAt" TIMESTAMP(3),
ADD COLUMN     "nextRetryAt" TIMESTAMP(3),
ADD COLUMN     "related" TEXT NOT NULL,
ADD COLUMN     "startedAt" TIMESTAMP(3),
ADD COLUMN     "status" TEXT NOT NULL DEFAULT 'queued',
ADD COLUMN     "type" TEXT NOT NULL,
ALTER COLUMN "emailId" DROP NOT NULL;

-- CreateIndex
CREATE INDEX "Application_role_idx" ON "Application"("role");

-- CreateIndex
CREATE INDEX "AuditLog_actor_idx" ON "AuditLog"("actor");

-- CreateIndex
CREATE INDEX "AuditLog_result_idx" ON "AuditLog"("result");

-- CreateIndex
CREATE INDEX "Candidate_processingStatus_idx" ON "Candidate"("processingStatus");

-- CreateIndex
CREATE INDEX "Candidate_interviewStatus_idx" ON "Candidate"("interviewStatus");

-- CreateIndex
CREATE INDEX "Candidate_receivedAt_idx" ON "Candidate"("receivedAt");

-- CreateIndex
CREATE INDEX "Candidate_appliedRole_idx" ON "Candidate"("appliedRole");

-- CreateIndex
CREATE INDEX "ExcelSyncRun_createdAt_idx" ON "ExcelSyncRun"("createdAt");

-- CreateIndex
CREATE INDEX "ProcessingJob_type_idx" ON "ProcessingJob"("type");

-- CreateIndex
CREATE INDEX "ProcessingJob_status_idx" ON "ProcessingJob"("status");

-- CreateIndex
CREATE INDEX "ProcessingJob_createdAt_idx" ON "ProcessingJob"("createdAt");

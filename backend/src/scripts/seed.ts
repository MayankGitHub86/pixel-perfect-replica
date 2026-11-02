import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function main() {
  if (process.env.NODE_ENV === 'production' && !process.argv.includes('--force')) {
    console.error('ERROR: Refusing to seed synthetic data in production environment.');
    process.exit(1);
  }

  console.log('Seeding development synthetic candidate records...');

  // Delete previous synthetic records to keep seed clean and reproducible
  await prisma.auditLog.deleteMany({ where: { actor: 'dev-seed' } });
  await prisma.application.deleteMany({ where: { sourceEmail: { contains: 'synthetic.dev' } } });
  await prisma.processingJob.deleteMany({ where: { related: { contains: 'SYNTHETIC' } } });
  await prisma.attachment.deleteMany({ where: { filename: { contains: 'SYNTHETIC' } } });
  await prisma.emailMessage.deleteMany({ where: { fromEmail: { contains: 'synthetic.dev' } } });
  await prisma.candidate.deleteMany({ where: { email: { contains: 'synthetic.dev' } } });

  const syntheticCandidates = [
    {
      name: '[DEV-SYNTHETIC] Sarah Jenkins',
      email: 'sarah.jenkins@synthetic.dev',
      phone: '+1 (555) 234-5678',
      college: 'Stanford University',
      degree: 'B.S. Computer Science',
      graduationYear: 2022,
      skills: ['TypeScript', 'React', 'Node.js', 'PostgreSQL', 'Docker'],
      experienceYears: 3.5,
      appliedRole: 'Senior Full Stack Engineer',
      processingStatus: 'processed',
      interviewStatus: 'scheduled',
      resumeRef: 'resumes/synthetic_sarah_jenkins.pdf',
      resumeFileName: '[DEV-SYNTHETIC] Sarah_Jenkins_Resume.pdf',
      sourceEmailId: 'SYNTH-MSG-001',
      reviewFields: [],
      receivedAt: new Date(Date.now() - 2 * 86_400_000),
    },
    {
      name: '[DEV-SYNTHETIC] Marcus Vance',
      email: 'marcus.vance@synthetic.dev',
      phone: '+1 (555) 345-6789',
      college: 'MIT',
      degree: 'M.S. Data Science',
      graduationYear: 2021,
      skills: ['Python', 'PyTorch', 'FastAPI', 'Redis', 'SQL'],
      experienceYears: 4.0,
      appliedRole: 'Machine Learning Engineer',
      processingStatus: 'needs_review',
      interviewStatus: 'not_scheduled',
      resumeRef: 'resumes/synthetic_marcus_vance.pdf',
      resumeFileName: '[DEV-SYNTHETIC] Marcus_Vance_Resume.pdf',
      sourceEmailId: 'SYNTH-MSG-002',
      reviewFields: ['degree', 'graduationYear'],
      receivedAt: new Date(Date.now() - 4 * 86_400_000),
    },
    {
      name: '[DEV-SYNTHETIC] Elena Rostova',
      email: 'elena.rostova@synthetic.dev',
      phone: '+1 (555) 456-7890',
      college: 'University of Michigan',
      degree: 'B.A. Graphic Design',
      graduationYear: 2020,
      skills: ['Figma', 'UI/UX', 'Design Systems', 'CSS', 'Tailwind'],
      experienceYears: 5.0,
      appliedRole: 'Lead Product Designer',
      processingStatus: 'processed',
      interviewStatus: 'completed',
      resumeRef: 'resumes/synthetic_elena_rostova.pdf',
      resumeFileName: '[DEV-SYNTHETIC] Elena_Rostova_Portfolio.pdf',
      sourceEmailId: 'SYNTH-MSG-003',
      reviewFields: [],
      receivedAt: new Date(Date.now() - 6 * 86_400_000),
    },
    {
      name: '[DEV-SYNTHETIC] David Kim',
      email: 'david.kim@synthetic.dev',
      phone: '+1 (555) 567-8901',
      college: 'UC Berkeley',
      degree: 'B.S. Electrical Engineering',
      graduationYear: 2023,
      skills: ['Go', 'Kubernetes', 'AWS', 'Terraform', 'Prometheus'],
      experienceYears: 2.0,
      appliedRole: 'DevOps Engineer',
      processingStatus: 'new',
      interviewStatus: 'not_scheduled',
      resumeRef: 'resumes/synthetic_david_kim.pdf',
      resumeFileName: '[DEV-SYNTHETIC] David_Kim_CV.pdf',
      sourceEmailId: 'SYNTH-MSG-004',
      reviewFields: [],
      receivedAt: new Date(Date.now() - 1 * 86_400_000),
    },
    {
      name: '[DEV-SYNTHETIC] Priya Sharma',
      email: 'priya.sharma@synthetic.dev',
      phone: '+1 (555) 678-9012',
      college: 'Carnegie Mellon University',
      degree: 'M.S. Software Engineering',
      graduationYear: 2019,
      skills: ['Java', 'Spring Boot', 'Microservices', 'Kafka', 'PostgreSQL'],
      experienceYears: 6.5,
      appliedRole: 'Staff Backend Architect',
      processingStatus: 'processed',
      interviewStatus: 'not_scheduled',
      resumeRef: 'resumes/synthetic_priya_sharma.pdf',
      resumeFileName: '[DEV-SYNTHETIC] Priya_Sharma_Resume.docx',
      sourceEmailId: 'SYNTH-MSG-005',
      reviewFields: [],
      receivedAt: new Date(Date.now() - 3 * 86_400_000),
    },
  ];

  for (const c of syntheticCandidates) {
    const candidate = await prisma.candidate.create({
      data: c,
    });

    // Create synthetic source email
    const email = await prisma.emailMessage.create({
      data: {
        messageId: c.sourceEmailId,
        subject: `Application for ${c.appliedRole} - ${c.name}`,
        fromEmail: c.email,
        receivedAt: c.receivedAt,
        status: c.processingStatus === 'needs_review' ? 'needs_review' : 'processed',
        candidateId: candidate.id,
        history: [
          { at: c.receivedAt.toISOString(), step: 'email_ingest', status: 'succeeded' },
          { at: new Date(c.receivedAt.getTime() + 5000).toISOString(), step: 'attachment_parse', status: 'succeeded' },
          { at: new Date(c.receivedAt.getTime() + 10000).toISOString(), step: 'field_extraction', status: 'succeeded' },
        ],
      },
    });

    // Create synthetic attachment
    await prisma.attachment.create({
      data: {
        emailId: email.id,
        candidateId: candidate.id,
        filename: c.resumeFileName,
        mimeType: c.resumeFileName.endsWith('.pdf') ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        sizeBytes: 145000,
        storageKey: c.resumeRef,
        status: 'done',
      },
    });

    // Create synthetic application
    await prisma.application.create({
      data: {
        candidateId: candidate.id,
        role: c.appliedRole,
        sourceEmail: c.email,
        status: c.interviewStatus === 'scheduled' ? 'interview' : c.processingStatus === 'processed' ? 'under_review' : 'new',
        submissionDate: c.receivedAt,
      },
    });

    // Create synthetic processing job
    await prisma.processingJob.create({
      data: {
        type: 'field_extraction',
        related: `[DEV-SYNTHETIC] ${candidate.name}`,
        status: 'succeeded',
        attempts: 1,
        startedAt: c.receivedAt,
        completedAt: new Date(c.receivedAt.getTime() + 12000),
      },
    });

    // Create synthetic audit log
    await prisma.auditLog.create({
      data: {
        actor: 'dev-seed',
        action: 'Candidate ingested from email',
        related: candidate.id,
        result: 'success',
        requestId: `SEED-REQ-${candidate.id.slice(0, 8)}`,
        entity: 'Candidate',
        entityId: candidate.id,
      },
    });
  }

  console.log(`Successfully seeded ${syntheticCandidates.length} synthetic candidate records.`);
}

main()
  .catch((e) => {
    console.error('Seed script error:', e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });


import { zodResolver } from "@hookform/resolvers/zod";
import { useEffect } from "react";
import { useForm } from "react-hook-form";
import { z } from "zod";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useUpdateCandidate } from "@/hooks/use-api";
import type { Candidate, CandidatePatch, InterviewStatus } from "@/lib/api/types";

const opt = (max: number) => z.string().trim().max(max);
const schema = z.object({
  name: opt(120),
  phone: opt(30).refine((v) => !v || /^[+\d][\d\s\-()]{6,}$/.test(v), "Enter a valid phone number"),
  college: opt(160),
  degree: opt(80),
  graduationYear: z.string().trim().refine((v) => !v || (/^\d{4}$/.test(v) && +v >= 1960 && +v <= 2035), "Year must be 1960–2035"),
  appliedRole: opt(80),
  interviewStatus: z.enum(["not_scheduled", "scheduled", "completed", "no_show"]),
});
type Form = z.infer<typeof schema>;

const toForm = (c: Candidate): Form => ({
  name: c.name ?? "", phone: c.phone ?? "", college: c.college ?? "", degree: c.degree ?? "",
  graduationYear: c.graduationYear?.toString() ?? "", appliedRole: c.appliedRole ?? "", interviewStatus: c.interviewStatus,
});

export function CandidateEditDialog({ candidate, onClose }: { candidate: Candidate | null; onClose: () => void }) {
  const mut = useUpdateCandidate();
  const form = useForm<Form>({ resolver: zodResolver(schema) });
  useEffect(() => { if (candidate) form.reset(toForm(candidate)); }, [candidate]); // eslint-disable-line react-hooks/exhaustive-deps

  const submit = form.handleSubmit((v) => {
    if (!candidate) return;
    const n = (s: string) => (s ? s : null);
    const patch: CandidatePatch = {
      name: n(v.name), phone: n(v.phone), college: n(v.college), degree: n(v.degree),
      graduationYear: v.graduationYear ? +v.graduationYear : null, appliedRole: n(v.appliedRole),
      interviewStatus: v.interviewStatus as InterviewStatus,
    };
    mut.mutate({ id: candidate.id, patch }, { onSuccess: onClose });
  });

  const fields: [keyof Form, string][] = [["name", "Full name"], ["phone", "Phone"], ["college", "College"], ["degree", "Degree"], ["graduationYear", "Graduation year"], ["appliedRole", "Applied role"]];
  return (
    <Dialog open={!!candidate} onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="sm:max-w-lg">
        <DialogHeader>
          <DialogTitle>Edit {candidate?.id}</DialogTitle>
          <DialogDescription>Email and skills come from the source email and CV and can't be edited here. Leave a field blank if unknown.</DialogDescription>
        </DialogHeader>
        <form onSubmit={submit} className="grid gap-3 sm:grid-cols-2" noValidate>
          {fields.map(([k, label]) => (
            <div key={k} className="space-y-1.5">
              <Label htmlFor={k}>{label}</Label>
              <Input id={k} {...form.register(k)} aria-invalid={!!form.formState.errors[k]} />
              {form.formState.errors[k] && <p className="text-xs text-destructive">{form.formState.errors[k]?.message}</p>}
            </div>
          ))}
          <div className="space-y-1.5">
            <Label>Interview status</Label>
            <Select value={form.watch("interviewStatus")} onValueChange={(v) => form.setValue("interviewStatus", v as Form["interviewStatus"])}>
              <SelectTrigger><SelectValue /></SelectTrigger>
              <SelectContent>
                <SelectItem value="not_scheduled">Not scheduled</SelectItem>
                <SelectItem value="scheduled">Scheduled</SelectItem>
                <SelectItem value="completed">Completed</SelectItem>
                <SelectItem value="no_show">No show</SelectItem>
              </SelectContent>
            </Select>
          </div>
          <DialogFooter className="sm:col-span-2">
            <Button type="button" variant="outline" onClick={onClose}>Cancel</Button>
            <Button type="submit" disabled={mut.isPending}>{mut.isPending ? "Saving…" : "Save changes"}</Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}

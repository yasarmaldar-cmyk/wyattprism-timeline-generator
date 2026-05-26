import { NextRequest, NextResponse } from "next/server";
import { timelineXlsxBuffer } from "@/lib/excel";
import { computeTimeline } from "@/lib/engine";
import type { ProjectInputs } from "@/lib/types";

export const runtime = "nodejs";

function safeFilename(s: string) {
  return s.replace(/[^A-Za-z0-9._-]/g, "_");
}

type SendWyattprismBody = ProjectInputs & {
  wp_project_id?: string;
  wp_project_code?: string;
};

export async function POST(req: NextRequest) {
  try {
    const inputs = (await req.json()) as SendWyattprismBody;
    const url = new URL(req.url);
    const mode = url.searchParams.get("send");

    // Mode 1 — preview (returns JSON of computed dates)
    if (url.searchParams.get("preview") === "1") {
      const result = computeTimeline(inputs);
      return NextResponse.json({
        anchors: {
          kick_off: result.anchors.kick_off.toISOString(),
          closure: result.anchors.closure.toISOString(),
          board_meeting: result.anchors.board_meeting?.toISOString() ?? null,
          agm: result.anchors.agm?.toISOString() ?? null,
        },
        tasks: result.tasks.map((t) => ({
          id: t.id,
          name: t.name,
          phase: t.phase,
          responsibility: t.responsibility,
          startDate: t.startDate.toISOString(),
          endDate: t.endDate.toISOString(),
          durationDays:
            Math.round((t.endDate.getTime() - t.startDate.getTime()) / 86400000) + 1,
        })),
      });
    }

    // Mode 2 — send to Wyattprism shell (server-side POST keeps the callback
    // key out of the browser).
    if (mode === "wyattprism") {
      const shellUrl = process.env.SHELL_URL;
      const callbackKey = process.env.SHELL_CALLBACK_KEY;
      if (!shellUrl || !callbackKey) {
        return NextResponse.json(
          {
            ok: false,
            error:
              "Timeline Generator is not configured to talk to a Wyattprism shell. Set SHELL_URL and SHELL_CALLBACK_KEY env vars.",
          },
          { status: 500 }
        );
      }
      if (!inputs.wp_project_id) {
        return NextResponse.json(
          { ok: false, error: "Missing wp_project_id." },
          { status: 400 }
        );
      }

      const result = computeTimeline(inputs);
      const payload = {
        wp_project_id: inputs.wp_project_id,
        wp_project_code: inputs.wp_project_code,
        saved_at: new Date().toISOString(),
        client_name: inputs.clientName,
        report_type: inputs.reportType,
        reporting_period: inputs.reportingPeriod,
        anchors: {
          kick_off: result.anchors.kick_off.toISOString(),
          closure: result.anchors.closure.toISOString(),
          board_meeting: result.anchors.board_meeting?.toISOString() ?? null,
          agm: result.anchors.agm?.toISOString() ?? null,
        },
        tasks: result.tasks.map((t) => ({
          name: t.name,
          phase: t.phase,
          plannedStart: t.startDate.toISOString(),
          plannedEnd: t.endDate.toISOString(),
          responsibility: t.responsibility,
        })),
      };

      const callbackUrl = `${shellUrl.replace(/\/$/, "")}/api/timeline-callback?key=${encodeURIComponent(callbackKey)}`;
      const cbRes = await fetch(callbackUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const cbJson: { ok?: boolean; error?: string; taskCount?: number } = await cbRes
        .json()
        .catch(() => ({}));
      if (!cbRes.ok || !cbJson.ok) {
        return NextResponse.json(
          {
            ok: false,
            error: cbJson.error ?? `Shell returned HTTP ${cbRes.status}`,
          },
          { status: 502 }
        );
      }
      return NextResponse.json({ ok: true, taskCount: cbJson.taskCount });
    }

    // Mode 3 (default) — Excel download
    const buf = await timelineXlsxBuffer(inputs);
    const body = new Uint8Array(buf);
    const fname = `${safeFilename(inputs.clientName)}_${inputs.reportType}_${inputs.reportingPeriod || "timeline"}.xlsx`;
    return new NextResponse(body, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${fname}"`,
      },
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : "Unknown error";
    return NextResponse.json({ error: msg }, { status: 400 });
  }
}

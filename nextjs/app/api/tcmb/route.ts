import { NextRequest, NextResponse } from "next/server";

export async function GET(req: NextRequest) {
  const date = req.nextUrl.searchParams.get("date");
  if (!date) return NextResponse.json({ error: "date required" }, { status: 400 });

  const url = `https://www.tcmb.gov.tr/kurlar/${date.slice(0, 7).replace("-", "")}/${date.replace(/-/g, "")}.xml`;
  try {
    const res = await fetch(url, { next: { revalidate: 3600 } });
    if (!res.ok) return NextResponse.json({ error: "not found" }, { status: 404 });
    const xml = await res.text();
    return new NextResponse(xml, {
      headers: { "Content-Type": "text/xml; charset=utf-8" },
    });
  } catch {
    return NextResponse.json({ error: "fetch failed" }, { status: 500 });
  }
}

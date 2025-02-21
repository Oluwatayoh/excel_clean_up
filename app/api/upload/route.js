import { NextResponse } from "next/server";
import { processExcel } from "@/lib/processExcel";

export async function POST(request) {
  try {
    const formData = await request.formData();
    const file = formData.get("excelFile");

    if (!file) {
      return NextResponse.json({ message: "No file uploaded." }, { status: 400 });
    }

    // Validate file type (ensure it's an Excel file)
    if (!file.name.endsWith(".xlsx")) {
      return NextResponse.json({ message: "Invalid file type. Only .xlsx files are allowed." }, { status: 400 });
    }

    // Convert file to Buffer
    const buffer = Buffer.from(await file.arrayBuffer());

    // Process the Excel file (processExcel now works with Buffers)
    const processedBuffer = await processExcel(buffer);

    // Convert processed Buffer into a Response for download
    return new NextResponse(processedBuffer, {
      headers: {
        "Content-Disposition": `attachment; filename="output.xlsx"`,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    });
  } catch (error) {
    console.error("Error processing file:", error);
    return NextResponse.json({ message: "An error occurred while processing the file." }, { status: 500 });
  }
}

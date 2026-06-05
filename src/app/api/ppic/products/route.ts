// app/api/ppic/products/route.ts
import { NextResponse } from "next/server";
import { getPool } from "@/lib/config";
import sql from "mssql";

export async function GET() {
  try {
    const pool = await getPool();

    // Ambil daftar produk unik dari BOM yang aktif
    const result = await pool.request().query(`
      SELECT DISTINCT 
        hd.ItemID as kode,
        ISNULL(ghd.ItemName, '') as nama,
        hd.TransID
      FROM taPackingHD hd
      INNER JOIN taGoods ghd ON hd.ItemID = ghd.ItemID
      WHERE EXISTS (
        SELECT 1 FROM taPackingDT dt 
        WHERE dt.TransID = hd.TransID
      )
      ORDER BY hd.ItemID
    `);

    return NextResponse.json({
      success: true,
      data: result.recordset,
    });
  } catch (error) {
    console.error("Error fetching products:", error);
    return NextResponse.json(
      { success: false, error: "Failed to fetch products" },
      { status: 500 },
    );
  }
}

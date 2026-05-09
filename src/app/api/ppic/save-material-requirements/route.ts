import { NextResponse } from "next/server";
import sql from "mssql";
import { getPool } from "@/lib/config";

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const {
      calculation_id,
      calculation_name, // TAMBAHKAN FIELD INI
      user_id,
      po_list,
      total_po,
      material_data,
      total_materials,
      total_kebutuhan,
      total_kekurangan,
      material_aman,
      material_kurang,
      material_habis,
      stock_date,
      notes,
    } = body;

    // Validasi
    if (!calculation_id || !material_data) {
      return NextResponse.json(
        { success: false, error: "Missing required fields" },
        { status: 400 },
      );
    }

    // Koneksi ke database
    const pool = await getPool();

    // Cek apakah calculation_id sudah ada
    const checkResult = await pool
      .request()
      .input("calculation_id", sql.VarChar, calculation_id)
      .query(
        `SELECT id FROM material_requirements_history WHERE calculation_id = @calculation_id`,
      );

    let result;

    if (checkResult.recordset.length > 0) {
      // Update existing record
      result = await pool
        .request()
        .input("calculation_id", sql.VarChar, calculation_id)
        .input(
          "calculation_name",
          sql.NVarChar,
          calculation_name || calculation_id,
        ) // TAMBAHKAN
        .input("user_id", sql.VarChar, user_id || "system")
        .input("po_list", sql.NVarChar, JSON.stringify(po_list))
        .input("total_po", sql.Int, total_po)
        .input("material_data", sql.NVarChar, JSON.stringify(material_data))
        .input("total_materials", sql.Int, total_materials)
        .input("total_kebutuhan", sql.Decimal(15, 2), total_kebutuhan || 0)
        .input("total_kekurangan", sql.Decimal(15, 2), total_kekurangan || 0)
        .input("material_aman", sql.Int, material_aman || 0)
        .input("material_kurang", sql.Int, material_kurang || 0)
        .input("material_habis", sql.Int, material_habis || 0)
        .input("stock_date", sql.Date, stock_date || new Date())
        .input("notes", sql.NVarChar, notes || "").query(`
          UPDATE material_requirements_history 
          SET 
            calculation_name = @calculation_name,
            user_id = @user_id,
            po_list = @po_list,
            total_po = @total_po,
            material_data = @material_data,
            total_materials = @total_materials,
            total_kebutuhan = @total_kebutuhan,
            total_kekurangan = @total_kekurangan,
            material_aman = @material_aman,
            material_kurang = @material_kurang,
            material_habis = @material_habis,
            stock_date = @stock_date,
            notes = @notes,
            updated_at = GETDATE()
          WHERE calculation_id = @calculation_id
        `);

      return NextResponse.json({
        success: true,
        message: "Material requirements updated successfully",
        id: null,
        calculation_id: calculation_id,
      });
    } else {
      // Insert new record
      result = await pool
        .request()
        .input("calculation_id", sql.VarChar, calculation_id)
        .input(
          "calculation_name",
          sql.NVarChar,
          calculation_name || calculation_id,
        ) // TAMBAHKAN
        .input("user_id", sql.VarChar, user_id || "system")
        .input("po_list", sql.NVarChar, JSON.stringify(po_list))
        .input("total_po", sql.Int, total_po)
        .input("material_data", sql.NVarChar, JSON.stringify(material_data))
        .input("total_materials", sql.Int, total_materials)
        .input("total_kebutuhan", sql.Decimal(15, 2), total_kebutuhan || 0)
        .input("total_kekurangan", sql.Decimal(15, 2), total_kekurangan || 0)
        .input("material_aman", sql.Int, material_aman || 0)
        .input("material_kurang", sql.Int, material_kurang || 0)
        .input("material_habis", sql.Int, material_habis || 0)
        .input("stock_date", sql.Date, stock_date || new Date())
        .input("notes", sql.NVarChar, notes || "").query(`
          INSERT INTO material_requirements_history (
            calculation_id,
            calculation_name,
            calculation_date,
            user_id,
            po_list,
            total_po,
            material_data,
            total_materials,
            total_kebutuhan,
            total_kekurangan,
            material_aman,
            material_kurang,
            material_habis,
            stock_date,
            notes
          ) VALUES (
            @calculation_id,
            @calculation_name,
            GETDATE(),
            @user_id,
            @po_list,
            @total_po,
            @material_data,
            @total_materials,
            @total_kebutuhan,
            @total_kekurangan,
            @material_aman,
            @material_kurang,
            @material_habis,
            @stock_date,
            @notes
          );
          SELECT SCOPE_IDENTITY() AS id;
        `);

      return NextResponse.json({
        success: true,
        message: "Material requirements saved successfully",
        id: result?.recordset?.[0]?.id || null,
        calculation_id: calculation_id,
      });
    }
  } catch (error) {
    console.error("Error saving material requirements:", error);
    return NextResponse.json(
      {
        success: false,
        error:
          "Failed to save material requirements: " + (error as Error).message,
      },
      { status: 500 },
    );
  }
}

// GET endpoint untuk mengambil data yang sudah disimpan
export async function GET(request: Request) {
  try {
    const { searchParams } = new URL(request.url);
    const calculation_id = searchParams.get("calculation_id");
    const limit = parseInt(searchParams.get("limit") || "50");
    const offset = parseInt(searchParams.get("offset") || "0");

    const pool = await getPool();

    let query = `
      SELECT 
        id, 
        calculation_id, 
        calculation_name,
        calculation_date,
        user_id,
        po_list,
        total_po,
        material_data,
        total_materials,
        total_kebutuhan,
        total_kekurangan,
        material_aman,
        material_kurang,
        material_habis,
        stock_date,
        notes,
        created_at,
        updated_at
      FROM material_requirements_history
    `;

    const request_db = pool.request();

    if (calculation_id) {
      query += ` WHERE calculation_id = @calculation_id`;
      request_db.input("calculation_id", sql.VarChar, calculation_id);
    }

    query += ` ORDER BY calculation_date DESC OFFSET @offset ROWS FETCH NEXT @limit ROWS ONLY`;
    request_db.input("offset", sql.Int, offset);
    request_db.input("limit", sql.Int, limit);

    const result = await request_db.query(query);

    // Parse JSON strings back to objects
    const records = result.recordset.map((record) => ({
      ...record,
      po_list: record.po_list ? JSON.parse(record.po_list) : [],
      material_data: record.material_data
        ? JSON.parse(record.material_data)
        : [],
    }));

    return NextResponse.json({
      success: true,
      data: records,
      total: records.length,
    });
  } catch (error) {
    console.error("Error fetching material requirements:", error);
    return NextResponse.json(
      {
        success: false,
        error: "Failed to fetch data: " + (error as Error).message,
      },
      { status: 500 },
    );
  }
}

// DELETE endpoint untuk menghapus data
export async function DELETE(request: Request) {
  try {
    const { searchParams } = new URL(request.url);
    const calculation_id = searchParams.get("calculation_id");

    if (!calculation_id) {
      return NextResponse.json(
        { success: false, error: "calculation_id is required" },
        { status: 400 },
      );
    }

    const pool = await getPool();

    const result = await pool
      .request()
      .input("calculation_id", sql.VarChar, calculation_id)
      .query(
        `DELETE FROM material_requirements_history WHERE calculation_id = @calculation_id`,
      );

    return NextResponse.json({
      success: true,
      message: "Data deleted successfully",
      affectedRows: result.rowsAffected?.[0] || 0,
    });
  } catch (error) {
    console.error("Error deleting material requirements:", error);
    return NextResponse.json(
      { success: false, error: "Failed to delete data" },
      { status: 500 },
    );
  }
}

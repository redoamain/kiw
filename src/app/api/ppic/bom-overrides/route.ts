// app/api/ppic/bom-overrides/route.ts
import { NextRequest, NextResponse } from "next/server";
import sql from "mssql";
import { getPool } from "@/lib/config";

const normalizeItemId = (id: string): string => {
  if (!id) return "";
  return id.trim().toUpperCase();
};

// GET: Ambil semua BOM overrides
export async function GET() {
  try {
    const pool = await getPool();
    const result = await pool.request().query(`
      SELECT 
        id,
        original_item_id,
        replacement_item_id,
        replacement_item_name,
        replacement_item_name2,
        is_active,
        target_kode_barang,
        target_kode_barangs,
        created_by,
        created_at,
        updated_at
      FROM bom_overrides
      ORDER BY created_at DESC
    `);

    return NextResponse.json({
      success: true,
      data: result.recordset,
    });
  } catch (error) {
    console.error("Error fetching BOM overrides:", error);
    return NextResponse.json(
      { success: false, error: "Failed to fetch BOM overrides" },
      { status: 500 },
    );
  }
}

// POST: Tambah atau update BOM override
export async function POST(request: NextRequest) {
  let transaction: sql.Transaction | null = null;

  try {
    const body = await request.json();
    const {
      id,
      originalItemId,
      replacementItemId,
      replacementItemName,
      replacementItemName2,
      isActive,
      targetKodeBarang,
      targetKodeBarangs,
      createdBy = "system",
    } = body;

    if (!originalItemId || !replacementItemId) {
      return NextResponse.json(
        {
          success: false,
          error: "Original item and replacement item are required",
        },
        { status: 400 },
      );
    }

    const pool = await getPool();
    transaction = new sql.Transaction(pool);
    await transaction.begin();

    let resultData: any = null;

    if (id) {
      // Update existing override
      const updateRequest = transaction.request();
      await updateRequest
        .input("originalItemId", sql.NVarChar, originalItemId)
        .input("replacementItemId", sql.NVarChar, replacementItemId)
        .input(
          "replacementItemName",
          sql.NVarChar,
          replacementItemName || replacementItemId,
        )
        .input(
          "replacementItemName2",
          sql.NVarChar,
          replacementItemName2 || replacementItemId,
        )
        .input("isActive", sql.Bit, isActive !== undefined ? isActive : true)
        .input("targetKodeBarang", sql.NVarChar, targetKodeBarang || null)
        .input(
          "targetKodeBarangs",
          sql.NVarChar,
          targetKodeBarangs ? JSON.stringify(targetKodeBarangs) : null,
        )
        .input("id", sql.Int, parseInt(String(id))).query(`
          UPDATE bom_overrides 
          SET 
            original_item_id = @originalItemId,
            replacement_item_id = @replacementItemId,
            replacement_item_name = @replacementItemName,
            replacement_item_name2 = @replacementItemName2,
            is_active = @isActive,
            target_kode_barang = @targetKodeBarang,
            target_kode_barangs = @targetKodeBarangs,
            updated_at = GETDATE()
          WHERE id = @id
        `);

      const getRequest = transaction.request();
      const getResult = await getRequest
        .input("id", sql.Int, parseInt(String(id)))
        .query(`SELECT * FROM bom_overrides WHERE id = @id`);

      resultData = getResult.recordset[0];
      await transaction.commit();
    } else {
      // Insert new override
      const insertRequest = transaction.request();
      const result = await insertRequest
        .input("originalItemId", sql.NVarChar, originalItemId)
        .input("replacementItemId", sql.NVarChar, replacementItemId)
        .input(
          "replacementItemName",
          sql.NVarChar,
          replacementItemName || replacementItemId,
        )
        .input(
          "replacementItemName2",
          sql.NVarChar,
          replacementItemName2 || replacementItemId,
        )
        .input("isActive", sql.Bit, isActive !== undefined ? isActive : true)
        .input("targetKodeBarang", sql.NVarChar, targetKodeBarang || null)
        .input(
          "targetKodeBarangs",
          sql.NVarChar,
          targetKodeBarangs ? JSON.stringify(targetKodeBarangs) : null,
        )
        .input("createdBy", sql.NVarChar, createdBy).query(`
          INSERT INTO bom_overrides (
            original_item_id,
            replacement_item_id,
            replacement_item_name,
            replacement_item_name2,
            is_active,
            target_kode_barang,
            target_kode_barangs,
            created_by,
            created_at,
            updated_at
          ) VALUES (
            @originalItemId,
            @replacementItemId,
            @replacementItemName,
            @replacementItemName2,
            @isActive,
            @targetKodeBarang,
            @targetKodeBarangs,
            @createdBy,
            GETDATE(),
            GETDATE()
          );
          SELECT SCOPE_IDENTITY() as id;
        `);

      const insertedId = result.recordset[0]?.id;

      const getRequest = transaction.request();
      const getResult = await getRequest
        .input("id", sql.Int, insertedId)
        .query(`SELECT * FROM bom_overrides WHERE id = @id`);

      resultData = getResult.recordset[0];
      await transaction.commit();
    }

    return NextResponse.json({
      success: true,
      data: resultData,
    });
  } catch (error) {
    if (transaction) {
      try {
        await transaction.rollback();
      } catch (rollbackError) {
        console.error("Rollback error:", rollbackError);
      }
    }
    console.error("Error saving BOM override:", error);
    return NextResponse.json(
      { success: false, error: "Failed to save BOM override" },
      { status: 500 },
    );
  }
}

// DELETE dan PATCH sama seperti sebelumnya...
// DELETE: Hapus BOM override
export async function DELETE(request: NextRequest) {
  let transaction: sql.Transaction | null = null;

  try {
    const { searchParams } = new URL(request.url);
    const id = searchParams.get("id");

    if (!id) {
      return NextResponse.json(
        { success: false, error: "ID is required" },
        { status: 400 },
      );
    }

    const pool = await getPool();
    transaction = new sql.Transaction(pool);
    await transaction.begin();

    const deleteRequest = transaction.request();
    await deleteRequest
      .input("id", sql.Int, parseInt(String(id)))
      .query(`DELETE FROM bom_overrides WHERE id = @id`);

    await transaction.commit();

    return NextResponse.json({
      success: true,
      message: "BOM override deleted successfully",
    });
  } catch (error) {
    if (transaction) {
      try {
        await transaction.rollback();
      } catch (rollbackError) {
        console.error("Rollback error:", rollbackError);
      }
    }
    console.error("Error deleting BOM override:", error);
    return NextResponse.json(
      { success: false, error: "Failed to delete BOM override" },
      { status: 500 },
    );
  }
}

// PATCH: Toggle active status
export async function PATCH(request: NextRequest) {
  let transaction: sql.Transaction | null = null;

  try {
    const body = await request.json();
    const { id, isActive } = body;

    if (!id) {
      return NextResponse.json(
        { success: false, error: "ID is required" },
        { status: 400 },
      );
    }

    const pool = await getPool();
    transaction = new sql.Transaction(pool);
    await transaction.begin();

    const updateRequest = transaction.request();
    await updateRequest
      .input("isActive", sql.Bit, isActive)
      .input("id", sql.Int, parseInt(String(id))).query(`
        UPDATE bom_overrides 
        SET is_active = @isActive, updated_at = GETDATE()
        WHERE id = @id
      `);

    // Get the updated record
    const getRequest = transaction.request();
    const getResult = await getRequest
      .input("id", sql.Int, parseInt(String(id)))
      .query(`SELECT * FROM bom_overrides WHERE id = @id`);

    await transaction.commit();

    return NextResponse.json({
      success: true,
      data: getResult.recordset[0],
    });
  } catch (error) {
    if (transaction) {
      try {
        await transaction.rollback();
      } catch (rollbackError) {
        console.error("Rollback error:", rollbackError);
      }
    }
    console.error("Error toggling BOM override:", error);
    return NextResponse.json(
      { success: false, error: "Failed to toggle BOM override" },
      { status: 500 },
    );
  }
}

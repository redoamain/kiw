import { NextResponse } from "next/server";
import { getPool } from "@/lib/config";
import sql from "mssql";

const CACHE_DURATION = 5 * 60 * 1000; // 5 menit
const cache = new Map();




export async function GET(request: Request) {
  const url = new URL(request.url);
  const itemid = url.searchParams.get("itemid");

  try {
    const pool = await getPool();
    
    let query = `
      SELECT 
        hd.TransID,
        hd.ItemID AS itemidHD,
        ISNULL(ghd.ItemName, '') AS itemnamehd,
        ISNULL(ghd.ItemName2, '') AS itemnamehd2,
        dt.ItemID,
        ISNULL(gdt.ItemName, '') AS ItemName,
        ISNULL(gdt.ItemName2, '') AS ItemName2,
        ISNULL(dt.BahanQty, 0) AS BahanQty,
        ISNULL(dt.BahanPackSatuan, '') AS BahanPackSatuan,
        ISNULL(gdt.Spec, '') AS Spec,
        ISNULL(gdt.warnac, '') AS warna,
        ISNULL(gdt.bahan, '') AS bahan
      FROM taPackingHD hd
      INNER JOIN taPackingDT dt ON hd.TransID = dt.TransID
      INNER JOIN taGoods ghd ON hd.ItemID = ghd.ItemID
      LEFT JOIN taGoods gdt ON dt.ItemID = gdt.ItemID
    `;
    
    if (itemid && itemid.trim() !== "") {
      query += ` WHERE hd.ItemID = '${itemid}'`;
    }
    
    query += ` ORDER BY hd.ItemID, dt.ItemID`;
    
    const result = await pool.request().query(query);
    
    return NextResponse.json(result.recordset);
    
  } catch (error: any) {
    console.error("Error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}

// app/api/bom/route.ts - Tambahkan method PUT
export async function PUT(request: Request) {
  try {
    const body = await request.json();
    const { TransID, ItemID, newItemID, BahanQty, BahanPackSatuan, action } =
      body;

    console.log("PUT request received:", {
      TransID,
      ItemID,
      newItemID,
      BahanQty,
      BahanPackSatuan,
      action,
    });

    if (!TransID || !ItemID) {
      return NextResponse.json(
        { error: "TransID dan ItemID diperlukan" },
        { status: 400 },
      );
    }

    const pool = await getPool();

    if (action === "delete") {
      // Delete komponen
      const result = await pool
        .request()
        .input("TransID", sql.VarChar(7), TransID)
        .input("ItemID", sql.VarChar(25), ItemID).query(`
          DELETE FROM taPackingDT 
          WHERE TransID = @TransID AND ItemID = @ItemID
        `);

      if (result.rowsAffected[0] === 0) {
        return NextResponse.json(
          { error: "Komponen tidak ditemukan" },
          { status: 404 },
        );
      }

      clearCache();

      return NextResponse.json({
        success: true,
        message: "Komponen berhasil dihapus",
        TransID,
        ItemID,
      });
    } else if (action === "update") {
      // Update komponen
      const updateItemID = newItemID || ItemID;

      // Cek apakah perlu update ItemID (rename)
      if (updateItemID !== ItemID) {
        // Delete old and insert new
        const transaction = new sql.Transaction(pool);
        await transaction.begin();

        try {
          // Delete old
          await transaction
            .request()
            .input("TransID", sql.VarChar(7), TransID)
            .input("ItemID", sql.VarChar(25), ItemID)
            .query(
              `DELETE FROM taPackingDT WHERE TransID = @TransID AND ItemID = @ItemID`,
            );

          // Insert new
          await transaction
            .request()
            .input("TransID", sql.VarChar(7), TransID)
            .input("ItemID", sql.VarChar(25), updateItemID)
            .input("BahanQty", sql.Decimal(18, 5), BahanQty)
            .input("BahanPackSatuan", sql.VarChar(10), BahanPackSatuan || null)
            .query(`
              INSERT INTO taPackingDT (TransID, ItemID, BahanQty, BahanPackSatuan)
              VALUES (@TransID, @ItemID, @BahanQty, @BahanPackSatuan)
            `);

          await transaction.commit();
        } catch (error) {
          await transaction.rollback();
          throw error;
        }
      } else {
        // Just update quantity and satuan
        const result = await pool
          .request()
          .input("TransID", sql.VarChar(7), TransID)
          .input("ItemID", sql.VarChar(25), ItemID)
          .input("BahanQty", sql.Decimal(18, 5), BahanQty)
          .input("BahanPackSatuan", sql.VarChar(10), BahanPackSatuan || null)
          .query(`
            UPDATE taPackingDT 
            SET BahanQty = @BahanQty, BahanPackSatuan = @BahanPackSatuan
            WHERE TransID = @TransID AND ItemID = @ItemID
          `);

        if (result.rowsAffected[0] === 0) {
          return NextResponse.json(
            { error: "Komponen tidak ditemukan" },
            { status: 404 },
          );
        }
      }

      clearCache();

      return NextResponse.json({
        success: true,
        message: "Komponen berhasil diupdate",
        TransID,
        ItemID: updateItemID,
        BahanQty,
        BahanPackSatuan,
      });
    } else {
      return NextResponse.json(
        { error: "Action tidak valid. Gunakan 'update' atau 'delete'" },
        { status: 400 },
      );
    }
  } catch (error) {
    console.error("Error updating BOM:", error);
    return NextResponse.json(
      { error: "Gagal mengupdate BOM", details: String(error) },
      { status: 500 },
    );
  }
}
// PATCH untuk bulk update atau add component
export async function PATCH(request: Request) {
  try {
    const body = await request.json();
    const { action, TransID, components } = body;

    console.log("PATCH request:", { action, TransID, components });

    const pool = await getPool();

    if (action === "add" && components && Array.isArray(components)) {
      // Add multiple components
      const transaction = new sql.Transaction(pool);
      await transaction.begin();

      try {
        for (const comp of components) {
          const { ItemID, BahanQty, BahanPackSatuan } = comp;

          // Cek apakah komponen sudah ada
          const checkResult = await transaction
            .request()
            .input("TransID", sql.VarChar(7), TransID)
            .input("ItemID", sql.VarChar(25), ItemID).query(`
              SELECT COUNT(*) as count FROM taPackingDT 
              WHERE TransID = @TransID AND ItemID = @ItemID
            `);

          if (checkResult.recordset[0].count > 0) {
            await transaction.rollback();
            return NextResponse.json(
              {
                success: false,
                message: `Komponen ${ItemID} sudah ada dalam BOM ini`,
              },
              { status: 400 },
            );
          }

          await transaction
            .request()
            .input("TransID", sql.VarChar(7), TransID)
            .input("ItemID", sql.VarChar(25), ItemID)
            .input("BahanQty", sql.Decimal(18, 5), BahanQty || 0)
            .input("BahanPackSatuan", sql.VarChar(10), BahanPackSatuan || null)
            .query(`
              INSERT INTO taPackingDT (TransID, ItemID, BahanQty, BahanPackSatuan)
              VALUES (@TransID, @ItemID, @BahanQty, @BahanPackSatuan)
            `);
        }

        await transaction.commit();

        // Clear cache after update
        clearCache();

        return NextResponse.json({
          success: true,
          message: `${components.length} komponen berhasil ditambahkan`,
          TransID,
        });
      } catch (error) {
        await transaction.rollback();
        throw error;
      }
    } else {
      return NextResponse.json(
        { error: "Action tidak valid. Gunakan 'add'" },
        { status: 400 },
      );
    }
  } catch (error) {
    console.error("Error in BOM operation:", error);
    return NextResponse.json(
      { error: "Gagal memproses BOM", details: String(error) },
      { status: 500 },
    );
  }
}
// Function to clear cache
function clearCache() {
  cache.clear();
  console.log("[CACHE CLEARED] All cache entries removed after update");
}

// Clear expired cache entries
if (typeof setInterval !== "undefined") {
  setInterval(
    () => {
      const now = Date.now();
      let deletedCount = 0;
      cache.forEach((value, key) => {
        if (now - value.timestamp > CACHE_DURATION) {
          cache.delete(key);
          deletedCount++;
        }
      });
      if (deletedCount > 0) {
        console.log(
          `[CACHE CLEANED] Removed ${deletedCount} expired cache entries`,
        );
      }
    },
    60 * 60 * 1000,
  );
}

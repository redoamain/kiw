import { NextResponse } from "next/server";
import { getPool } from "@/lib/config";
import sql from "mssql";

const CACHE_DURATION = 5 * 60 * 1000; // 5 menit
const cache = new Map();

export async function GET(request: Request) {
  const url = new URL(request.url);
  const itemid = url.searchParams.get("itemid");
  const searchType = url.searchParams.get("searchType") || "itemid";
  const page = parseInt(url.searchParams.get("page") || "1");
  const limit = parseInt(url.searchParams.get("limit") || "50");
  const fetchAll = url.searchParams.get("fetchAll") === "true";

  const cacheKey = `bom-${searchType}-${itemid || "all"}-page${page}-limit${limit}-all${fetchAll}`;
  const cached = cache.get(cacheKey);

  if (cached && Date.now() - cached.timestamp < CACHE_DURATION) {
    console.log(`[CACHE HIT] BOM for ${itemid || "all"}`);
    return NextResponse.json(cached.data);
  }

  console.log(
    `[CACHE MISS] Fetching BOM for ${itemid || "all"}, fetchAll: ${fetchAll}`,
  );

  try {
    const pool = await getPool();
    let result;
    let totalCount = 0;
    let totalPages = 1;

    if (fetchAll) {
      // Ambil semua data tanpa pagination
      console.log("Fetching ALL BOM data...");

      if (searchType === "itemname" && itemid && itemid.trim() !== "") {
        result = await pool
          .request()
          .input("itemid", sql.NVarChar, `%${itemid}%`).query(`
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
              ISNULL(gdt.Mark, '') AS Departemen,
              ISNULL(gdt.Spec, '') AS Spec,
              ISNULL(gdt.warnac, '') AS warna,
              ISNULL(gdt.bahan, '') AS bahan,
              ISNULL(got.NamaJenis, '') AS NamaJenis
            FROM taPackingHD hd
            INNER JOIN taPackingDT dt ON hd.TransID = dt.TransID
            INNER JOIN taGoods ghd ON hd.ItemID = ghd.ItemID
            INNER JOIN taGoods gdt ON dt.ItemID = gdt.ItemID
            INNER JOIN taKindofGoods got ON gdt.KodeJenis = got.KodeJenis
            WHERE ghd.ItemName LIKE @itemid
            ORDER BY hd.ItemID, dt.ItemID
          `);
      } else {
        let searchParam = "%";
        if (itemid && itemid.trim() !== "" && itemid !== "%") {
          searchParam = `%${itemid}%`;
        }

        result = await pool
          .request()
          .input("itemid", sql.VarChar(50), searchParam).query(`
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
              ISNULL(gdt.Mark, '') AS Departemen,
              ISNULL(gdt.Spec, '') AS Spec,
              ISNULL(gdt.warnac, '') AS warna,
              ISNULL(gdt.bahan, '') AS bahan,
              ISNULL(got.NamaJenis, '') AS NamaJenis
            FROM taPackingHD hd
            INNER JOIN taPackingDT dt ON hd.TransID = dt.TransID
            INNER JOIN taGoods ghd ON hd.ItemID = ghd.ItemID
            INNER JOIN taGoods gdt ON dt.ItemID = gdt.ItemID
            INNER JOIN taKindofGoods got ON gdt.KodeJenis = got.KodeJenis
            WHERE ghd.ItemID LIKE @itemid OR @itemid = '%'
            ORDER BY hd.ItemID, dt.ItemID
          `);
      }

      totalCount = result.recordset.length;
      console.log(`Fetched ${totalCount} records (all data)`);
    } else {
      // Dengan pagination
      if (searchType === "itemname" && itemid && itemid.trim() !== "") {
        const countResult = await pool
          .request()
          .input("itemid", sql.NVarChar, `%${itemid}%`).query(`
            SELECT COUNT(DISTINCT hd.ItemID) as total
            FROM taPackingHD hd
            INNER JOIN taGoods ghd ON hd.ItemID = ghd.ItemID
            WHERE ghd.ItemName LIKE @itemid
          `);
        totalCount = countResult.recordset[0]?.total || 0;
        totalPages = Math.ceil(totalCount / limit);

        const offset = (page - 1) * limit;
        result = await pool
          .request()
          .input("itemid", sql.NVarChar, `%${itemid}%`)
          .input("offset", sql.Int, offset)
          .input("limit", sql.Int, limit).query(`
            SELECT * FROM (
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
                ISNULL(gdt.Mark, '') AS Departemen,
                ISNULL(gdt.Spec, '') AS Spec,
                ISNULL(gdt.warnac, '') AS warna,
                ISNULL(gdt.bahan, '') AS bahan,
                ISNULL(got.NamaJenis, '') AS NamaJenis,
                ROW_NUMBER() OVER (ORDER BY hd.ItemID, dt.ItemID) as rownum
              FROM taPackingHD hd
              INNER JOIN taPackingDT dt ON hd.TransID = dt.TransID
              INNER JOIN taGoods ghd ON hd.ItemID = ghd.ItemID
              INNER JOIN taGoods gdt ON dt.ItemID = gdt.ItemID
              INNER JOIN taKindofGoods got ON gdt.KodeJenis = got.KodeJenis
              WHERE ghd.ItemName LIKE @itemid
            ) t
            WHERE rownum > @offset AND rownum <= @offset + @limit
            ORDER BY itemidHD, ItemID
          `);
      } else {
        let searchParam = "%";
        if (itemid && itemid.trim() !== "" && itemid !== "%") {
          searchParam = `%${itemid}%`;
        }

        const countResult = await pool
          .request()
          .input("itemid", sql.VarChar(50), searchParam).query(`
            SELECT COUNT(DISTINCT hd.ItemID) as total
            FROM taPackingHD hd
            WHERE hd.ItemID LIKE @itemid OR @itemid = '%'
          `);
        totalCount = countResult.recordset[0]?.total || 0;
        totalPages = Math.ceil(totalCount / limit);

        result = await pool
          .request()
          .input("itemid", sql.VarChar(50), searchParam)
          .input("offset", sql.Int, (page - 1) * limit)
          .input("limit", sql.Int, limit).query(`
            SELECT * FROM (
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
                ISNULL(gdt.Mark, '') AS Departemen,
                ISNULL(gdt.Spec, '') AS Spec,
                ISNULL(gdt.warnac, '') AS warna,
                ISNULL(gdt.bahan, '') AS bahan,
                ISNULL(got.NamaJenis, '') AS NamaJenis,
                ROW_NUMBER() OVER (ORDER BY hd.ItemID, dt.ItemID) as rownum
              FROM taPackingHD hd
              INNER JOIN taPackingDT dt ON hd.TransID = dt.TransID
              INNER JOIN taGoods ghd ON hd.ItemID = ghd.ItemID
              INNER JOIN taGoods gdt ON dt.ItemID = gdt.ItemID
              INNER JOIN taKindofGoods got ON gdt.KodeJenis = got.KodeJenis
              WHERE ghd.ItemID LIKE @itemid OR @itemid = '%'
            ) t
            WHERE rownum > @offset AND rownum <= @offset + @limit
            ORDER BY itemidHD, ItemID
          `);
      }

      console.log(
        `Fetched ${result.recordset.length} records (page ${page} of ${totalPages})`,
      );
    }

    const transformedData = result.recordset.map((row: any) => ({
      TransID: row.TransID,
      itemidHD: row.itemidHD,
      itemnamehd: row.itemnamehd || "",
      itemnamehd2: row.itemnamehd2 || "",
      ItemID: row.ItemID,
      ItemName: row.ItemName || "",
      ItemName2: row.ItemName2 || "",
      BahanQty: row.BahanQty || 0,
      BahanPackSatuan: row.BahanPackSatuan || "",
      Departemen: row.Departemen || "",
      NamaJenis: row.NamaJenis || "",
      Spec: row.Spec || "",
      warna: row.warna || "",
      bahan: row.bahan || "",
    }));

    const response = {
      data: transformedData,
      pagination: fetchAll
        ? null
        : {
            currentPage: page,
            totalPages: totalPages,
            totalItems: totalCount,
            itemsPerPage: limit,
          },
      isAllData: fetchAll,
    };

    cache.set(cacheKey, {
      data: response,
      timestamp: Date.now(),
    });

    return NextResponse.json(response);
  } catch (error) {
    console.error("Error fetching BOM:", error);
    return NextResponse.json(
      { error: "Terjadi kesalahan saat mengambil data BOM" },
      { status: 500 },
    );
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

/* eslint-disable react-hooks/exhaustive-deps */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
// app/production-plan/page.tsx
"use client";
import { useState, useEffect, useMemo, useCallback } from "react";
import axios from "axios";
import * as XLSX from "xlsx";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  Sheet,
  SheetContent,
  SheetDescription,
  SheetFooter,
  SheetHeader,
  SheetTitle,
} from "@/components/ui/sheet";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import { Badge } from "@/components/ui/badge";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import {
  Pagination,
  PaginationContent,
  PaginationItem,
  PaginationLink,
  PaginationNext,
  PaginationPrevious,
} from "@/components/ui/pagination";
import { Progress } from "@/components/ui/progress";
import {
  AlertCircle,
  CheckCircle2,
  Download,
  RefreshCw,
  Search,
  Package,
  Factory,
  Calendar,
  ClipboardList,
  Lock,
  Unlock,
  Eye,
  Loader2,
  Info,
  TreePine,
  Table as TableIcon,
  ChevronUp,
  ChevronDown,
  ChevronsUpDown,
  Database,
  Trash2,
  ChevronRight,
  ChevronLeft,
} from "lucide-react";
// Import itemsWithVariants di bagian atas file
import {
  hasVariant,
  getVariantInfo,
  itemsWithVariants,
} from "./itemsWithVariants";
// ==================== TIPE DATA ====================
interface ProductionOrder {
  No_SPK: string;
  Tanggal_Order: string;
  Nama_PO: string;
  Kode_Barang: string;
  QTY: number;
  isCombined?: boolean;
  combinedItems?: ProductionOrder[];
  originalOrders?: ProductionOrder[];
  Completed?: boolean;
  FinishedDate?: string;
}

interface BomItem {
  ItemID: string;
  ItemName: string;
  ItemName2: string;
  Qty: number;
  Level: number;
  Departemen: string;
  NamaJenis: string;
  ParentItemID?: string;
  children?: BomItem[];
  _uniqueKey?: string;
  _originalIndex?: number;
  _isOverridden?: boolean;
  _originalItemId?: string;
}

interface StockItem {
  itemid: string;
  itemname: string;
  stockAkhir: number; // Untuk SaldoAkhir
  physicalStock: number; // Untuk SaldoAkhirFisik (Stok Wincp)
  committedQty?: number;
  reservedQty?: number;
}

interface ProductionPlan {
  order: ProductionOrder;
  bom?: {
    flat: BomItem[];
    tree: BomItem[];
    combinedBoms?: {
      [kodeBarang: string]: {
        flat: BomItem[];
        tree: BomItem[];
      };
    };
  };
  stock?: StockItem[];
  selected: boolean;
  loadingBom: boolean;
  committed: boolean;
  CommitID?: number;
  stockLastUpdated?: string;
  error?: string;
  // TAMBAHKAN FIELD INI
  materialSummary?: {
    totalNeeded: number;
    totalShortage: number;
    status: "CUKUP" | "KURANG" | "BELUM_DIHITUNG";
    itemCount: number;
  };
}

interface CommittedPO {
  CommitID: number;
  noSPK: string;
  kodeBarang: string;
  namaPO: string;
  qty: number;
  tanggalCommit: string;
  userID: string;
  status: string;
  totalMaterials: number;
  totalQtyReserved: number;
}

interface StockReservation {
  reservationID: number;
  CommitID: number;
  itemID: string;
  itemName: string;
  reservedQty: number;
  reservationDate: string;
  status: string;
  expiryDate: string;
  noSPK: string;
  namaPO?: string;
}
interface BomOverride {
  id?: number;
  originalItemId: string;
  replacementItemId: string;
  replacementItemName: string;
  replacementItemName2: string;
  isActive: boolean;
  targetKodeBarang?: string; // TAMBAHKAN: hanya berlaku untuk produk tertentu
  targetKodeBarangs?: string[]; // ATAU untuk multiple produk
}
// ==================== FUNGSI BANTU ====================
const normalizeItemId = (id: string): string => {
  if (!id) return "";
  return id.trim().toUpperCase();
};

const isINJECTIONDepartment = (departemen?: string): boolean => {
  if (!departemen) return false;
  const deptString = String(departemen).trim().toUpperCase();
  return (
    deptString.includes("INJEKSI-BB") ||
    deptString === "INJEKSI-BB" ||
    deptString.includes("INJEKSI BB") ||
    deptString.includes("INJEKSI_BB") ||
    deptString === "INJEKSIBB"
  );
};

const filterOnlyComponents = (bomItems: BomItem[]): BomItem[] => {
  return bomItems.filter((item) => Number(item.Level) > 0);
};

const buildTreeStructure = (flatBom: BomItem[]): BomItem[] => {
  if (!flatBom || flatBom.length === 0) return [];

  const itemMap = new Map<string, BomItem>();
  const rootItems: BomItem[] = [];

  flatBom.forEach((item) => {
    const treeItem = { ...item, children: [] as BomItem[] };
    itemMap.set(normalizeItemId(item.ItemID), treeItem);
  });

  flatBom.forEach((item) => {
    const normalizedId = normalizeItemId(item.ItemID);
    const treeItem = itemMap.get(normalizedId)!;
    const parentId = item.ParentItemID
      ? normalizeItemId(item.ParentItemID)
      : null;

    if (!parentId || parentId === normalizedId || !itemMap.has(parentId)) {
      rootItems.push(treeItem);
    } else {
      const parent = itemMap.get(parentId);
      if (parent && parent.children) {
        parent.children.push(treeItem);
      }
    }
  });

  return rootItems;
};

// ==================== FUNGSI PERHITUNGAN MATERIAL ====================
const calculateMaterialNeeds = (
  bom: BomItem[],
  productionQty: number,
  stock: StockItem[] = [],
) => {
  if (!bom) return { totalNeeded: 0, totalShortage: 0, items: [] };

  const componentsOnly = filterOnlyComponents(bom);

  let totalNeeded = 0;
  let totalShortage = 0;
  const items = componentsOnly.map((item) => {
    const needed = item.Qty * productionQty;
    const normalizedItemId = normalizeItemId(item.ItemID);
    const stockItem = stock.find(
      (s) => normalizeItemId(s.itemid) === normalizedItemId,
    );
    const availableStock = stockItem?.stockAkhir || 0;
    const shortage = Math.max(0, needed - availableStock);

    totalNeeded += needed;
    totalShortage += shortage;

    return {
      ...item,
      needed,
      availableStock,
      shortage,
    };
  });

  return { totalNeeded, totalShortage, items };
};

const calculateMaterialNeedsForCombinedPO = (
  bom: { flat: BomItem[]; tree: BomItem[]; combinedBoms?: any },
  productionOrders: ProductionOrder[],
  stock: StockItem[] = [],
) => {
  if (!bom || !productionOrders || productionOrders.length === 0) {
    return { totalNeeded: 0, totalShortage: 0, items: [] };
  }

  const materialMap = new Map<
    string,
    {
      item: BomItem;
      totalNeeded: number;
    }
  >();

  productionOrders.forEach((po) => {
    let bomForThisItem: BomItem[] = [];

    if (bom.combinedBoms && bom.combinedBoms[po.Kode_Barang]) {
      bomForThisItem = bom.combinedBoms[po.Kode_Barang].flat;
    } else {
      bomForThisItem = bom.flat;
    }

    const componentsOnly = filterOnlyComponents(bomForThisItem);

    componentsOnly.forEach((item) => {
      const neededForThisPO = item.Qty * po.QTY;
      const normalizedId = normalizeItemId(item.ItemID);
      const existing = materialMap.get(normalizedId);

      if (existing) {
        existing.totalNeeded += neededForThisPO;
      } else {
        materialMap.set(normalizedId, {
          item: item,
          totalNeeded: neededForThisPO,
        });
      }
    });
  });

  let totalNeeded = 0;
  let totalShortage = 0;
  const items = Array.from(materialMap.values()).map((material) => {
    const normalizedId = normalizeItemId(material.item.ItemID);
    const stockItem = stock.find(
      (s) => normalizeItemId(s.itemid) === normalizedId,
    );
    const availableStock = stockItem?.stockAkhir || 0;
    const shortage = Math.max(0, material.totalNeeded - availableStock);

    totalNeeded += material.totalNeeded;
    totalShortage += shortage;

    return {
      ...material.item,
      needed: material.totalNeeded,
      availableStock,
      shortage,
    };
  });

  return { totalNeeded, totalShortage, items };
};

// ==================== FUNGSI AMBIL STOK ====================
const fetchStockForItem = async (
  itemId: string,
  orderDate: string,
): Promise<StockItem | null> => {
  try {
    if (!itemId) return null;
    if (!orderDate) orderDate = new Date().toISOString().split("T")[0];

    const apiUrl = `/api/stock/ppic?tgl1=${orderDate}&tgl2=${orderDate}&loc=%25&periodeR=201905&kategori=%25&itemid=${encodeURIComponent(itemId)}`;

    const response = await fetch(apiUrl, {
      method: "GET",
      headers: { "Content-Type": "application/json" },
    });

    if (!response.ok) {
      return {
        itemid: normalizeItemId(itemId),
        itemname: itemId,
        stockAkhir: 0,
        physicalStock: 0,
        committedQty: 0,
        reservedQty: 0,
      };
    }

    const result = await response.json();
    let dataArray = null;

    if (Array.isArray(result)) dataArray = result;
    else if (result.data && Array.isArray(result.data)) dataArray = result.data;
    else if (result.stockData && Array.isArray(result.stockData))
      dataArray = result.stockData;

    if (dataArray && dataArray.length > 0) {
      const itemData = dataArray[0];

      // 🔥 AMBIL KEDUA NILAI STOK
      let saldoAkhir = 0; // Untuk kolom "Stok Akhir"
      let saldoAkhirFisik = 0; // Untuk kolom "Stok Wincp"

      // Ambil SaldoAkhir
      if (
        typeof itemData.SaldoAkhir === "number" &&
        itemData.SaldoAkhir !== undefined
      ) {
        saldoAkhir = Math.max(0, itemData.SaldoAkhir);
      } else if (typeof itemData.stockAkhir === "number") {
        saldoAkhir = Math.max(0, itemData.stockAkhir);
      }

      // Ambil SaldoAkhirFisik (Stok Wincp)
      if (typeof itemData.SaldoAkhirFisik === "number") {
        saldoAkhirFisik = Math.max(0, itemData.SaldoAkhirFisik);
      } else if (typeof itemData.physicalStock === "number") {
        saldoAkhirFisik = Math.max(0, itemData.physicalStock);
      } else if (typeof itemData.totalkgs === "number") {
        saldoAkhirFisik = Math.max(0, itemData.totalkgs);
      }

      console.log(
        `Stock untuk ${itemId}: SaldoAkhir=${saldoAkhir}, SaldoAkhirFisik=${saldoAkhirFisik}`,
      );

      return {
        itemid: normalizeItemId(
          itemData.KodeBarang || itemData.itemid || itemData.ItemID || itemId,
        ),
        itemname:
          itemData.NamaBarang ||
          itemData.itemname ||
          itemData.ItemName ||
          itemId,
        stockAkhir: saldoAkhir, // Untuk kolom "Stok Akhir"
        physicalStock: saldoAkhirFisik, // Untuk kolom "Stok Wincp"
        committedQty: itemData.TotalCommitted || 0,
        reservedQty: itemData.TotalReserved || 0,
      };
    }

    return {
      itemid: normalizeItemId(itemId),
      itemname: itemId,
      stockAkhir: 0,
      physicalStock: 0,
      committedQty: 0,
      reservedQty: 0,
    };
  } catch (err) {
    console.error(`Error fetching stock for ${itemId}:`, err);
    return {
      itemid: normalizeItemId(itemId),
      itemname: itemId,
      stockAkhir: 0,
      physicalStock: 0,
      committedQty: 0,
      reservedQty: 0,
    };
  }
};

const fetchStockForItemsWithCommitment = async (
  itemIds: string[],
  orderDate: string,
): Promise<StockItem[]> => {
  if (!itemIds || itemIds.length === 0) return [];

  const stockData: StockItem[] = [];
  const uniqueItemIds = Array.from(new Set(itemIds))
    .filter((id) => id && id.trim() !== "")
    .map((id) => normalizeItemId(id));

  for (let i = 0; i < uniqueItemIds.length; i++) {
    const stockItem = await fetchStockForItem(uniqueItemIds[i], orderDate);
    if (stockItem) stockData.push(stockItem);
    if (i < uniqueItemIds.length - 1)
      await new Promise((resolve) => setTimeout(resolve, 100));
  }

  return stockData;
};

// ==================== KOMPONEN TOAST NOTIFIKASI ====================
const ToastNotification: React.FC<{
  message: string;
  type: "success" | "error" | "info" | "loading";
  onClose: () => void;
}> = ({ message, type, onClose }) => {
  useEffect(() => {
    const timer = setTimeout(() => {
      onClose();
    }, 5000);
    return () => clearTimeout(timer);
  }, [onClose]);

  const bgColor = {
    success: "bg-green-500",
    error: "bg-red-500",
    info: "bg-blue-500",
    loading: "bg-yellow-500",
  }[type];

  const icon = {
    success: <CheckCircle2 className="h-5 w-5" />,
    error: <AlertCircle className="h-5 w-5" />,
    info: <Info className="h-5 w-5" />,
    loading: <Loader2 className="h-5 w-5 animate-spin" />,
  }[type];

  return (
    <div
      className={`fixed bottom-4 right-4 z-50 ${bgColor} text-white px-4 py-3 rounded-lg shadow-lg flex items-center gap-3 animate-in slide-in-from-bottom-5`}
    >
      {icon}
      <span>{message}</span>
      <button onClick={onClose} className="ml-2 hover:opacity-80">
        ×
      </button>
    </div>
  );
};

// ==================== KOMPONEN COMMITTED PO PANEL ====================
const CommittedPOsPanel: React.FC<{
  committedPOs: CommittedPO[];
  stockReservations: StockReservation[];
  onRefresh: () => void;
  onUncommit?: (noSPK: string) => void;
}> = ({ committedPOs, stockReservations, onRefresh, onUncommit }) => {
  const [expanded, setExpanded] = useState(false);
  const [showUnique, setShowUnique] = useState(true);
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState<string>("COMMITTED");
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  const [uncommitting, setUncommitting] = useState<string | null>(null);
  const [commitDetailDrawer, setCommitDetailDrawer] = useState({
    open: false,
    spk: "",
    commits: [] as CommittedPO[],
  });

  const openCommitDetailDrawer = (noSPK: string) => {
    const allCommitsForSPK = committedPOs
      .filter((p) => p.noSPK === noSPK)
      .sort((a, b) => b.CommitID - a.CommitID);
    setCommitDetailDrawer({
      open: true,
      spk: noSPK,
      commits: allCommitsForSPK,
    });
  };

  const filteredPOs = useMemo(() => {
    let filtered = committedPOs;
    if (searchTerm.trim()) {
      const term = searchTerm.toLowerCase();
      filtered = filtered.filter(
        (po) =>
          po.noSPK.toLowerCase().includes(term) ||
          po.namaPO.toLowerCase().includes(term) ||
          po.kodeBarang.toLowerCase().includes(term),
      );
    }
    if (statusFilter !== "all")
      filtered = filtered.filter((po) => po.status === statusFilter);
    return filtered;
  }, [committedPOs, searchTerm, statusFilter]);

  const displayedPOs = useMemo(() => {
    if (showUnique) {
      const poMap = new Map<string, CommittedPO>();
      filteredPOs.forEach((po) => {
        const existing = poMap.get(po.noSPK);
        if (!existing || po.CommitID > existing.CommitID)
          poMap.set(po.noSPK, po);
      });
      return Array.from(poMap.values()).sort((a, b) => b.CommitID - a.CommitID);
    }
    return filteredPOs.sort((a, b) => b.CommitID - a.CommitID);
  }, [filteredPOs, showUnique]);

  const totalPages = Math.ceil(displayedPOs.length / itemsPerPage);
  const paginatedPOs = displayedPOs.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage,
  );

  // Di dalam komponen CommittedPOsPanel, perbaiki handleUncommit
  const handleUncommit = async (noSPK: string) => {
    if (
      !confirm(
        `Apakah Anda yakin ingin uncommit PO ${noSPK}? Stok akan dikembalikan.`,
      )
    )
      return;
    setUncommitting(noSPK);
    try {
      const response = await fetch("/api/ppic/uncommit-po", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ noSPK, userID: "system" }),
      });

      // Baca response sebagai text
      const responseText = await response.text();
      console.log("Uncommit response from history panel:", responseText);

      // Anggap sukses jika response OK
      if (response.ok) {
        if (onUncommit) onUncommit(noSPK);
        onRefresh();
      } else {
        // Coba parse JSON untuk error message
        try {
          const result = JSON.parse(responseText);
          if (!result.success) {
            console.warn("Uncommit warning:", result.error);
          }
        } catch (e) {
          // Ignore parse error
        }
        // Tetap panggil onRefresh karena mungkin berhasil
        if (onUncommit) onUncommit(noSPK);
        onRefresh();
      }
    } catch (error) {
      console.error("Error uncommit:", error);
      // Tetap refresh
      onRefresh();
    } finally {
      setUncommitting(null);
    }
  };

  const activeReservations = stockReservations.filter(
    (r) => r.status === "RESERVED",
  );
  const totalReservedQty = activeReservations.reduce(
    (sum, r) => sum + r.reservedQty,
    0,
  );

  const getPageNumbers = () => {
    const pages = [];
    const maxVisiblePages = 5;
    let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
    const endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);
    if (endPage - startPage + 1 < maxVisiblePages) {
      startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }
    for (let i = startPage; i <= endPage; i++) pages.push(i);
    return pages;
  };

  const startIndex = (currentPage - 1) * itemsPerPage + 1;
  const endIndex = Math.min(currentPage * itemsPerPage, displayedPOs.length);

  return (
    <Card>
      <CardHeader>
        <div className="flex justify-between items-center flex-wrap gap-4">
          <CardTitle className="flex items-center gap-2">
            <ClipboardList className="h-5 w-5" />
            PO yang Stoknya sudah di hitung
          </CardTitle>
          <div className="flex gap-2">
            <Button
              variant={showUnique ? "default" : "outline"}
              size="sm"
              onClick={() => {
                setShowUnique(!showUnique);
                setCurrentPage(1);
              }}
            >
              {showUnique ? "Show Unique" : "Show All"}
            </Button>
            <Button variant="outline" size="sm" onClick={onRefresh}>
              <RefreshCw className="h-4 w-4" /> Refresh
            </Button>
            <Button
              variant="outline"
              size="sm"
              onClick={() => setExpanded(!expanded)}
            >
              {expanded ? "Sembunyikan" : "Tampilkan"}
            </Button>
          </div>
        </div>
      </CardHeader>
      {expanded && (
        <CardContent>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
            <Card>
              <CardContent className="pt-6">
                <div className="text-sm text-muted-foreground">Total PO</div>
                <div className="text-2xl font-bold">{displayedPOs.length}</div>
              </CardContent>
            </Card>
            <Card>
              <CardContent className="pt-6">
                <div className="text-sm text-muted-foreground">
                  Active Reservations
                </div>
                <div className="text-2xl font-bold text-green-600">
                  {activeReservations.length}
                </div>
              </CardContent>
            </Card>
            <Card>
              <CardContent className="pt-6">
                <div className="text-sm text-muted-foreground">
                  Total Qty Reserved
                </div>
                <div className="text-2xl font-bold text-orange-600">
                  {totalReservedQty.toLocaleString()}
                </div>
              </CardContent>
            </Card>
            <Card>
              <CardContent className="pt-6">
                <div className="text-sm text-muted-foreground">
                  Items Reserved
                </div>
                <div className="text-2xl font-bold text-purple-600">
                  {new Set(activeReservations.map((r) => r.itemID)).size}
                </div>
              </CardContent>
            </Card>
          </div>

          <div className="flex flex-col sm:flex-row gap-4 mb-4">
            <div className="flex-1 relative">
              <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-muted-foreground" />
              <Input
                placeholder="Cari No SPK, Nama PO, atau Kode Barang..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="pl-8"
              />
            </div>
            <Select value={statusFilter} onValueChange={setStatusFilter}>
              <SelectTrigger className="w-48">
                <SelectValue placeholder="Filter Status" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="COMMITTED">Committed</SelectItem>
                <SelectItem value="UNCOMMITTED">Uncommitted</SelectItem>
              </SelectContent>
            </Select>
          </div>

          <div className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Commit ID</TableHead>
                  <TableHead>No SPK</TableHead>
                  <TableHead>PO</TableHead>

                  <TableHead>Tanggal</TableHead>
                  <TableHead className="text-center">Status</TableHead>
                  <TableHead className="text-right">Materials</TableHead>
                  <TableHead className="text-center">Aksi</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {paginatedPOs.map((po) => (
                  <TableRow key={po.CommitID}>
                    <TableCell className="font-mono text-xs">
                      {po.CommitID}
                    </TableCell>
                    <TableCell className="font-medium">{po.noSPK}</TableCell>
                    <TableCell>
                      <div className="font-medium">{po.namaPO}</div>
                    </TableCell>

                    <TableCell>
                      {new Date(po.tanggalCommit).toLocaleDateString("id-ID")}
                    </TableCell>
                    <TableCell className="text-center">
                      <Badge
                        variant={
                          po.status === "COMMITTED" ? "default" : "secondary"
                        }
                      >
                        {po.status}
                      </Badge>
                    </TableCell>
                    <TableCell className="text-right">
                      <div>{po.totalMaterials} items</div>
                      <div className="text-xs">
                        {po.totalQtyReserved.toLocaleString()} qty
                      </div>
                    </TableCell>
                    <TableCell className="text-center">
                      <div className="flex gap-2 justify-center">
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => openCommitDetailDrawer(po.noSPK)}
                        >
                          Detail
                        </Button>
                        {po.status === "COMMITTED" && (
                          <Button
                            variant="destructive"
                            size="sm"
                            onClick={() => handleUncommit(po.noSPK)}
                            disabled={uncommitting === po.noSPK}
                          >
                            {uncommitting === po.noSPK ? (
                              <RefreshCw className="h-3 w-3 animate-spin" />
                            ) : (
                              <Unlock className="h-3 w-3" />
                            )}
                            Uncommit
                          </Button>
                        )}
                      </div>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>

          {displayedPOs.length > 0 && (
            <div className="flex flex-col sm:flex-row justify-between items-center gap-4 mt-4">
              <div className="text-sm text-muted-foreground">
                Menampilkan {startIndex}-{endIndex} dari {displayedPOs.length}{" "}
                data
              </div>
              <div className="flex items-center gap-4">
                <div className="flex items-center gap-2">
                  <Label className="text-sm">Items per page:</Label>
                  <Select
                    value={itemsPerPage.toString()}
                    onValueChange={(v) => {
                      setItemsPerPage(parseInt(v));
                      setCurrentPage(1);
                    }}
                  >
                    <SelectTrigger className="w-20">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="5">5</SelectItem>
                      <SelectItem value="10">10</SelectItem>
                      <SelectItem value="20">20</SelectItem>
                      <SelectItem value="50">50</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                <Pagination>
                  <PaginationContent>
                    <PaginationItem>
                      <PaginationPrevious
                        onClick={() =>
                          setCurrentPage((p) => Math.max(1, p - 1))
                        }
                        className={
                          currentPage === 1
                            ? "pointer-events-none opacity-50"
                            : ""
                        }
                      />
                    </PaginationItem>
                    {getPageNumbers().map((page) => (
                      <PaginationItem key={page}>
                        <PaginationLink
                          onClick={() => setCurrentPage(page)}
                          isActive={currentPage === page}
                        >
                          {page}
                        </PaginationLink>
                      </PaginationItem>
                    ))}
                    <PaginationItem>
                      <PaginationNext
                        onClick={() =>
                          setCurrentPage((p) => Math.min(totalPages, p + 1))
                        }
                        className={
                          currentPage === totalPages
                            ? "pointer-events-none opacity-50"
                            : ""
                        }
                      />
                    </PaginationItem>
                  </PaginationContent>
                </Pagination>
              </div>
            </div>
          )}
        </CardContent>
      )}

      <Sheet
        open={commitDetailDrawer.open}
        onOpenChange={(open) =>
          setCommitDetailDrawer((prev) => ({ ...prev, open }))
        }
      >
        <SheetContent className="sm:max-w-lg overflow-y-auto">
          <SheetHeader>
            <SheetTitle>Detail Commit SPK {commitDetailDrawer.spk}</SheetTitle>
            <SheetDescription>
              Total: {commitDetailDrawer.commits.length} records
            </SheetDescription>
          </SheetHeader>
          <div className="space-y-3 py-4">
            {commitDetailDrawer.commits.map((commit, idx) => (
              <Card key={idx}>
                <CardContent className="pt-6">
                  <div className="flex justify-between">
                    <Badge>ID: {commit.CommitID}</Badge>
                    <Badge
                      variant={
                        commit.status === "COMMITTED" ? "default" : "secondary"
                      }
                    >
                      {commit.status}
                    </Badge>
                  </div>
                  <div className="grid grid-cols-2 gap-3 mt-3">
                    <div>
                      <Label>Tanggal</Label>
                      <div className="font-medium">
                        {new Date(commit.tanggalCommit).toLocaleDateString(
                          "id-ID",
                        )}
                      </div>
                    </div>
                    <div>
                      <Label>Qty</Label>
                      <div className="font-medium text-right">
                        {commit.qty.toLocaleString()}
                      </div>
                    </div>
                    <div>
                      <Label>Materials</Label>
                      <div className="font-medium">
                        {commit.totalMaterials} items
                      </div>
                    </div>
                    <div>
                      <Label>Qty Reserved</Label>
                      <div className="font-medium text-right">
                        {commit.totalQtyReserved.toLocaleString()}
                      </div>
                    </div>
                  </div>
                </CardContent>
              </Card>
            ))}
          </div>
        </SheetContent>
      </Sheet>
    </Card>
  );
};

// ==================== KOMPONEN PREVIEW DIALOG ====================
// Komponen HistoryPanel dengan Sheet (bukan modal) dan Pagination
const HistoryPanel: React.FC<{
  calculations: any[];
  onLoadCalculation: (calculation: any) => void;
  onRefresh: () => void;
  onDeleteCalculation?: (calculationId: string) => void;
}> = ({ calculations, onLoadCalculation, onRefresh, onDeleteCalculation }) => {
  const [expanded, setExpanded] = useState(true); // Default expanded
  const [selectedCalc, setSelectedCalc] = useState<any>(null);
  const [detailSheetOpen, setDetailSheetOpen] = useState(false);
  const [searchHistory, setSearchHistory] = useState("");
  const [dateFilter, setDateFilter] = useState("");

  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);
  const [isMobile, setIsMobile] = useState(false);

  // Detect mobile screen
  useEffect(() => {
    const checkMobile = () => {
      setIsMobile(window.innerWidth < 768);
    };
    checkMobile();
    window.addEventListener("resize", checkMobile);
    return () => window.removeEventListener("resize", checkMobile);
  }, []);

  // Filter history berdasarkan search
  const filteredCalculations = useMemo(() => {
    if (!calculations.length) return [];

    let filtered = [...calculations];

    if (searchHistory) {
      const term = searchHistory.toLowerCase();
      filtered = filtered.filter(
        (calc) =>
          calc.calculation_id?.toLowerCase().includes(term) ||
          calc.user_id?.toLowerCase().includes(term) ||
          calc.notes?.toLowerCase().includes(term) ||
          calc.calculation_name?.toLowerCase().includes(term),
      );
    }

    if (dateFilter) {
      filtered = filtered.filter(
        (calc) => calc.calculation_date?.split("T")[0] === dateFilter,
      );
    }

    return filtered;
  }, [calculations, searchHistory, dateFilter]);

  // Pagination calculations
  const totalItems = filteredCalculations.length;
  const totalPages = Math.ceil(totalItems / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedCalculations = filteredCalculations.slice(
    startIndex,
    endIndex,
  );

  // Reset to first page when filters change
  useEffect(() => {
    setCurrentPage(1);
  }, [searchHistory, dateFilter]);

  const getStatusSummary = (calc: any) => {
    const total =
      calc.material_aman + calc.material_kurang + calc.material_habis;
    const amanPercent = total > 0 ? (calc.material_aman / total) * 100 : 0;
    return { total, amanPercent };
  };

  // Handle page change
  const handlePageChange = (page: number) => {
    setCurrentPage(page);
    // Scroll to top of table when page changes
    const tableElement = document.getElementById("history-table");
    if (tableElement) {
      tableElement.scrollIntoView({ behavior: "smooth" });
    }
  };

  // Handle page size change
  const handlePageSizeChange = (newSize: number) => {
    setPageSize(newSize);
    setCurrentPage(1);
  };

  // Generate page numbers for pagination
  const getPageNumbers = () => {
    const pageNumbers = [];
    const maxVisible = isMobile ? 3 : 5;
    let startPage = Math.max(1, currentPage - Math.floor(maxVisible / 2));
    let endPage = Math.min(totalPages, startPage + maxVisible - 1);

    if (endPage - startPage + 1 < maxVisible) {
      startPage = Math.max(1, endPage - maxVisible + 1);
    }

    for (let i = startPage; i <= endPage; i++) {
      pageNumbers.push(i);
    }

    return pageNumbers;
  };

  if (!expanded) {
    return (
      <Card>
        <CardHeader>
          <div className="flex justify-between items-center">
            <CardTitle
              className="flex items-center gap-2 cursor-pointer"
              onClick={() => setExpanded(true)}
            >
              <Database className="h-5 w-5" />
              History Perhitungan Material Tersimpan
              <Badge variant="secondary" className="ml-2">
                {calculations.length} Data
              </Badge>
            </CardTitle>
            <Button
              variant="outline"
              size="sm"
              onClick={() => setExpanded(true)}
            >
              <Eye className="h-4 w-4 mr-2" />
              Tampilkan History
            </Button>
          </div>
        </CardHeader>
      </Card>
    );
  }

  return (
    <>
      <Card>
        <CardHeader>
          <div className="flex justify-between items-center flex-wrap gap-4">
            <CardTitle className="flex items-center gap-2">
              <Database className="h-5 w-5" />
              History Perhitungan Material (Tersimpan di Database)
              <Badge variant="secondary" className="ml-2">
                {calculations.length} Data
              </Badge>
            </CardTitle>
            <div className="flex gap-2">
              <Button variant="outline" size="sm" onClick={onRefresh}>
                <RefreshCw className="h-4 w-4 mr-2" />
                Refresh
              </Button>
              <Button
                variant="outline"
                size="sm"
                onClick={() => setExpanded(false)}
              >
                Sembunyikan
              </Button>
            </div>
          </div>
        </CardHeader>
        <CardContent>
          {/* Filter Section */}
          <div className="flex gap-4 mb-4 flex-wrap">
            <div className="flex-1 relative min-w-50">
              <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-muted-foreground" />
              <Input
                placeholder="Cari ID Perhitungan, Nama, atau User..."
                value={searchHistory}
                onChange={(e) => setSearchHistory(e.target.value)}
                className="pl-8"
              />
            </div>
            <div className="w-48">
              <Input
                type="date"
                placeholder="Filter Tanggal"
                value={dateFilter}
                onChange={(e) => setDateFilter(e.target.value)}
              />
            </div>
            {(searchHistory || dateFilter) && (
              <Button
                variant="ghost"
                size="sm"
                onClick={() => {
                  setSearchHistory("");
                  setDateFilter("");
                }}
              >
                Clear Filter
              </Button>
            )}
          </div>

          {/* Page Size Selector and Info */}
          {filteredCalculations.length > 0 && (
            <div className="flex justify-between items-center mb-4 flex-wrap gap-2">
              <div className="flex items-center gap-2 text-sm text-muted-foreground">
                <span>Menampilkan</span>
                <select
                  value={pageSize}
                  onChange={(e) => handlePageSizeChange(Number(e.target.value))}
                  className="border rounded px-2 py-1 text-sm"
                >
                  <option value={5}>5</option>
                  <option value={10}>10</option>
                  <option value={20}>20</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                </select>
                <span>data dari {totalItems} total</span>
              </div>
              <div className="text-sm text-muted-foreground">
                Halaman {currentPage} dari {totalPages || 1}
              </div>
            </div>
          )}

          {filteredCalculations.length === 0 ? (
            <div className="text-center py-8 text-muted-foreground">
              <Database className="h-12 w-12 mx-auto mb-2 opacity-50" />
              <p>Belum ada data perhitungan yang tersimpan</p>
              <p className="text-sm">
                Silahkan preview PO dan klik Simpan ke Database
              </p>
            </div>
          ) : (
            <>
              <div className="overflow-x-auto" id="history-table">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead className="w-12.5">No</TableHead>
                      <TableHead>Nama Perhitungan</TableHead>
                      <TableHead>Tanggal</TableHead>
                      <TableHead>User</TableHead>
                      <TableHead className="text-center">Total PO</TableHead>
                      <TableHead className="text-center">
                        Total Material
                      </TableHead>
                      <TableHead>Status</TableHead>
                      <TableHead className="text-right">Kebutuhan</TableHead>
                      <TableHead className="text-right">Kekurangan</TableHead>
                      <TableHead className="text-center">Aksi</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {paginatedCalculations.map((calc, idx) => {
                      const { total, amanPercent } = getStatusSummary(calc);
                      const globalIndex = startIndex + idx + 1;
                      return (
                        <TableRow key={calc.id}>
                          <TableCell>{globalIndex}</TableCell>
                          <TableCell>
                            <div className="font-medium">
                              {calc.calculation_name || calc.calculation_id}
                            </div>
                            <div className="text-xs text-muted-foreground font-mono">
                              ID: {calc.calculation_id}
                            </div>
                          </TableCell>
                          <TableCell className="whitespace-nowrap">
                            {new Date(calc.calculation_date).toLocaleString(
                              "id-ID",
                            )}
                          </TableCell>
                          <TableCell>{calc.user_id}</TableCell>
                          <TableCell className="text-center font-bold">
                            {calc.total_po}
                          </TableCell>
                          <TableCell className="text-center font-bold">
                            {calc.total_materials}
                          </TableCell>
                          <TableCell>
                            <div className="flex gap-1 flex-wrap">
                              {calc.material_aman > 0 && (
                                <Badge className="bg-green-500">
                                  ✅ {calc.material_aman}
                                </Badge>
                              )}
                              {calc.material_kurang > 0 && (
                                <Badge className="bg-orange-500">
                                  ⚠️ {calc.material_kurang}
                                </Badge>
                              )}
                              {calc.material_habis > 0 && (
                                <Badge className="bg-red-500">
                                  ❌ {calc.material_habis}
                                </Badge>
                              )}
                            </div>
                            <div className="text-xs text-muted-foreground mt-1">
                              Keamanan: {amanPercent.toFixed(0)}%
                            </div>
                          </TableCell>
                          <TableCell className="text-right">
                            {calc.total_kebutuhan?.toLocaleString() || 0}
                          </TableCell>
                          <TableCell className="text-right text-red-600 font-bold">
                            {calc.total_kekurangan?.toLocaleString() || 0}
                          </TableCell>
                          <TableCell>
                            <div className="flex gap-2">
                              <Button
                                variant="outline"
                                size="sm"
                                onClick={() => {
                                  setSelectedCalc(calc);
                                  setDetailSheetOpen(true);
                                }}
                              >
                                <Eye className="h-3 w-3 mr-1" />
                                Detail
                              </Button>
                              <Button
                                variant="secondary"
                                size="sm"
                                onClick={() => onLoadCalculation(calc)}
                              >
                                <Download className="h-3 w-3 mr-1" />
                                Load
                              </Button>
                              {onDeleteCalculation && (
                                <Button
                                  variant="destructive"
                                  size="sm"
                                  onClick={() => {
                                    if (
                                      confirm(
                                        `Hapus perhitungan "${calc.calculation_name || calc.calculation_id}"?`,
                                      )
                                    ) {
                                      onDeleteCalculation(calc.calculation_id);
                                    }
                                  }}
                                >
                                  <Trash2 className="h-3 w-3" />
                                </Button>
                              )}
                            </div>
                          </TableCell>
                        </TableRow>
                      );
                    })}
                  </TableBody>
                </Table>
              </div>

              {/* Pagination Component */}
              {totalPages > 1 && (
                <div className="flex justify-center items-center gap-2 mt-6 flex-wrap">
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => handlePageChange(1)}
                    disabled={currentPage === 1}
                  >
                    <ChevronLeft className="h-4 w-4 mr-1" />
                    <span className={isMobile ? "hidden" : "inline"}>
                      First
                    </span>
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => handlePageChange(currentPage - 1)}
                    disabled={currentPage === 1}
                  >
                    <ChevronLeft className="h-4 w-4" />
                    <span className={isMobile ? "hidden" : "ml-1"}>
                      Previous
                    </span>
                  </Button>

                  <div className="flex gap-1">
                    {currentPage > 2 && totalPages > 5 && (
                      <>
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => handlePageChange(1)}
                        >
                          1
                        </Button>
                        {currentPage > 3 && (
                          <span className="px-2 self-center">...</span>
                        )}
                      </>
                    )}

                    {getPageNumbers().map((pageNum) => (
                      <Button
                        key={pageNum}
                        variant={
                          currentPage === pageNum ? "default" : "outline"
                        }
                        size="sm"
                        onClick={() => handlePageChange(pageNum)}
                        className={currentPage === pageNum ? "bg-blue-600" : ""}
                      >
                        {pageNum}
                      </Button>
                    ))}

                    {currentPage < totalPages - 1 && totalPages > 5 && (
                      <>
                        {currentPage < totalPages - 2 && (
                          <span className="px-2 self-center">...</span>
                        )}
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => handlePageChange(totalPages)}
                        >
                          {totalPages}
                        </Button>
                      </>
                    )}
                  </div>

                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => handlePageChange(currentPage + 1)}
                    disabled={currentPage === totalPages}
                  >
                    <span className={isMobile ? "hidden" : "mr-1"}>Next</span>
                    <ChevronRight className="h-4 w-4" />
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => handlePageChange(totalPages)}
                    disabled={currentPage === totalPages}
                  >
                    <span className={isMobile ? "hidden" : "inline"}>Last</span>
                    <ChevronRight className="h-4 w-4 ml-1" />
                  </Button>
                </div>
              )}

              {/* Mobile optimized pagination - simple */}
              {isMobile && totalPages > 1 && (
                <div className="flex justify-center items-center gap-2 mt-4">
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => handlePageChange(currentPage - 1)}
                    disabled={currentPage === 1}
                  >
                    Previous
                  </Button>
                  <span className="text-sm">
                    {currentPage} / {totalPages}
                  </span>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => handlePageChange(currentPage + 1)}
                    disabled={currentPage === totalPages}
                  >
                    Next
                  </Button>
                </div>
              )}

              {/* Info text about visible data */}
              <div className="text-center text-xs text-muted-foreground mt-4">
                Menampilkan data {startIndex + 1} -{" "}
                {Math.min(endIndex, totalItems)} dari {totalItems} total
                perhitungan
              </div>
            </>
          )}
        </CardContent>
      </Card>

      {/* Detail Sheet - menggunakan Sheet seperti preview dialog */}
      <Sheet open={detailSheetOpen} onOpenChange={setDetailSheetOpen}>
        <SheetContent side="bottom" className="h-[90vh] p-0 flex flex-col">
          <div className="flex-shrink-0 border-b px-6 py-4 bg-white">
            <SheetHeader className="text-left">
              <SheetTitle className="flex items-center gap-2">
                <Database className="h-5 w-5" />
                Detail Perhitungan:{" "}
                {selectedCalc?.calculation_name || selectedCalc?.calculation_id}
              </SheetTitle>
              <SheetDescription>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mt-2">
                  <div>
                    <div className="text-xs text-muted-foreground">Tanggal</div>
                    <div className="font-medium">
                      {selectedCalc?.calculation_date &&
                        new Date(selectedCalc.calculation_date).toLocaleString(
                          "id-ID",
                        )}
                    </div>
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground">User</div>
                    <div className="font-medium">{selectedCalc?.user_id}</div>
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground">
                      Total PO
                    </div>
                    <div className="font-medium">{selectedCalc?.total_po}</div>
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground">
                      Total Material
                    </div>
                    <div className="font-medium">
                      {selectedCalc?.total_materials}
                    </div>
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground">
                      Total Kebutuhan
                    </div>
                    <div className="font-medium">
                      {selectedCalc?.total_kebutuhan?.toLocaleString()}
                    </div>
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground text-red-600">
                      Total Kekurangan
                    </div>
                    <div className="font-medium text-red-600">
                      {selectedCalc?.total_kekurangan?.toLocaleString()}
                    </div>
                  </div>
                </div>
              </SheetDescription>
            </SheetHeader>
          </div>

          {/* Table Material Data with pagination in sheet */}
          <div className="flex-1 overflow-auto px-6 py-4 min-h-0">
            <div className="border rounded-md h-full">
              <div
                className="overflow-auto"
                style={{ maxHeight: "calc(90vh - 280px)" }}
              >
                <Table>
                  <TableHeader className="sticky top-0 bg-gray-50 z-10">
                    <TableRow>
                      <TableHead>Kode Material</TableHead>
                      <TableHead>Nama Material</TableHead>
                      <TableHead>Departemen</TableHead>
                      <TableHead className="text-right">Kebutuhan</TableHead>
                      <TableHead className="text-right">Stok Wincp</TableHead>
                      <TableHead className="text-right">Reserved</TableHead>
                      <TableHead className="text-right">Dibutuhkan</TableHead>
                      <TableHead className="text-right">Available</TableHead>
                      <TableHead>Status</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {selectedCalc?.material_data
                      ?.slice(0, 100)
                      .map((item: any, idx: number) => (
                        <TableRow key={idx}>
                          <TableCell className="font-mono text-xs">
                            {item["Kode Material"]}
                          </TableCell>
                          <TableCell className="text-sm">
                            {item["Nama Material"]}
                          </TableCell>
                          <TableCell>{item["Departemen"]}</TableCell>
                          <TableCell className="text-right">
                            {item["Total Kebutuhan"]}
                          </TableCell>
                          <TableCell className="text-right">
                            {item["Stok Wincp (Real)"]}
                          </TableCell>
                          <TableCell className="text-right">
                            {item["Qty Reserved (PO Lain)"]}
                          </TableCell>
                          <TableCell className="text-right">
                            {item["Total Dibutuhkan"]}
                          </TableCell>
                          <TableCell className="text-right">
                            {item["Qty Available"]}
                          </TableCell>
                          <TableCell
                            className={
                              item["Status Stock"] === "AMAN"
                                ? "text-green-600 font-bold"
                                : item["Status Stock"] === "KURANG"
                                  ? "text-orange-600 font-bold"
                                  : "text-red-600 font-bold"
                            }
                          >
                            {item["Status Stock"]}
                          </TableCell>
                        </TableRow>
                      ))}
                  </TableBody>
                </Table>
                {selectedCalc?.material_data?.length > 100 && (
                  <div className="text-center text-muted-foreground text-sm py-4">
                    Menampilkan 100 dari {selectedCalc?.material_data?.length}{" "}
                    material
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Footer */}
          <div className="flex-shrink-0 border-t px-6 py-4 bg-white">
            <div className="flex justify-end gap-2">
              <Button
                variant="outline"
                onClick={() => setDetailSheetOpen(false)}
              >
                Tutup
              </Button>
              <Button
                onClick={() => {
                  onLoadCalculation(selectedCalc);
                  setDetailSheetOpen(false);
                }}
              >
                <Download className="h-4 w-4 mr-2" />
                Load Perhitungan Ini
              </Button>
            </div>
          </div>
        </SheetContent>
      </Sheet>
    </>
  );
};
// ==================== KOMPONEN BOM OVERRIDE MANAGER (DENGAN DATABASE) - LENGKAP ====================
const BomOverrideManager: React.FC<{
  overrides: BomOverride[];
  onAdd: (override: BomOverride) => Promise<void>;
  onUpdate: (index: number, override: BomOverride) => Promise<void>;
  onDelete: (index: number, id: number) => Promise<void>;
  onToggle: (index: number, id: number, isActive: boolean) => Promise<void>;
  onClose: () => void;
  loading?: boolean;
}> = ({ overrides, onAdd, onUpdate, onDelete, onToggle, onClose, loading = false }) => {
  const [showAddForm, setShowAddForm] = useState(false);
  const [newOverride, setNewOverride] = useState<Partial<BomOverride>>({
    originalItemId: "",
    replacementItemId: "",
    replacementItemName: "",
    replacementItemName2: "",
    isActive: true,
  });
  const [searchOriginal, setSearchOriginal] = useState("");
  const [searchReplacement, setSearchReplacement] = useState("");
  const [searchResults, setSearchResults] = useState<
    Array<{ itemid: string; itemname: string; itemname2?: string }>
  >([]);
  const [searching, setSearching] = useState(false);
  const [selectedField, setSelectedField] = useState<
    "original" | "replacement"
  >("original");
  const [editId, setEditId] = useState<number | null>(null);
  const [saving, setSaving] = useState(false);
  
  // ==================== TAMBAHKAN STATE UNTUK TARGET BOM ====================
  const [targetType, setTargetType] = useState<"all" | "specific">("all");
  const [selectedTargets, setSelectedTargets] = useState<string[]>([]);
  const [availableProducts, setAvailableProducts] = useState<Array<{kode: string, nama: string}>>([]);
  const [searchProduct, setSearchProduct] = useState("");
  const [loadingProducts, setLoadingProducts] = useState(false);

  // Load daftar produk yang tersedia dari API
  const loadAvailableProducts = async () => {
    setLoadingProducts(true);
    try {
      // Ambil dari API /api/ppic/products atau dari data orders yang sudah ada
      // Karena mungkin belum ada API, kita bisa ambil dari orders yang sudah dimuat
      // Untuk sementara, kita buat API call terlebih dahulu
      const response = await fetch("/api/ppic/products");
      if (response.ok) {
        const result = await response.json();
        if (result.success) {
          setAvailableProducts(result.data);
          return;
        }
      }
      // Fallback: ambil dari sessionStorage atau buat array kosong
      const cachedProducts = sessionStorage.getItem("availableProducts");
      if (cachedProducts) {
        setAvailableProducts(JSON.parse(cachedProducts));
      } else {
        // Jika tidak ada API, kita bisa fetch dari orders yang ada di komponen induk
        // Tapi karena ini komponen terpisah, kita akan panggil props tambahan
        console.warn("Tidak dapat memuat daftar produk");
      }
    } catch (error) {
      console.error("Error loading products:", error);
    } finally {
      setLoadingProducts(false);
    }
  };

  useEffect(() => {
    loadAvailableProducts();
  }, []);

  // Simpan ke sessionStorage
  useEffect(() => {
    if (availableProducts.length > 0) {
      sessionStorage.setItem("availableProducts", JSON.stringify(availableProducts));
    }
  }, [availableProducts]);

  const searchItems = async (query: string) => {
    if (!query || query.length < 2) {
      setSearchResults([]);
      return;
    }

    setSearching(true);
    try {
      const response = await fetch(
        `/api/stock/ppic?itemid=${encodeURIComponent(query)}&limit=20`,
      );
      const result = await response.json();
      let items: Array<{ itemid: string; itemname: string; itemname2?: string }> = [];

      if (Array.isArray(result)) {
        items = result.map((r: any) => ({
          itemid: r.KodeBarang || r.itemid || r.ItemID,
          itemname: r.NamaBarang || r.itemname || r.ItemName,
          itemname2: r.NamaBarang2 || r.itemname2 || r.ItemName2,
        }));
      } else if (result.data && Array.isArray(result.data)) {
        items = result.data.map((r: any) => ({
          itemid: r.KodeBarang || r.itemid || r.ItemID,
          itemname: r.NamaBarang || r.itemname || r.ItemName,
          itemname2: r.NamaBarang2 || r.itemname2 || r.ItemName2,
        }));
      }

      setSearchResults(items.filter((i) => i.itemid && i.itemid !== ""));
    } catch (error) {
      console.error("Error searching items:", error);
      setSearchResults([]);
    } finally {
      setSearching(false);
    }
  };

  const handleSelectItem = (item: { itemid: string; itemname: string; itemname2?: string }) => {
    console.log("Selected item:", item, "Field:", selectedField);
    
    if (selectedField === "original") {
      setNewOverride((prev) => {
        const updated = {
          ...prev,
          originalItemId: item.itemid,
        };
        console.log("Updated newOverride (original):", updated);
        return updated;
      });
      setSearchOriginal(item.itemid);
    } else {
      setNewOverride((prev) => {
        const updated = {
          ...prev,
          replacementItemId: item.itemid,
          replacementItemName: item.itemname,
          replacementItemName2: item.itemname2 || item.itemname,
        };
        console.log("Updated newOverride (replacement):", updated);
        return updated;
      });
      setSearchReplacement(item.itemid);
    }
    setSearchResults([]);
  };

  const handleAddOrUpdateOverride = async () => {
    const originalId = newOverride.originalItemId?.trim();
    const replacementId = newOverride.replacementItemId?.trim();

    if (!originalId || originalId === "") {
      alert("Mohon pilih item asli terlebih dahulu");
      return;
    }

    if (!replacementId || replacementId === "") {
      alert("Mohon pilih item pengganti terlebih dahulu");
      return;
    }

    setSaving(true);
    try {
      const overrideData: any = {
        originalItemId: originalId,
        replacementItemId: replacementId,
        replacementItemName: newOverride.replacementItemName || replacementId,
        replacementItemName2: newOverride.replacementItemName2 || replacementId,
        isActive: newOverride.isActive ?? true,
      };

      // Tambahkan target jika specific
      if (targetType === "specific" && selectedTargets.length > 0) {
        if (selectedTargets.length === 1) {
          overrideData.targetKodeBarang = selectedTargets[0];
        } else {
          overrideData.targetKodeBarangs = selectedTargets;
        }
      }

      if (editId !== null) {
        await onUpdate(editId, { ...overrideData, id: editId });
      } else {
        await onAdd(overrideData);
      }

      // Reset form
      setNewOverride({
        originalItemId: "",
        replacementItemId: "",
        replacementItemName: "",
        replacementItemName2: "",
        isActive: true,
      });
      setSelectedTargets([]);
      setTargetType("all");
      setShowAddForm(false);
      setSearchOriginal("");
      setSearchReplacement("");
      setEditId(null);
    } catch (error) {
      console.error("Error saving override:", error);
      alert("Gagal menyimpan perubahan BOM: " + (error instanceof Error ? error.message : "Unknown error"));
    } finally {
      setSaving(false);
    }
  };

  const handleEdit = (override: BomOverride, index: number) => {
    setEditId(override.id || null);
    setNewOverride({
      originalItemId: override.originalItemId,
      replacementItemId: override.replacementItemId,
      replacementItemName: override.replacementItemName,
      replacementItemName2: override.replacementItemName2,
      isActive: override.isActive,
    });
    setSearchOriginal(override.originalItemId);
    setSearchReplacement(override.replacementItemId);
    
    // Parse target yang sudah ada
    if ((override as any).targetKodeBarang) {
      setTargetType("specific");
      setSelectedTargets([(override as any).targetKodeBarang]);
    } else if ((override as any).targetKodeBarangs && (override as any).targetKodeBarangs.length > 0) {
      setTargetType("specific");
      setSelectedTargets((override as any).targetKodeBarangs);
    } else {
      setTargetType("all");
      setSelectedTargets([]);
    }
    
    setShowAddForm(true);
  };

  const handleCancelEdit = () => {
    setEditId(null);
    setNewOverride({
      originalItemId: "",
      replacementItemId: "",
      replacementItemName: "",
      replacementItemName2: "",
      isActive: true,
    });
    setSearchOriginal("");
    setSearchReplacement("");
    setSelectedTargets([]);
    setTargetType("all");
    setShowAddForm(false);
  };

  return (
    <Sheet open={true} onOpenChange={onClose}>
      <SheetContent className="sm:max-w-lg overflow-y-auto">
        <SheetHeader>
          <SheetTitle className="flex items-center gap-2">
            <TreePine className="h-5 w-5" />
            Pengaturan Perubahan BOM
          </SheetTitle>
          <SheetDescription>
            Ganti material tertentu dengan material lain untuk kebutuhan
            produksi khusus. Perubahan ini akan tersimpan di database dan tidak
            hilang meskipun halaman di-refresh.
          </SheetDescription>
        </SheetHeader>

        <div className="space-y-4 py-4">
          {/* Daftar Override Aktif */}
          <div>
            <div className="flex justify-between items-center mb-3">
              <h3 className="font-semibold">Daftar Perubahan BOM</h3>
              <Button
                variant="outline"
                size="sm"
                onClick={() => {
                  setEditId(null);
                  setShowAddForm(!showAddForm);
                }}
                disabled={loading || saving}
              >
                + Tambah Perubahan
              </Button>
            </div>

            {loading ? (
              <div className="text-center py-8">
                <Loader2 className="h-8 w-8 animate-spin mx-auto" />
                <p className="mt-2 text-muted-foreground">Memuat data...</p>
              </div>
            ) : overrides.length === 0 ? (
              <div className="text-center py-8 text-muted-foreground border rounded-lg">
                <TreePine className="h-12 w-12 mx-auto mb-2 opacity-50" />
                <p>Belum ada perubahan BOM</p>
                <p className="text-sm">Klik tombol di atas untuk menambahkan</p>
              </div>
            ) : (
              <div className="space-y-2">
                {overrides.map((override, idx) => (
                  <Card
                    key={override.id || idx}
                    className={!override.isActive ? "opacity-60" : ""}
                  >
                    <CardContent className="pt-4">
                      <div className="flex justify-between items-start">
                        <div className="flex-1">
                          <div className="flex items-center gap-2 mb-2">
                            <Badge
                              variant={
                                override.isActive ? "default" : "secondary"
                              }
                            >
                              {override.isActive ? "AKTIF" : "NONAKTIF"}
                            </Badge>
                            <span className="text-xs text-muted-foreground">
                              ID: {override.id || idx + 1}
                            </span>
                          </div>
                          <div className="grid grid-cols-2 gap-2 text-sm">
                            <div>
                              <span className="text-muted-foreground">
                                Item Asli:
                              </span>
                              <div className="font-mono text-xs break-all">
                                {override.originalItemId}
                              </div>
                            </div>
                            <div>
                              <span className="text-muted-foreground">
                                Item Pengganti:
                              </span>
                              <div className="font-mono text-xs break-all">
                                {override.replacementItemId}
                              </div>
                              <div className="text-xs truncate">
                                {override.replacementItemName}
                              </div>
                            </div>
                          </div>
                          {/* Tampilkan target jika ada */}
                          {((override as any).targetKodeBarang || (override as any).targetKodeBarangs) && (
                            <div className="mt-2 text-xs text-blue-600">
                              🎯 Berlaku untuk:{' '}
                              {(override as any).targetKodeBarang || 
                               (override as any).targetKodeBarangs?.join(", ")}
                            </div>
                          )}
                        </div>
                        <div className="flex gap-1">
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={() => handleEdit(override, idx)}
                            disabled={saving}
                          >
                            <svg
                              xmlns="http://www.w3.org/2000/svg"
                              className="h-3 w-3"
                              fill="none"
                              viewBox="0 0 24 24"
                              stroke="currentColor"
                            >
                              <path
                                strokeLinecap="round"
                                strokeLinejoin="round"
                                strokeWidth={2}
                                d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"
                              />
                            </svg>
                          </Button>
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={() =>
                              onToggle(idx, override.id!, !override.isActive)
                            }
                            disabled={saving}
                          >
                            {override.isActive ? (
                              <Lock className="h-3 w-3" />
                            ) : (
                              <Unlock className="h-3 w-3" />
                            )}
                          </Button>
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={() => {
                              if (
                                confirm(
                                  `Hapus perubahan BOM untuk ${override.originalItemId}?`,
                                )
                              ) {
                                onDelete(idx, override.id!);
                              }
                            }}
                            disabled={saving}
                          >
                            <Trash2 className="h-3 w-3" />
                          </Button>
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                ))}
              </div>
            )}
          </div>
          
          {/* Form Tambah/Edit Override */}
          {showAddForm && (
            <div className="border rounded-lg p-4 space-y-3">
              <h3 className="font-semibold text-sm">
                {editId !== null
                  ? "Edit Perubahan BOM"
                  : "Tambah Perubahan BOM Baru"}
              </h3>

              {/* Item Asli */}
              <div>
                <Label>Item Asli (yang akan diganti)</Label>
                <div className="relative">
                  <Input
                    placeholder="Cari item asli..."
                    value={searchOriginal}
                    onChange={(e) => {
                      setSearchOriginal(e.target.value);
                      setSelectedField("original");
                      searchItems(e.target.value);
                    }}
                    disabled={saving}
                  />
                  {searchResults.length > 0 && selectedField === "original" && (
                    <div className="absolute z-10 top-full left-0 right-0 bg-white border rounded-md shadow-lg mt-1 max-h-48 overflow-auto">
                      {searching ? (
                        <div className="px-3 py-2 text-center text-muted-foreground">
                          <Loader2 className="h-4 w-4 animate-spin inline mr-2" />
                          Mencari...
                        </div>
                      ) : (
                        searchResults.map((item) => (
                          <div
                            key={item.itemid}
                            className="px-3 py-2 hover:bg-gray-100 cursor-pointer text-sm"
                            onClick={() => handleSelectItem(item)}
                          >
                            <div className="font-mono text-xs">
                              {item.itemid}
                            </div>
                            <div className="text-xs truncate">
                              {item.itemname}
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  )}
                </div>
                {newOverride.originalItemId && (
                  <div className="mt-1 text-xs text-green-600">
                    ✓ Item asli dipilih: {newOverride.originalItemId}
                  </div>
                )}
              </div>

              {/* Item Pengganti */}
              <div>
                <Label>Item Pengganti</Label>
                <div className="relative">
                  <Input
                    placeholder="Cari item pengganti..."
                    value={searchReplacement}
                    onChange={(e) => {
                      setSearchReplacement(e.target.value);
                      setSelectedField("replacement");
                      searchItems(e.target.value);
                    }}
                    disabled={saving}
                  />
                  {searchResults.length > 0 &&
                    selectedField === "replacement" && (
                      <div className="absolute z-10 top-full left-0 right-0 bg-white border rounded-md shadow-lg mt-1 max-h-48 overflow-auto">
                        {searching ? (
                          <div className="px-3 py-2 text-center text-muted-foreground">
                            <Loader2 className="h-4 w-4 animate-spin inline mr-2" />
                            Mencari...
                          </div>
                        ) : (
                          searchResults.map((item) => (
                            <div
                              key={item.itemid}
                              className="px-3 py-2 hover:bg-gray-100 cursor-pointer text-sm"
                              onClick={() => handleSelectItem(item)}
                            >
                              <div className="font-mono text-xs">
                                {item.itemid}
                              </div>
                              <div className="text-xs truncate">
                                {item.itemname}
                              </div>
                              {item.itemname2 && (
                                <div className="text-xs text-muted-foreground">
                                  {item.itemname2}
                                </div>
                              )}
                            </div>
                          ))
                        )}
                      </div>
                    )}
                </div>
                {newOverride.replacementItemId && (
                  <div className="mt-1 text-xs text-green-600">
                    ✓ Item pengganti dipilih: {newOverride.replacementItemId} -{" "}
                    {newOverride.replacementItemName}
                  </div>
                )}
              </div>

              {/* ==================== TARGET BOM (TAMBAHKAN INI) ==================== */}
              <div>
                <Label>Berlaku Untuk</Label>
                <Select 
                  value={targetType} 
                  onValueChange={(val: "all" | "specific") => setTargetType(val)}
                  disabled={saving}
                >
                  <SelectTrigger>
                    <SelectValue placeholder="Pilih cakupan" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">Semua BOM (Global)</SelectItem>
                    <SelectItem value="specific">Hanya BOM Tertentu</SelectItem>
                  </SelectContent>
                </Select>
              </div>

              {targetType === "specific" && (
                <div>
                  <Label>Pilih BOM/Produk</Label>
                  <div className="relative">
                    <Input
                      placeholder="Cari produk..."
                      value={searchProduct}
                      onChange={(e) => setSearchProduct(e.target.value)}
                      disabled={saving || loadingProducts}
                    />
                    {searchProduct && (
                      <div className="absolute z-10 top-full left-0 right-0 bg-white border rounded-md shadow-lg mt-1 max-h-48 overflow-auto">
                        {loadingProducts ? (
                          <div className="px-3 py-2 text-center text-muted-foreground">
                            <Loader2 className="h-4 w-4 animate-spin inline mr-2" />
                            Memuat produk...
                          </div>
                        ) : (
                          availableProducts
                            .filter(p => 
                              p.kode.toLowerCase().includes(searchProduct.toLowerCase()) || 
                              p.nama.toLowerCase().includes(searchProduct.toLowerCase())
                            )
                            .map((product) => (
                              <div
                                key={product.kode}
                                className="px-3 py-2 hover:bg-gray-100 cursor-pointer text-sm"
                                onClick={() => {
                                  if (!selectedTargets.includes(product.kode)) {
                                    setSelectedTargets([...selectedTargets, product.kode]);
                                  }
                                  setSearchProduct("");
                                }}
                              >
                                <div className="font-mono text-xs">{product.kode}</div>
                                <div className="text-xs truncate">{product.nama}</div>
                              </div>
                            ))
                        )}
                      </div>
                    )}
                  </div>
                  <div className="flex flex-wrap gap-1 mt-2">
                    {selectedTargets.map((target) => (
                      <Badge key={target} variant="secondary" className="gap-1">
                        {target}
                        <button
                          onClick={() => setSelectedTargets(selectedTargets.filter(t => t !== target))}
                          className="ml-1 hover:text-red-500"
                        >
                          ×
                        </button>
                      </Badge>
                    ))}
                  </div>
                  {selectedTargets.length === 0 && (
                    <p className="text-xs text-muted-foreground mt-1">
                      Belum ada produk dipilih. Override tidak akan berlaku jika kosong.
                    </p>
                  )}
                </div>
              )}

              <div className="flex gap-2 pt-2">
                <Button
                  size="sm"
                  onClick={handleAddOrUpdateOverride}
                  disabled={
                    saving ||
                    !newOverride.originalItemId ||
                    !newOverride.replacementItemId
                  }
                >
                  {saving ? (
                    <>
                      <Loader2 className="h-3 w-3 animate-spin mr-1" />
                      Menyimpan...
                    </>
                  ) : editId !== null ? (
                    "Update Perubahan"
                  ) : (
                    "Simpan Perubahan"
                  )}
                </Button>
                <Button
                  variant="outline"
                  size="sm"
                  onClick={handleCancelEdit}
                  disabled={saving}
                >
                  Batal
                </Button>
              </div>
            </div>
          )}
          
          {/* Informasi */}
          <Alert>
            <Info className="h-4 w-4" />
            <AlertTitle>Informasi Penting</AlertTitle>
            <AlertDescription className="text-xs">
              <ul className="list-disc list-inside space-y-1">
                <li>
                  ⚠️ Perubahan BOM bersifat <strong>SPESIFIK per ITEM</strong> -
                  hanya item yang dipilih yang akan diganti
                </li>
                <li>
                  ✅ Contoh: Jika Anda mengatur override untuk item "RAW-001"
                  diganti "RAW-002", maka hanya item "RAW-001" yang akan berubah
                  di BOM
                </li>
                <li>
                  ✅ Item lain yang tidak masuk dalam daftar override akan tetap
                  menggunakan item asli dari BOM
                </li>
                <li>
                  🎯 Anda juga dapat membatasi override hanya untuk BOM/produk tertentu
                </li>
                <li>
                  ✅ Perubahan tersimpan di DATABASE, tidak akan hilang meskipun
                  refresh halaman
                </li>
                <li>
                  ✅ Anda dapat mengaktifkan/nonaktifkan perubahan kapan saja
                </li>
                <li>
                  ✅ Satu item hanya bisa memiliki SATU override aktif (jika
                  ingin mengganti, edit atau nonaktifkan yang lama)
                </li>
              </ul>
            </AlertDescription>
          </Alert>
        </div>

        <SheetFooter>
          <Button onClick={onClose} disabled={saving}>
            Tutup
          </Button>
        </SheetFooter>
      </SheetContent>
    </Sheet>
  );
};
// ==================== KOMPONEN UTAMA ====================
export default function ProductionPlanPage() {
  const [orders, setOrders] = useState<ProductionPlan[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [exportLoading, setExportLoading] = useState(false);
  const [committing, setCommitting] = useState<string | null>(null);
  const [committedPOs, setCommittedPOs] = useState<CommittedPO[]>([]);
  const [stockReservations, setStockReservations] = useState<
    StockReservation[]
  >([]);
  const [dateFilter, setDateFilter] = useState({ startDate: "", endDate: "" });
  const [searchQuery, setSearchQuery] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(10);
  const [showCommitted, setShowCommitted] = useState(true);
  const [exportProgress, setExportProgress] = useState({
    visible: false,
    current: 0,
    total: 0,
    message: "",
  });
  // Di dalam komponen ProductionPlanPage, tambahkan state:
  const [previewDialog, setPreviewDialog] = useState({
    open: false,
    data: null as any[] | null,
    fileName: "",
    totalPO: 0,
    totalMaterial: 0,
  });
  const [isMounted, setIsMounted] = useState(false);
  // Di dalam ProductionPlanPage component, tambahkan state
  const [savingToDatabase, setSavingToDatabase] = useState(false);
  const [savedCalculations, setSavedCalculations] = useState<any[]>([]);
  const [showHistoryPanel, setShowHistoryPanel] = useState(false);
  const [bomOverrides, setBomOverrides] = useState<BomOverride[]>([]);
  const [loadingOverrides, setLoadingOverrides] = useState(false);
  // Di bagian state declarations, tambahkan:
  const [bomOverrideDialog, setBomOverrideDialog] = useState({
    open: false,
    originalItemId: "",
    originalItemName: "",
    currentReplacement: "",
  });
 const loadBomOverridesFromDatabase = useCallback(async () => {
   setLoadingOverrides(true);
   try {
     const response = await fetch("/api/ppic/bom-overrides");
     if (!response.ok) {
       throw new Error(`HTTP error! status: ${response.status}`);
     }
     const result = await response.json();

     if (result.success && Array.isArray(result.data)) {
       const overrides = result.data.map((item: any) => ({
         id: item.id,
         originalItemId: item.original_item_id,
         replacementItemId: item.replacement_item_id,
         replacementItemName:
           item.replacement_item_name || item.replacement_item_id,
         replacementItemName2:
           item.replacement_item_name2 || item.replacement_item_id,
         isActive: item.is_active === 1 || item.is_active === true,
         targetKodeBarang: item.target_kode_barang,
         targetKodeBarangs: item.target_kode_barangs
           ? JSON.parse(item.target_kode_barangs)
           : undefined,
       }));
       setBomOverrides(overrides);

       // Simpan ke localStorage sebagai backup
       localStorage.setItem("bomOverrides", JSON.stringify(overrides));
     } else {
       console.warn("Invalid response format:", result);
       const savedOverrides = localStorage.getItem("bomOverrides");
       if (savedOverrides) {
         setBomOverrides(JSON.parse(savedOverrides));
       }
     }
   } catch (error) {
     console.error("Error loading BOM overrides from database:", error);
     const savedOverrides = localStorage.getItem("bomOverrides");
     if (savedOverrides) {
       try {
         setBomOverrides(JSON.parse(savedOverrides));
       } catch (e) {
         console.error("Error loading from localStorage:", e);
       }
     }
   } finally {
     setLoadingOverrides(false);
   }
 }, []);
  // Fungsi untuk menambah BOM override
  // Di fungsi addBomOverride, pastikan response data memiliki id
  const addBomOverride = async (override: BomOverride) => {
    try {
      const response = await fetch("/api/ppic/bom-overrides", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          originalItemId: override.originalItemId,
          replacementItemId: override.replacementItemId,
          replacementItemName: override.replacementItemName,
          replacementItemName2: override.replacementItemName2,
          isActive: override.isActive,
          createdBy: "current_user",
        }),
      });

      const result = await response.json();
      if (result.success) {
        setBomOverrides((prev) => [
          ...prev,
          {
            ...override,
            id: result.data?.id || Date.now(), // Pastikan ada id
          },
        ]);
        showToast(
          `✅ Perubahan BOM untuk ${override.originalItemId} berhasil ditambahkan`,
          "success",
        );
        await refreshAllData();
      } else {
        throw new Error(result.error);
      }
    } catch (error) {
      console.error("Error adding BOM override:", error);
      showToast(`❌ Gagal menambahkan perubahan BOM`, "error");
    }
  };

  // Fungsi untuk update BOM override
  const updateBomOverride = async (index: number, override: BomOverride) => {
    try {
      const response = await fetch("/api/ppic/bom-overrides", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          id: override.id,
          originalItemId: override.originalItemId,
          replacementItemId: override.replacementItemId,
          replacementItemName: override.replacementItemName,
          replacementItemName2: override.replacementItemName2,
          isActive: override.isActive,
        }),
      });

      const result = await response.json();
      if (result.success) {
        setBomOverrides((prev) =>
          prev.map((o, i) => (i === index ? override : o)),
        );
        showToast(
          `✅ Perubahan BOM untuk ${override.originalItemId} berhasil diupdate`,
          "success",
        );
        await refreshAllData();
      } else {
        throw new Error(result.error);
      }
    } catch (error) {
      console.error("Error updating BOM override:", error);
      showToast(`❌ Gagal mengupdate perubahan BOM`, "error");
    }
  };

  // Fungsi untuk delete BOM override
  const deleteBomOverride = async (index: number, id: number) => {
    try {
      const response = await fetch(`/api/ppic/bom-overrides?id=${id}`, {
        method: "DELETE",
      });

      const result = await response.json();
      if (result.success) {
        const deletedOverride = bomOverrides[index];
        setBomOverrides((prev) => prev.filter((_, i) => i !== index));
        showToast(
          `✅ Perubahan BOM untuk ${deletedOverride.originalItemId} berhasil dihapus`,
          "success",
        );
        await refreshAllData();
      } else {
        throw new Error(result.error);
      }
    } catch (error) {
      console.error("Error deleting BOM override:", error);
      showToast(`❌ Gagal menghapus perubahan BOM`, "error");
    }
  };

  // Fungsi untuk toggle active status
  const toggleBomOverride = async (
    index: number,
    id: number,
    isActive: boolean,
  ) => {
    try {
      const response = await fetch("/api/ppic/bom-overrides", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id, isActive }),
      });

      const result = await response.json();
      if (result.success) {
        setBomOverrides((prev) =>
          prev.map((o, i) => (i === index ? { ...o, isActive } : o)),
        );
        showToast(
          `✅ Perubahan BOM ${isActive ? "diaktifkan" : "dinonaktifkan"}`,
          "success",
        );
        await refreshAllData();
      } else {
        throw new Error(result.error);
      }
    } catch (error) {
      console.error("Error toggling BOM override:", error);
      showToast(`❌ Gagal mengubah status BOM override`, "error");
    }
  };

  // Load BOM overrides saat komponen mount
  useEffect(() => {
    loadBomOverridesFromDatabase();
  }, [loadBomOverridesFromDatabase]);

  // Simpan ke localStorage sebagai backup saat overrides berubah (opsional)
  useEffect(() => {
    localStorage.setItem("bomOverrides", JSON.stringify(bomOverrides));
  }, [bomOverrides]);
  // Fungsi untuk menerapkan BOM override pada BOM (HANYA UNTUK BOM TERTENTU)
  const applyBomOverrides = (
    bomFlat: BomItem[],
    kodeBarang?: string,
  ): BomItem[] => {
    if (!bomOverrides.length) return bomFlat;

    const activeOverrides = bomOverrides.filter((o) => o.isActive);
    if (!activeOverrides.length) return bomFlat;

    // Filter override berdasarkan target kode barang
    let applicableOverrides = activeOverrides;

    if (kodeBarang) {
      // Hanya ambil override yang:
      // 1. Tidak memiliki target (global) ATAU
      // 2. Targetnya match dengan kodeBarang saat ini
      applicableOverrides = activeOverrides.filter((override) => {
        // Jika tidak ada target, berarti berlaku untuk semua (global)
        if (
          !override.targetKodeBarang &&
          (!override.targetKodeBarangs ||
            override.targetKodeBarangs.length === 0)
        ) {
          return true;
        }
        // Jika ada targetKodeBarang spesifik
        if (
          override.targetKodeBarang &&
          override.targetKodeBarang === kodeBarang
        ) {
          return true;
        }
        // Jika ada targetKodeBarangs (array)
        if (
          override.targetKodeBarangs &&
          override.targetKodeBarangs.includes(kodeBarang)
        ) {
          return true;
        }
        return false;
      });
    }

    if (applicableOverrides.length === 0) return bomFlat;

    // Buat map untuk override berdasarkan originalItemId
    const overrideMap = new Map<string, BomOverride>();
    applicableOverrides.forEach((override) => {
      const normalizedOriginalId = normalizeItemId(override.originalItemId);
      overrideMap.set(normalizedOriginalId, override);
    });

    console.log(
      `Active BOM Overrides for ${kodeBarang || "ALL"}:`,
      applicableOverrides.length,
    );
    console.log(
      "Overrides:",
      applicableOverrides.map((o) => ({
        original: o.originalItemId,
        replacement: o.replacementItemId,
        target: o.targetKodeBarang || o.targetKodeBarangs || "ALL",
      })),
    );

    // Hanya ganti item yang MATCH dengan originalItemId di override
    let overrideCount = 0;
    const result = bomFlat.map((item) => {
      const normalizedId = normalizeItemId(item.ItemID);
      const override = overrideMap.get(normalizedId);

      if (override) {
        overrideCount++;
        console.log(
          `OVERRIDE for ${kodeBarang || "ALL"}: ${item.ItemID} -> ${override.replacementItemId}`,
        );
        return {
          ...item,
          ItemID: override.replacementItemId,
          ItemName: override.replacementItemName,
          ItemName2: override.replacementItemName2,
          _originalItemId: item.ItemID,
          _isOverridden: true,
        };
      }
      return item;
    });

    console.log(
      `Total override applied for ${kodeBarang || "ALL"}: ${overrideCount}`,
    );
    return result;
  };
  // Load BOM overrides dari localStorage saat komponen mount
  useEffect(() => {
    const savedOverrides = localStorage.getItem("bomOverrides");
    if (savedOverrides) {
      try {
        setBomOverrides(JSON.parse(savedOverrides));
      } catch (e) {
        console.error("Error loading BOM overrides:", e);
      }
    }
  }, []);

  // Simpan BOM overrides ke localStorage setiap kali berubah
  useEffect(() => {
    localStorage.setItem("bomOverrides", JSON.stringify(bomOverrides));
  }, [bomOverrides]);
  // Fungsi untuk menghapus perhitungan
  const deleteCalculation = async (calculationId: string) => {
    try {
      const response = await fetch(
        `/api/ppic/save-material-requirements?calculation_id=${calculationId}`,
        {
          method: "DELETE",
        },
      );
      const result = await response.json();
      if (result.success) {
        showToast(
          `✅ Perhitungan ${calculationId} berhasil dihapus`,
          "success",
        );
        await loadSavedCalculations();
      } else {
        throw new Error(result.error || "Failed to delete");
      }
    } catch (error) {
      console.error("Error deleting calculation:", error);
      showToast(`❌ Gagal menghapus: ${error}`, "error");
    }
  };

  // Fungsi untuk memuat perhitungan yang tersimpan ke preview
  const loadCalculationToPreview = (calculation: any) => {
    setPreviewDialog({
      open: true,
      data: calculation.material_data,
      fileName: calculation.calculation_id,
      totalPO: calculation.total_po,
      totalMaterial: calculation.total_materials,
    });
  };
  // Toast notification state
  const [toast, setToast] = useState<{
    message: string;
    type: "success" | "error" | "info" | "loading";
    id: number;
  } | null>(null);

  const showToast = (
    message: string,
    type: "success" | "error" | "info" | "loading",
  ) => {
    setToast({ message, type, id: Date.now() });
  };

  const hideToast = () => {
    setToast(null);
  };

  // Fix hydration mismatch
  useEffect(() => {
    setIsMounted(true);
  }, []);

  const filteredOrders = useMemo(() => {
    if (!searchQuery.trim()) return orders;
    const query = searchQuery.toLowerCase();
    return orders.filter(
      (order) =>
        order.order.No_SPK?.toLowerCase().includes(query) ||
        order.order.Nama_PO?.toLowerCase().includes(query) ||
        order.order.Kode_Barang?.toLowerCase().includes(query),
    );
  }, [orders, searchQuery]);

  const getAllKodeBarang = (kodeBarang: string): string[] => {
    if (kodeBarang.includes(" | ")) return kodeBarang.split(" | ");
    return [kodeBarang];
  };

  const combineDuplicatePOs = (
    orders: ProductionOrder[],
  ): ProductionOrder[] => {
    const poMap = new Map<string, ProductionOrder>();
    orders.forEach((order) => {
      const existing = poMap.get(order.No_SPK);
      if (existing) {
        if (!existing.combinedItems)
          existing.combinedItems = [{ ...existing, isCombined: false }];
        existing.combinedItems.push({ ...order, isCombined: true });
        if (existing.Nama_PO !== order.Nama_PO)
          existing.Nama_PO = `${existing.Nama_PO} | ${order.Nama_PO}`;
        if (existing.Kode_Barang !== order.Kode_Barang)
          existing.Kode_Barang = `${existing.Kode_Barang} | ${order.Kode_Barang}`;
        if (new Date(order.Tanggal_Order) > new Date(existing.Tanggal_Order))
          existing.Tanggal_Order = order.Tanggal_Order;
      } else {
        poMap.set(order.No_SPK, {
          ...order,
          combinedItems: [{ ...order, isCombined: false }],
        });
      }
    });
    return Array.from(poMap.values());
  };

  const combineBoms = (boms: {
    [kodeBarang: string]: { flat: BomItem[]; tree: BomItem[] };
  }) => {
    const combinedFlat: BomItem[] = [];
    const itemMap = new Map<string, BomItem>();
    Object.values(boms).forEach((bom) => {
      bom.flat.forEach((item) => {
        const normalizedId = normalizeItemId(item.ItemID);
        const existing = itemMap.get(normalizedId);
        if (existing) existing.Qty += item.Qty;
        else {
          itemMap.set(normalizedId, { ...item });
          combinedFlat.push({ ...item });
        }
      });
    });
    const combinedTree: BomItem[] = [];
    Object.values(boms).forEach((bom) => {
      if (bom.tree) combinedTree.push(...bom.tree);
    });
    return { flat: combinedFlat, tree: combinedTree, combinedBoms: boms };
  };

  const loadCommittedPOs = async () => {
    try {
      console.log("🔄 Loading committed POs from API...");
      const response = await fetch("/api/ppic/committed-pos", {
        cache: "no-store",
        headers: { "Cache-Control": "no-cache" },
      });

      const responseText = await response.text();
      console.log("Raw committed POs response:", responseText);

      let result;
      try {
        result = JSON.parse(responseText);
      } catch (e) {
        console.error("JSON parse error:", e);
        return [];
      }

      if (result.success) {
        const newCommittedPOs = result.data.committedPOs || [];
        const newReservations = result.data.reservations || [];

        console.log(`✅ Loaded ${newCommittedPOs.length} committed POs`);
        console.log(
          "Committed PO details:",
          newCommittedPOs.map((po: any) => ({
            noSPK: po.noSPK,
            status: po.status,
            CommitID: po.CommitID,
          })),
        );

        setCommittedPOs(newCommittedPOs);
        setStockReservations(newReservations);

        return newCommittedPOs;
      }
      return [];
    } catch (error) {
      console.error("Error loading committed POs:", error);
      return [];
    }
  };

  const fetchOrders = async (
    startDate?: string,
    endDate?: string,
    existingCommittedPOs?: CommittedPO[],
    showCommittedParam?: boolean, // Parameter baru
  ) => {
    try {
      setLoading(true);
      setError("");
      let url = "/api/ppic";
      const params = new URLSearchParams();
      if (startDate) params.append("startDate", startDate);
      if (endDate) params.append("endDate", endDate);
      if (params.toString()) url += `?${params.toString()}`;

      const response = await axios.get<ProductionOrder[]>(url);
      const activeOrders = response.data.filter((order) => !order.Completed);
      const combinedOrders = combineDuplicatePOs(activeOrders);

      const committedList = existingCommittedPOs || committedPOs;
      const committedSPKs = new Set(
        committedList
          .filter((po) => po.status === "COMMITTED")
          .map((po) => po.noSPK),
      );

      // 🔥 Gunakan showCommittedParam atau showCommitted dari state
      const shouldShowCommitted =
        showCommittedParam !== undefined ? showCommittedParam : showCommitted;

      console.log(`showCommitted = ${shouldShowCommitted}`); // Debug log

      let filteredOrdersData;
      if (shouldShowCommitted) {
        // Tampilkan SEMUA PO (termasuk yang sudah di-commit)
        filteredOrdersData = combinedOrders;
        console.log(
          `Menampilkan semua PO (termasuk committed): ${combinedOrders.length} total`,
        );
      } else {
        // Hanya PO yang belum di-commit
        filteredOrdersData = combinedOrders.filter(
          (order) => !committedSPKs.has(order.No_SPK),
        );
        console.log(
          `Menampilkan PO aktif: ${filteredOrdersData.length} dari ${combinedOrders.length} total`,
        );
      }

      const productionPlans: ProductionPlan[] = filteredOrdersData.map(
        (order) => {
          // Tandai PO yang sudah di-commit
          const isCommitted = committedSPKs.has(order.No_SPK);
          const commitInfo = committedList.find(
            (po) => po.noSPK === order.No_SPK && po.status === "COMMITTED",
          );

          return {
            order: {
              ...order,
              QTY: order.QTY || 0,
              originalOrders: order.combinedItems || [order],
            },
            selected: false,
            loadingBom: false,
            committed: isCommitted,
            CommitID: commitInfo?.CommitID,
            bom: undefined,
            stock: undefined,
            stockLastUpdated: undefined,
          };
        },
      );

      console.log(`📊 Total PO ditampilkan: ${productionPlans.length}`);
      setOrders(productionPlans);
      setCurrentPage(1);
    } catch (err) {
      setError(
        "Gagal mengambil data produksi: " +
          (err instanceof Error ? err.message : "Unknown error"),
      );
    } finally {
      setLoading(false);
    }
  };
  // Perbaiki fungsi refreshAllData
  const refreshAllData = useCallback(async () => {
    showToast("Memuat ulang data...", "loading");
    setLoading(true);
    try {
      const latestCommittedPOs = await loadCommittedPOs();
      // Kirim showCommitted ke fetchOrders
      await fetchOrders(
        dateFilter.startDate,
        dateFilter.endDate,
        latestCommittedPOs,
        showCommitted,
      );
      showToast("Data berhasil dimuat ulang", "success");
    } catch (error) {
      console.error("Error refreshing data:", error);
      showToast("Gagal memuat ulang data", "error");
    } finally {
      setLoading(false);
    }
  }, [
    dateFilter.startDate,
    dateFilter.endDate,
    showCommitted,
    loadCommittedPOs,
  ]); // Tambahkan loadCommittedPOs ke dependensi
  // ==================== FUNGSI LOAD BOM UNTUK COMMIT ====================
  const loadBomForCommit = async (
    plan: ProductionPlan,
  ): Promise<ProductionPlan> => {
    if (plan.bom && plan.stock) {
      return plan;
    }

    const allKodeBarang = getAllKodeBarang(plan.order.Kode_Barang);
    const combinedBoms: {
      [kodeBarang: string]: { flat: BomItem[]; tree: BomItem[] };
    } = {};
    let allItemIds: string[] = [];

    for (const kb of allKodeBarang) {
      try {
        const bomResponse = await axios.get(
          `/api/bom/ppic?itemid=${encodeURIComponent(kb)}`,
        );

        // 🔥 AMBIL BOM FLAT ASLI
        const bomFlat = bomResponse.data.flat;

        // 🔥 TERAPKAN BOM OVERRIDE (GANTI ITEM TERTENTU)
        const bomFlatWithOverrides = applyBomOverrides(bomFlat, kb);

        // 🔥 BUILD TREE STRUCTURE DENGAN BOM YANG SUDAH DIOVERRIDE
        const treeStructure = buildTreeStructure(bomFlatWithOverrides);

        // 🔥 SIMPAN KE combinedBoms
        combinedBoms[kb] = {
          flat: bomFlatWithOverrides,
          tree: treeStructure,
        };

        allItemIds.push(
          ...bomFlatWithOverrides.map((item: BomItem) =>
            normalizeItemId(item.ItemID),
          ),
        );
      } catch (err) {
        console.error(`Gagal load BOM untuk ${kb}:`, err);
      }
    }

    allItemIds = Array.from(new Set(allItemIds));
    const stockData = await fetchStockForItemsWithCommitment(
      allItemIds,
      plan.order.Tanggal_Order,
    );
    const finalBom = combineBoms(combinedBoms);

    return {
      ...plan,
      bom: finalBom,
      stock: stockData,
      stockLastUpdated: new Date().toISOString(),
    };
  };

  // ==================== FUNGSI COMMIT PO (DIPERBAIKI) ====================
  const commitPO = async (index: number) => {
    // Dapatkan data dari paginatedOrders berdasarkan index
    const plan = paginatedOrders[index];
    if (!plan) return;

    const spk = plan.order.No_SPK;

    // Cari index asli di orders
    const originalIndex = orders.findIndex((o) => o.order.No_SPK === spk);
    if (originalIndex === -1) {
      showToast(`PO ${spk} tidak ditemukan`, "error");
      return;
    }

    showToast(`Mempersiapkan commit PO ${spk}...`, "loading");
    setCommitting(spk);

    try {
      let updatedPlan = plan;

      // Load BOM jika belum ada
      if (!plan.bom || !plan.stock) {
        setOrders((prev) =>
          prev.map((item, i) =>
            i === originalIndex ? { ...item, loadingBom: true } : item,
          ),
        );
        showToast(`Memuat BOM untuk PO ${spk}...`, "loading");
        updatedPlan = await loadBomForCommit(plan);
        setOrders((prev) =>
          prev.map((item, i) => (i === originalIndex ? updatedPlan : item)),
        );
      }

      if (!updatedPlan.bom || !updatedPlan.stock) {
        throw new Error("BOM atau Stock tidak tersedia");
      }

      const isCombined =
        updatedPlan.order.combinedItems &&
        updatedPlan.order.combinedItems.length > 1;

      let materialNeeds;
      if (isCombined && updatedPlan.order.combinedItems) {
        materialNeeds = calculateMaterialNeedsForCombinedPO(
          updatedPlan.bom,
          updatedPlan.order.combinedItems,
          updatedPlan.stock,
        );
      } else {
        materialNeeds = calculateMaterialNeeds(
          updatedPlan.bom.flat,
          updatedPlan.order.QTY,
          updatedPlan.stock,
        );
      }

      const materialUsage = materialNeeds.items.map((item: any) => ({
        itemId: item.ItemID,
        itemName: item.ItemName,
        itemName2: item.ItemName2 || "",
        qtyPerUnit: item.Qty,
        totalNeeded: item.needed,
        stockBefore: item.availableStock,
        stockAfter: item.availableStock - item.needed,
        qtyUsed: item.needed,
        departemen: item.Departemen,
        level: Number(item.Level),
      }));

      showToast(`Menyimpan commit PO ${spk} ke database...`, "loading");

      const commitResponse = await fetch("/api/ppic/commit-po", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          noSPK: spk,
          kodeBarang: updatedPlan.order.Kode_Barang,
          namaPO: updatedPlan.order.Nama_PO,
          qty: updatedPlan.order.QTY,
          userID: "current_user",
          materialUsage: materialUsage,
          orderDate: updatedPlan.order.Tanggal_Order,
        }),
      });

      const responseText = await commitResponse.text();
      console.log("Commit response:", responseText);

      let commitResult = null;
      let commitID = null;

      try {
        commitResult = JSON.parse(responseText);
        commitID =
          commitResult.CommitID ||
          commitResult.commitID ||
          commitResult.data?.CommitID;
      } catch (e) {
        console.error("Failed to parse JSON response:", e);
      }

      if (commitResponse.ok) {
        const commitIDDisplay = commitID || "berhasil disimpan";
        showToast(
          `✅ PO ${spk} berhasil di-commit! Commit ID: ${commitIDDisplay}`,
          "success",
        );

        // Hapus PO dari state orders
        setOrders((prev) => prev.filter((order) => order.order.No_SPK !== spk));
        await loadCommittedPOs();

        // Reset halaman jika perlu
        if (filteredOrders.length <= 1 && currentPage > 1) {
          setCurrentPage(currentPage - 1);
        }
      } else {
        throw new Error(commitResult?.error || `HTTP ${commitResponse.status}`);
      }
    } catch (error: any) {
      console.error(`Gagal commit PO:`, error);
      showToast(
        `❌ Gagal commit PO ${spk}: ${error.message || "Unknown error"}`,
        "error",
      );
      setOrders((prev) =>
        prev.map((item, i) =>
          i === originalIndex ? { ...item, loadingBom: false } : item,
        ),
      );
    } finally {
      setCommitting(null);
    }
  };
  // ==================== FUNGSI UNCOMMIT (DIPERBAIKI) ====================
  const uncommitPO = async (noSPK: string) => {
    if (
      !confirm(
        `⚠️ PERINGATAN: Anda akan meng-uncommit PO ${noSPK}\n\nStok yang sudah di-reserve akan dikembalikan.\n\nPO akan muncul kembali di daftar aktif.\n\nApakah Anda yakin?`,
      )
    )
      return;

    showToast(`Memproses uncommit PO ${noSPK}...`, "loading");
    setCommitting(noSPK);

    try {
      const response = await fetch("/api/ppic/uncommit-po", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ noSPK, userID: "system" }),
      });

      // Baca response sebagai text terlebih dahulu
      const responseText = await response.text();
      console.log("Uncommit response:", responseText);

      // Coba parse JSON jika memungkinkan
      let result = null;
      let isSuccess = false;

      try {
        result = JSON.parse(responseText);
        isSuccess = result.success === true;
      } catch (e) {
        // Jika bukan JSON, cek status code
        console.log("Response is not JSON, checking status code");
        isSuccess = response.ok;
      }

      // Jika response status OK atau success true, anggap berhasil
      if (isSuccess || response.ok) {
        showToast(
          `✅ PO ${noSPK} berhasil di-uncommit! PO akan muncul kembali di daftar aktif.`,
          "success",
        );
        await refreshAllData();
      } else {
        // Tampilkan error tapi jangan throw jika sebenarnya berhasil
        const errorMsg = result?.error || "Unknown error";
        console.warn(`Uncommit warning: ${errorMsg}`);
        // Tetap refresh data karena mungkin tetap berhasil meskipun response error
        showToast(
          `ℹ️ PO ${noSPK} diproses, silahkan refresh manual jika belum muncul.`,
          "info",
        );
        await refreshAllData();
      }
    } catch (error: any) {
      console.error(`Gagal uncommit PO:`, error);
      // Jangan tampilkan error jika sebenarnya data sudah berubah
      showToast(
        `ℹ️ PO ${noSPK} sedang diproses, silahkan cek daftar PO.`,
        "info",
      );
      await refreshAllData();
    } finally {
      setCommitting(null);
    }
  };
  // ==================== FUNGSI LOAD DATA COMMITTED PO ====================
  // Tambahkan fungsi ini SEBELUM loadBomAndStockForExport
  const loadCommittedOrderData = async (
    committedPO: ProductionPlan,
  ): Promise<ProductionPlan> => {
    try {
      console.log(`Loading data for committed PO: ${committedPO.order.No_SPK}`);

      const allKodeBarang = getAllKodeBarang(committedPO.order.Kode_Barang);
      const combinedBoms: {
        [kodeBarang: string]: { flat: BomItem[]; tree: BomItem[] };
      } = {};
      let allItemIds: string[] = [];

      for (const kb of allKodeBarang) {
        try {
          const bomResponse = await axios.get(
            `/api/bom/ppic?itemid=${encodeURIComponent(kb)}`,
          );

          // 🔥 TERAPKAN BOM OVERRIDE DI SINI
          const bomFlat = bomResponse.data.flat;
          const bomFlatWithOverrides = applyBomOverrides(bomFlat, kb);
          const treeStructure = buildTreeStructure(bomFlatWithOverrides);

          combinedBoms[kb] = {
            flat: bomFlatWithOverrides,
            tree: treeStructure,
          };
          allItemIds.push(
            ...bomFlatWithOverrides.map((item: BomItem) =>
              normalizeItemId(item.ItemID),
            ),
          );
        } catch (err) {
          console.error(`Gagal load BOM untuk ${kb}:`, err);
        }
      }

      allItemIds = Array.from(new Set(allItemIds));
      const stockData = await fetchStockForItemsWithCommitment(
        allItemIds,
        committedPO.order.Tanggal_Order,
      );
      const finalBom = combineBoms(combinedBoms);

      return {
        ...committedPO,
        bom: finalBom,
        stock: stockData,
        stockLastUpdated: new Date().toISOString(),
      };
    } catch (error) {
      console.error(
        `Error loading data for committed PO ${committedPO.order.No_SPK}:`,
        error,
      );
      return committedPO;
    }
  };

  // ==================== FUNGSI LOAD BOM & STOCK UNTUK EXPORT ====================
  const loadBomAndStockForExport = async (
    selectedOrders: ProductionPlan[],
  ): Promise<ProductionPlan[]> => {
    const updatedOrders = [...selectedOrders];
    const totalOrders = updatedOrders.length;

    for (let i = 0; i < updatedOrders.length; i++) {
      const order = updatedOrders[i];

      if (order.bom && order.stock) {
        setExportProgress((prev) => ({
          ...prev,
          current: Math.floor(((i + 1) / totalOrders) * 50),
          message: `Menggunakan BOM yang sudah ada untuk ${order.order.No_SPK} (${i + 1}/${totalOrders})...`,
        }));
        continue;
      }

      if (order.committed) {
        setExportProgress({
          visible: true,
          current: Math.floor((i / totalOrders) * 50),
          total: 100,
          message: `Memuat data untuk committed PO ${order.order.No_SPK} (${i + 1}/${totalOrders})...`,
        });

        const loadedOrder = await loadCommittedOrderData(order);
        updatedOrders[i] = loadedOrder;
        continue;
      }

      console.log(`Loading BOM for active PO ${order.order.No_SPK}`);
      setExportProgress({
        visible: true,
        current: Math.floor((i / totalOrders) * 50),
        total: 100,
        message: `Memuat BOM dan stok untuk ${order.order.No_SPK} (${i + 1}/${totalOrders})...`,
      });

      const allKodeBarang = getAllKodeBarang(order.order.Kode_Barang);
      const combinedBoms: {
        [kodeBarang: string]: { flat: BomItem[]; tree: BomItem[] };
      } = {};
      let allItemIds: string[] = [];

      for (const kb of allKodeBarang) {
        try {
          const bomResponse = await axios.get(
            `/api/bom/ppic?itemid=${encodeURIComponent(kb)}`,
            { timeout: 30000 },
          );

          // 🔥 TERAPKAN BOM OVERRIDE DI SINI
          const bomFlat = bomResponse.data.flat;
          const bomFlatWithOverrides = applyBomOverrides(bomFlat, kb);
          const treeStructure = buildTreeStructure(bomFlatWithOverrides);

          combinedBoms[kb] = {
            flat: bomFlatWithOverrides,
            tree: treeStructure,
          };
          allItemIds.push(
            ...bomFlatWithOverrides.map((item: BomItem) =>
              normalizeItemId(item.ItemID),
            ),
          );
        } catch (err) {
          console.error(`Gagal load BOM untuk ${kb}:`, err);
        }
      }

      allItemIds = Array.from(new Set(allItemIds));
      const stockData = await fetchStockForItemsWithCommitment(
        allItemIds,
        order.order.Tanggal_Order,
      );
      const finalBom = combineBoms(combinedBoms);

      updatedOrders[i] = {
        ...order,
        bom: finalBom,
        stock: stockData,
        stockLastUpdated: new Date().toISOString(),
      };

      await new Promise((resolve) => setTimeout(resolve, 50));
    }

    return updatedOrders;
  };

  // ==================== FUNGSI EXPORT (LENGKAP DENGAN PERBAIKAN) ====================
  // ==================== FUNGSI CALCULATE ACCUMULATED QTY (DIPERBAIKI) ====================
  const calculateAccumulatedQty = (
    item: BomItem,
    flatBom: BomItem[],
    visited: Set<string> = new Set(),
  ): number => {
    const currentLevel = Number(item.Level);

    // Level 1: accumulated = qty per unit
    if (currentLevel === 1) {
      return item.Qty;
    }

    // Cegah infinite recursion
    const itemId = normalizeItemId(item.ItemID);
    if (visited.has(itemId)) {
      console.warn(`Circular reference detected for item: ${itemId}`);
      return item.Qty;
    }
    visited.add(itemId);

    // Cari parent (item dengan level lebih rendah 1)
    const parent = flatBom.find((p) => {
      const parentLevel = Number(p.Level);
      return parentLevel === currentLevel - 1 && p.ItemID !== item.ItemID;
    });

    if (!parent) {
      console.warn(
        `Parent tidak ditemukan untuk item ${item.ItemID} (Level ${currentLevel})`,
      );
      visited.delete(itemId);
      return item.Qty;
    }

    // Hitung accumulated parent terlebih dahulu (rekursif dengan visited set)
    const parentAccumulated = calculateAccumulatedQty(parent, flatBom, visited);

    // Hapus dari visited setelah selesai
    visited.delete(itemId);

    // Accumulated item = Qty item × (parentAccumulated / parent.Qty)
    if (parent.Qty === 0) {
      console.warn(`Parent Qty is zero for item ${parent.ItemID}`);
      return item.Qty;
    }

    const result = (item.Qty * parentAccumulated) / parent.Qty;

    // Validasi hasil
    if (isNaN(result) || !isFinite(result)) {
      console.warn(
        `Invalid accumulated qty for item ${item.ItemID}: ${result}`,
      );
      return item.Qty;
    }

    return result;
  };

  // Fungsi alternatif dengan pendekatan iteratif (lebih aman)
  const calculateAccumulatedQtyIterative = (
    item: BomItem,
    flatBom: BomItem[],
  ): number => {
    // Build map untuk akses cepat
    const itemMap = new Map<string, BomItem>();
    flatBom.forEach((b) => {
      itemMap.set(normalizeItemId(b.ItemID), b);
    });

    // Build parent-child relationship
    const childrenMap = new Map<string, BomItem[]>();
    flatBom.forEach((b) => {
      const parentId = b.ParentItemID ? normalizeItemId(b.ParentItemID) : null;
      if (parentId && parentId !== normalizeItemId(b.ItemID)) {
        if (!childrenMap.has(parentId)) {
          childrenMap.set(parentId, []);
        }
        childrenMap.get(parentId)!.push(b);
      }
    });

    // Hitung accumulated qty dari atas ke bawah (iterative)
    const accumulatedMap = new Map<string, number>();

    // Level 1 items
    const level1Items = flatBom.filter((b) => Number(b.Level) === 1);
    for (const level1Item of level1Items) {
      accumulatedMap.set(normalizeItemId(level1Item.ItemID), level1Item.Qty);
    }

    // Process level by level
    let currentLevel = 1;
    let maxLevel = Math.max(...flatBom.map((b) => Number(b.Level)), 0);

    while (currentLevel < maxLevel) {
      const nextLevelItems = flatBom.filter(
        (b) => Number(b.Level) === currentLevel + 1,
      );

      for (const nextItem of nextLevelItems) {
        const parentId = nextItem.ParentItemID
          ? normalizeItemId(nextItem.ParentItemID)
          : null;
        if (parentId && accumulatedMap.has(parentId)) {
          const parentAccumulated = accumulatedMap.get(parentId)!;
          const parent = itemMap.get(parentId);
          if (parent && parent.Qty > 0) {
            const accumulated = (nextItem.Qty * parentAccumulated) / parent.Qty;
            accumulatedMap.set(normalizeItemId(nextItem.ItemID), accumulated);
          } else {
            accumulatedMap.set(normalizeItemId(nextItem.ItemID), nextItem.Qty);
          }
        } else {
          accumulatedMap.set(normalizeItemId(nextItem.ItemID), nextItem.Qty);
        }
      }
      currentLevel++;
    }

    return accumulatedMap.get(normalizeItemId(item.ItemID)) || item.Qty;
  };

  // Gunakan fungsi ini di export (pilih salah satu)
  // Untuk amannya, gunakan fungsi iteratif
  const getAccumulatedQty = calculateAccumulatedQtyIterative;
  // Tambahkan fungsi untuk memuat data committed PO dari database

  const fetchMasterDataForItems = async (
    itemIds: string[],
  ): Promise<Map<string, any>> => {
    const masterDataMap = new Map<string, any>();
    const uniqueItemIds = Array.from(new Set(itemIds));
    for (const itemId of uniqueItemIds) {
      try {
        const response = await axios.get(
          `/api/master?check=true&itemId=${encodeURIComponent(itemId)}`,
        );
        if (
          response.data.success &&
          response.data.exists &&
          response.data.data
        ) {
          const masterItem = response.data.data;
          masterDataMap.set(normalizeItemId(itemId), {
            spec: masterItem.Spec || "-",
            warna: masterItem.warna || "-",
            bahan: masterItem.bahan || "-",
          });
        } else {
          masterDataMap.set(normalizeItemId(itemId), {
            spec: "-",
            warna: "-",
            bahan: "-",
          });
        }
      } catch (error) {
        masterDataMap.set(normalizeItemId(itemId), {
          spec: "-",
          warna: "-",
          bahan: "-",
        });
      }
    }
    return masterDataMap;
  };

  // Fungsi untuk membersihkan nama sheet
  const sanitizeSheetName = (name: string): string => {
    let cleanName = name.replace(/[\\/*?:\[\]]/g, "");
    cleanName = cleanName.replace(/\|/g, "");
    if (cleanName.length > 31) cleanName = cleanName.substring(0, 31);
    if (cleanName.trim() === "") cleanName = "DEPARTEMEN";
    return cleanName;
  };

  // Fungsi untuk mendapatkan nama sheet unik
  const getUniqueSheetName = (wb: XLSX.WorkBook, baseName: string): string => {
    let sheetName = sanitizeSheetName(baseName);
    let counter = 1;
    let uniqueName = sheetName;
    while (wb.SheetNames.includes(uniqueName)) {
      uniqueName = `${sheetName}_${counter}`;
      counter++;
    }
    return uniqueName;
  };

  // ==================== FUNGSI EXPORT (DENGAN AUTO SAVE KE DATABASE) ====================
  const exportSelectedToExcel = async (): Promise<void> => {
    try {
      setExportLoading(true);
      setExportProgress({
        visible: true,
        current: 0,
        total: 100,
        message: "Memulai proses export...",
      });
      showToast("Memulai export data...", "loading");

      const today = new Date().toISOString().split("T")[0];
      const selectedOrders = filteredOrders.filter((order) => order.selected);

      const exportingSPKs = new Set(
        selectedOrders.map((order) => order.order.No_SPK),
      );
      console.log("Exporting SPKs:", Array.from(exportingSPKs));

      if (selectedOrders.length === 0) {
        alert("Tidak ada PO yang dipilih untuk di-export!");
        setExportLoading(false);
        setExportProgress({
          visible: false,
          current: 0,
          total: 0,
          message: "",
        });
        return;
      }

      // ==================== BUAT NAMA FILE BERDASARKAN PO ====================
      let fileName = "";

      if (selectedOrders.length === 1) {
        const singlePO = selectedOrders[0];
        const poName = singlePO.order.Nama_PO || singlePO.order.No_SPK;
        const cleanFileName = poName
          .replace(/[\\/*?:"<>|]/g, "")
          .replace(/\s+/g, "_")
          .substring(0, 50);
        fileName = `${cleanFileName}_${today}.xlsx`;
      } else {
        const firstPO = selectedOrders[0];
        const firstPOName = firstPO.order.Nama_PO || firstPO.order.No_SPK;
        const cleanFirstPO = firstPOName
          .replace(/[\\/*?:"<>|]/g, "")
          .replace(/\s+/g, "_")
          .substring(0, 40);
        fileName = `${cleanFirstPO}_dan_${selectedOrders.length - 1}_lainnya_${today}.xlsx`;
      }

      setExportProgress({
        visible: true,
        current: 10,
        total: 100,
        message: `Memuat BOM & stok untuk ${selectedOrders.length} PO...`,
      });

      const ordersWithBom = await loadBomAndStockForExport(selectedOrders);

      setExportProgress({
        visible: true,
        current: 30,
        total: 100,
        message: "Mengambil data master...",
      });

      const allMaterialIds: string[] = [];
      for (const order of ordersWithBom) {
        if (order.bom?.flat) {
          order.bom.flat.forEach((item: BomItem) => {
            if (item.ItemID && Number(item.Level) > 0)
              allMaterialIds.push(normalizeItemId(item.ItemID));
          });
        }
      }
      const masterDataMap = await fetchMasterDataForItems(allMaterialIds);

      setExportProgress({
        visible: true,
        current: 50,
        total: 100,
        message: "Membuat file Excel...",
      });

      // ==================== BUAT MATERIAL DATA UNTUK SAVE KE DATABASE ====================
      // Ini adalah data yang akan disimpan ke database
      let materialDataForDatabase: any[] = [];

      // ==================== FUNGSI CALCULATE ACCUMULATED QTY UNTUK MATERIAL ====================
      const calculateAccumulatedQtyForMaterial = (
        flatBom: BomItem[],
      ): Map<string, number> => {
        const cache = new Map<string, number>();
        const itemMap = new Map<string, BomItem>();

        for (const item of flatBom) {
          itemMap.set(normalizeItemId(item.ItemID), item);
        }

        for (const item of flatBom) {
          const itemId = normalizeItemId(item.ItemID);
          const level = Number(item.Level);

          if (level === 1) {
            cache.set(itemId, item.Qty);
          } else {
            let parent: BomItem | undefined = undefined;

            if (item.ParentItemID) {
              parent = itemMap.get(normalizeItemId(item.ParentItemID));
            }

            if (!parent) {
              parent = flatBom.find((p) => Number(p.Level) === level - 1);
            }

            if (parent) {
              const parentId = normalizeItemId(parent.ItemID);
              const parentAccumulated = cache.get(parentId);
              if (parentAccumulated !== undefined) {
                cache.set(itemId, item.Qty * parentAccumulated);
              } else {
                cache.set(itemId, item.Qty);
              }
            } else {
              cache.set(itemId, item.Qty);
            }
          }
        }
        return cache;
      };

      // BUAT SET UNTUK MENAMPUNG SPK YANG SEDANG DIPREVIEW
      const exportingSPKsSet = new Set(
        selectedOrders.map((order) => order.order.No_SPK),
      );

      // RESERVATIONS BY ITEM (dengan filter self-reservation)
      const reservationsByItem = new Map<
        string,
        {
          totalQty: number;
          spkList: Set<{
            namaPO: string;
            qtyReserved: number;
          }>;
          itemName: string;
        }
      >();

      for (const reservation of stockReservations) {
        if (
          reservation.status !== "RESERVED" ||
          reservation.reservedQty <= 0 ||
          !reservation.noSPK
        )
          continue;

        if (exportingSPKsSet.has(reservation.noSPK)) {
          console.log(
            `Skipping self-reservation for PO: ${reservation.noSPK}, Qty: ${reservation.reservedQty}`,
          );
          continue;
        }

        const itemId = normalizeItemId(reservation.itemID);
        if (!reservationsByItem.has(itemId)) {
          reservationsByItem.set(itemId, {
            totalQty: 0,
            spkList: new Set(),
            itemName: reservation.itemName || itemId,
          });
        }

        const itemData = reservationsByItem.get(itemId)!;
        itemData.totalQty += reservation.reservedQty;

        const namaPO = reservation.namaPO || reservation.noSPK;

        let existing: { namaPO: string; qtyReserved: number } | undefined;
        for (const item of itemData.spkList) {
          if (item.namaPO === namaPO) {
            existing = item;
            break;
          }
        }

        if (existing) {
          existing.qtyReserved += reservation.reservedQty;
        } else {
          itemData.spkList.add({
            namaPO: namaPO,
            qtyReserved: reservation.reservedQty,
          });
        }
      }

      // MATERIAL AGGREGATION MAP
      const materialAggMap = new Map<string, any>();

      for (const order of ordersWithBom) {
        if (!order.bom || !order.stock) continue;

        const isCombined =
          order.order.combinedItems && order.order.combinedItems.length > 1;
        const barangJadiItems: Array<{
          kode: string;
          qty: number;
          nama: string;
        }> = [];

        if (isCombined && order.order.combinedItems) {
          order.order.combinedItems.forEach((item) =>
            barangJadiItems.push({
              kode: item.Kode_Barang,
              qty: item.QTY,
              nama: item.Nama_PO,
            }),
          );
        } else {
          barangJadiItems.push({
            kode: order.order.Kode_Barang,
            qty: order.order.QTY,
            nama: order.order.Nama_PO,
          });
        }

        for (const barangJadi of barangJadiItems) {
          let bomFlat: BomItem[] = [];
          if (isCombined && order.bom?.combinedBoms) {
            const bomItem = order.bom.combinedBoms[barangJadi.kode];
            if (bomItem) bomFlat = bomItem.flat;
          } else {
            bomFlat = order.bom.flat;
          }

          if (bomFlat.length === 0) continue;

          const components = bomFlat.filter(
            (b) => Number(b.Level) > 0 && !isINJECTIONDepartment(b.Departemen),
          );

          if (components.length === 0) continue;

          const accumulatedMap = calculateAccumulatedQtyForMaterial(bomFlat);
          const tempNeeds = new Map<string, number>();

          for (const component of components) {
            const materialId = normalizeItemId(component.ItemID);
            const accumulatedQty =
              accumulatedMap.get(materialId) || component.Qty;
            const needed = accumulatedQty * barangJadi.qty;
            tempNeeds.set(
              materialId,
              (tempNeeds.get(materialId) || 0) + needed,
            );
          }

          for (const [materialId, needed] of tempNeeds) {
            const masterInfo = masterDataMap.get(materialId) || {
              spec: "-",
              warna: "-",
              bahan: "-",
            };

            const stockItem = order.stock.find(
              (s) => normalizeItemId(s.itemid) === materialId,
            );

            const stockWincp = stockItem?.physicalStock || 0;
            const stockAkhir = stockItem?.stockAkhir || 0;

            const reservedData = reservationsByItem.get(materialId);
            const qtyReservedFromOtherPO = reservedData?.totalQty || 0;

            const component = components.find(
              (c) => normalizeItemId(c.ItemID) === materialId,
            );

            if (!materialAggMap.has(materialId)) {
              materialAggMap.set(materialId, {
                kode: materialId,
                nama: component?.ItemName || materialId,
                nama_china: component?.ItemName2 || "-",
                spec: masterInfo.spec,
                warna: masterInfo.warna,
                bahan: masterInfo.bahan,
                departemen: component?.Departemen || "UNKNOWN",
                totalNeeded: 0,
                stockWincp: stockWincp,
                stockAkhir: stockAkhir,
                qtyReserved: qtyReservedFromOtherPO,
                barangJadiSet: new Map(),
              });
            }

            const agg = materialAggMap.get(materialId);
            agg.totalNeeded += needed;

            if (!agg.barangJadiSet.has(barangJadi.kode)) {
              agg.barangJadiSet.set(barangJadi.kode, {
                qty: barangJadi.qty,
                nama: barangJadi.nama,
                kode: barangJadi.kode,
              });
            }
          }
        }
      }

      // ==================== BUAT MATERIAL DATA UNTUK SAVE KE DATABASE ====================
      for (const agg of materialAggMap.values()) {
        const barangJadiDetails: string[] = [];
        for (const [kode, info] of agg.barangJadiSet) {
          barangJadiDetails.push(`${kode} (${info.qty.toLocaleString()})`);
        }

        const reservationDetailsList: string[] = [];

        const reservedData = reservationsByItem.get(agg.kode);

        if (reservedData && reservedData.spkList.size > 0) {
          for (const spkReservation of reservedData.spkList) {
            let isExportingSPK = false;

            for (const order of selectedOrders) {
              if (
                order.order.No_SPK === spkReservation.namaPO ||
                order.order.Nama_PO === spkReservation.namaPO
              ) {
                isExportingSPK = true;
                break;
              }
            }

            if (!isExportingSPK) {
              reservationDetailsList.push(
                `${spkReservation.namaPO} (${spkReservation.qtyReserved.toLocaleString()})`,
              );
            }
          }
        }

        const reservedByText =
          reservationDetailsList.length > 0
            ? reservationDetailsList.join("\n")
            : "-";

        const variantInfo = getVariantInfo(agg.kode);

        const totalDibutuhkan = agg.totalNeeded + agg.qtyReserved;
        const available = agg.stockWincp - totalDibutuhkan;
        const kekurangan = totalDibutuhkan - agg.stockWincp;

        let status = "";
        if (agg.stockWincp >= totalDibutuhkan) {
          status = "AMAN";
        } else if (agg.stockWincp > 0) {
          status = "KURANG";
        } else {
          status = "HABIS";
        }

        materialDataForDatabase.push({
          "Kode Material": agg.kode,
          "Nama Material": agg.nama,
          "Nama China": agg.nama_china,
          Spesifikasi: agg.spec,
          Warna: agg.warna,
          Bahan: agg.bahan,
          Departemen: agg.departemen,
          "Barang Jadi": barangJadiDetails.join("\n"),
          "Total Kebutuhan": agg.totalNeeded.toLocaleString(),
          "Stok Wincp (Real)": agg.stockWincp.toLocaleString(),
          "Saldo Akhir": agg.stockAkhir.toLocaleString(),
          "Qty Reserved (PO Lain)": agg.qtyReserved.toLocaleString(),
          "Total Dibutuhkan": totalDibutuhkan.toLocaleString(),
          "Qty Available": available.toLocaleString(),
          "Reserved Oleh SPK": reservedByText,
          "Keterangan Variant": variantInfo,
          "Status Stock": status,
          Kekurangan: kekurangan > 0 ? kekurangan.toLocaleString() : "0",
        });
      }

      materialDataForDatabase.sort((a, b) =>
        a["Kode Material"].localeCompare(b["Kode Material"]),
      );

      // ==================== AUTO SAVE KE DATABASE ====================
      setExportProgress({
        visible: true,
        current: 70,
        total: 100,
        message: "Menyimpan ke database history...",
      });

      try {
        // Hitung summary untuk database
        const totalKebutuhan = materialDataForDatabase.reduce(
          (sum, item) =>
            sum +
            (parseInt(String(item["Total Kebutuhan"]).replace(/,/g, "")) || 0),
          0,
        );

        const totalKekurangan = materialDataForDatabase.reduce((sum, item) => {
          const kekurangan =
            parseInt(String(item["Kekurangan"]).replace(/,/g, "")) || 0;
          return sum + kekurangan;
        }, 0);

        const materialAman = materialDataForDatabase.filter(
          (item) => item["Status Stock"] === "AMAN",
        ).length;
        const materialKurang = materialDataForDatabase.filter(
          (item) => item["Status Stock"] === "KURANG",
        ).length;
        const materialHabis = materialDataForDatabase.filter(
          (item) => item["Status Stock"] === "HABIS",
        ).length;

        // Buat nama perhitungan
        let calculationName = "";
        if (selectedOrders.length === 1) {
          const singlePO = selectedOrders[0];
          calculationName = singlePO.order.Nama_PO || singlePO.order.No_SPK;
          calculationName = calculationName
            .replace(/[\\/*?:"<>|]/g, "")
            .replace(/\s+/g, "_")
            .substring(0, 50);
        } else {
          const firstPO = selectedOrders[0];
          const firstPOName = firstPO.order.Nama_PO || firstPO.order.No_SPK;
          const cleanFirstPO = firstPOName
            .replace(/[\\/*?:"<>|]/g, "")
            .replace(/\s+/g, "_")
            .substring(0, 40);
          calculationName = `${cleanFirstPO}_dan_${selectedOrders.length - 1}_lainnya`;
        }

        const timestamp = new Date()
          .toISOString()
          .replace(/[-:]/g, "")
          .split(".")[0];
        const calculationId = `${calculationName}_${timestamp}`;

        // Data PO yang dipilih
        const poList = selectedOrders.map((order) => ({
          no_spk: order.order.No_SPK,
          nama_po: order.order.Nama_PO,
          kode_barang: order.order.Kode_Barang,
          qty: order.order.QTY,
          tanggal_order: order.order.Tanggal_Order,
          is_combined:
            order.order.combinedItems && order.order.combinedItems.length > 1,
        }));

        const saveData = {
          calculation_id: calculationId,
          calculation_name: calculationName,
          user_id: "current_user",
          po_list: poList,
          total_po: selectedOrders.length,
          material_data: materialDataForDatabase,
          total_materials: materialDataForDatabase.length,
          total_kebutuhan: totalKebutuhan,
          total_kekurangan: totalKekurangan,
          material_aman: materialAman,
          material_kurang: materialKurang,
          material_habis: materialHabis,
          stock_date: today,
          notes: `Auto-saved from export - ${calculationName}`,
        };

        const saveResponse = await fetch(
          "/api/ppic/save-material-requirements",
          {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(saveData),
          },
        );

        const saveResult = await saveResponse.json();

        if (saveResult.success) {
          console.log("✅ Auto-saved to database:", calculationId);
          showToast(`✅ Data otomatis tersimpan ke history!`, "success");
          await loadSavedCalculations(); // Refresh history panel
        } else {
          console.warn("⚠️ Failed to auto-save:", saveResult.error);
          showToast(`⚠️ Gagal auto-save ke database`, "error");
        }
      } catch (saveError) {
        console.error("Error auto-saving to database:", saveError);
        // Jangan hentikan export jika save gagal
      }

      setExportProgress({
        visible: true,
        current: 80,
        total: 100,
        message: "Membuat file Excel...",
      });

      const wb = XLSX.utils.book_new();

      // ==================== SHEET 1: PO ====================
      const poData: any[] = [];
      ordersWithBom.forEach((order) => {
        const isCombined =
          order.order.combinedItems && order.order.combinedItems.length > 1;
        if (isCombined && order.order.combinedItems) {
          order.order.combinedItems.forEach((item) => {
            poData.push({
              "No SPK": order.order.No_SPK,
              "Tanggal Order": order.order.Tanggal_Order,
              "Tanggal Stok": today,
              "Nama PO": item.Nama_PO,
              "Kode Barang Jadi": item.Kode_Barang,
              "QTY PO": item.QTY,
            });
          });
        } else {
          poData.push({
            "No SPK": order.order.No_SPK,
            "Tanggal Order": order.order.Tanggal_Order,
            "Tanggal Stok": today,
            "Nama PO": order.order.Nama_PO,
            "Kode Barang Jadi": order.order.Kode_Barang,
            "QTY PO": order.order.QTY,
          });
        }
      });

      const wsPO = XLSX.utils.json_to_sheet(poData);
      wsPO["!cols"] = [
        { wch: 15 },
        { wch: 12 },
        { wch: 12 },
        { wch: 40 },
        { wch: 15 },
        { wch: 12 },
      ];
      XLSX.utils.book_append_sheet(wb, wsPO, "PO");

      // ==================== SHEET 2: BOM (FINAL - DENGAN PARENT TRACKING & DUPLICATE HANDLING) ====================
      const bomData: any[] = [];
      let totalINJECTIONRemoved = 0;

      // Build tree structure dengan mempertahankan semua node (termasuk duplicate kode di level berbeda)
      const buildTreeWithDuplicates = (flatBom: BomItem[]): BomItem[] => {
        if (!flatBom || flatBom.length === 0) return [];

        // Gunakan index sebagai key untuk membedakan node yang sama kodenya
        const nodeMap = new Map<string, BomItem>();
        const rootItems: BomItem[] = [];

        // Buat node dengan ID unik (gunakan kombinasi ItemID + Level + index)
        flatBom.forEach((item, idx) => {
          const uniqueKey = `${normalizeItemId(item.ItemID)}_L${item.Level}_${idx}`;
          const node: BomItem = {
            ...item,
            children: [],
          };
          nodeMap.set(uniqueKey, node);
        });

        // Bangun parent-child relationship
        flatBom.forEach((item, idx) => {
          const uniqueKey = `${normalizeItemId(item.ItemID)}_L${item.Level}_${idx}`;
          const node = nodeMap.get(uniqueKey);
          if (!node) return;

          const level = Number(item.Level);

          if (level === 1) {
            rootItems.push(node);
          } else {
            // Cari parent - bisa berdasarkan ParentItemID atau level
            let parentFound = false;

            // Cari parent berdasarkan ParentItemID
            if (item.ParentItemID) {
              const parentNormalizedId = normalizeItemId(item.ParentItemID);
              const parentKey = Array.from(nodeMap.keys()).find(
                (key) =>
                  key.startsWith(parentNormalizedId) &&
                  key.includes(`_L${level - 1}_`),
              );
              if (parentKey) {
                const parent = nodeMap.get(parentKey);
                if (parent) {
                  if (!parent.children) parent.children = [];
                  parent.children.push(node);
                  parentFound = true;
                }
              }
            }

            // Jika tidak ditemukan, cari parent berdasarkan level
            if (!parentFound) {
              const parentItem = flatBom.find(
                (p) => Number(p.Level) === level - 1,
              );
              if (parentItem) {
                const parentKey = `${normalizeItemId(parentItem.ItemID)}_L${parentItem.Level}_${flatBom.indexOf(parentItem)}`;
                const parent = nodeMap.get(parentKey);
                if (parent) {
                  if (!parent.children) parent.children = [];
                  parent.children.push(node);
                  parentFound = true;
                }
              }
            }

            if (!parentFound) {
              rootItems.push(node);
            }
          }
        });

        // Urutkan children
        const sortChildren = (nodes: BomItem[]) => {
          nodes.sort((a, b) => {
            if (Number(a.Level) !== Number(b.Level)) {
              return Number(a.Level) - Number(b.Level);
            }
            return (a.ItemID || "").localeCompare(b.ItemID || "");
          });
          nodes.forEach((node) => {
            if (node.children && node.children.length > 0) {
              sortChildren(node.children);
            }
          });
        };

        sortChildren(rootItems);
        return rootItems;
      };

      const formatIndentedName = (
        itemName: string,
        level: number,
        isDuplicate: boolean = false,
      ): string => {
        if (!itemName) return "";
        if (level === 1) return `📦 ${itemName}`;
        const indent = "  ".repeat(level - 1);
        const prefix = isDuplicate ? "↳ " : "└─ ";
        return `${indent}${prefix}${itemName}`;
      };

      // Fungsi untuk menghitung accumulated qty
      const calculateAccumulatedQtyForBOM = (
        flatBom: BomItem[],
      ): Map<string, number> => {
        const cache = new Map<string, number>();
        const itemMap = new Map<string, BomItem>();

        for (const item of flatBom) {
          const key = `${normalizeItemId(item.ItemID)}_L${item.Level}`;
          itemMap.set(key, item);
        }

        for (const item of flatBom) {
          const itemId = normalizeItemId(item.ItemID);
          const level = Number(item.Level);
          const key = `${itemId}_L${level}`;

          if (level === 1) {
            cache.set(key, item.Qty);
          } else {
            let parent: BomItem | undefined = undefined;

            if (item.ParentItemID) {
              const parentId = normalizeItemId(item.ParentItemID);
              const parentKey = `${parentId}_L${level - 1}`;
              parent = itemMap.get(parentKey);
            }

            if (!parent) {
              parent = flatBom.find((p) => Number(p.Level) === level - 1);
            }

            if (parent) {
              const parentId = normalizeItemId(parent.ItemID);
              const parentKey = `${parentId}_L${level - 1}`;
              const parentAccumulated = cache.get(parentKey);
              if (parentAccumulated !== undefined) {
                cache.set(key, item.Qty * parentAccumulated);
              } else {
                cache.set(key, item.Qty);
              }
            } else {
              cache.set(key, item.Qty);
            }
          }
        }
        return cache;
      };

      for (const order of ordersWithBom) {
        if (!order.bom) continue;
        const isCombined =
          order.order.combinedItems && order.order.combinedItems.length > 1;

        const processBom = (bomFlat: BomItem[], poQty: number, poItem: any) => {
          // Filter komponen (Level > 0) dan exclude INJECTION department
          const filteredBom = bomFlat.filter(
            (b) => Number(b.Level) > 0 && !isINJECTIONDepartment(b.Departemen),
          );
          const removedCount = bomFlat.filter(
            (b) => Number(b.Level) > 0 && isINJECTIONDepartment(b.Departemen),
          ).length;
          totalINJECTIONRemoved += removedCount;

          if (filteredBom.length === 0) return;

          // Hitung accumulated qty
          const accumulatedMap = calculateAccumulatedQtyForBOM(filteredBom);

          // HEADER untuk setiap PO
          bomData.push({
            "No SPK": order.order.No_SPK,
            "Kode Barang Jadi": poItem.Kode_Barang,
            "Nama Barang Jadi": poItem.Nama_PO,
            "QTY PO": poQty,
            Level: "HEADER",
            "Kode Komponen": "",
            "Nama Komponen": "",
            "Nama Komponen China": "",
            "Qty per Unit (BOM)": "",
            "Accumulated Qty": "",
            "Total Kebutuhan": "",
            Stok: "",
            Status: "",
            "Keterangan Perhitungan Accumulated": "",
          });

          // Build tree dengan mempertahankan semua node
          const treeStructure = buildTreeWithDuplicates(filteredBom);

          // Track item yang sudah ditampilkan untuk deteksi duplikat
          const displayedItems = new Map<string, number>();

          const traverseTree = (nodes: BomItem[]) => {
            for (const node of nodes) {
              const nodeLevel = Number(node.Level);
              const nodeId = normalizeItemId(node.ItemID);
              const stockItem = order.stock?.find(
                (s) => normalizeItemId(s.itemid) === nodeId,
              );

              const accumulatedKey = `${nodeId}_L${nodeLevel}`;
              const accumulatedQty =
                accumulatedMap.get(accumulatedKey) || node.Qty;
              const totalNeeded = accumulatedQty * poQty;
              const stock = stockItem?.stockAkhir || 0;
              const shortage = totalNeeded > stock;

              // Cek apakah ini duplikat (kode sama di level berbeda)
              const prevLevel = displayedItems.get(nodeId);
              const isDuplicate =
                prevLevel !== undefined && prevLevel !== nodeLevel;
              displayedItems.set(nodeId, nodeLevel);

              let calculationNote = "";
              if (nodeLevel === 1) {
                calculationNote = `Qty per Unit × QTY PO = ${node.Qty} × ${poQty} = ${totalNeeded.toLocaleString()}`;
              } else {
                const parentAccumulated = accumulatedQty / node.Qty;
                calculationNote = `Qty per Unit × Accumulated Parent × QTY PO = ${node.Qty} × ${parentAccumulated} × ${poQty} = ${totalNeeded.toLocaleString()}`;
              }

              bomData.push({
                "No SPK": "",
                "Kode Barang Jadi": "",
                "Nama Barang Jadi": "",
                "QTY PO": "",
                Level: node.Level,
                "Kode Komponen": node.ItemID,
                "Nama Komponen": formatIndentedName(
                  node.ItemName || node.ItemID,
                  nodeLevel,
                  isDuplicate,
                ),
                "Nama Komponen China": node.ItemName2 || "",
                "Qty per Unit (BOM)": node.Qty,
                "Accumulated Qty": accumulatedQty,
                "Total Kebutuhan": totalNeeded.toLocaleString(),
                Stok: stock.toLocaleString(),
                Status: shortage ? "KURANG" : "CUKUP",
                "Keterangan Perhitungan Accumulated": calculationNote,
              });

              if (node.children && node.children.length > 0) {
                traverseTree(node.children);
              }
            }
          };

          traverseTree(treeStructure);
          bomData.push({}); // Baris kosong sebagai separator antar BOM
        };

        if (isCombined && order.order.combinedItems) {
          for (const poItem of order.order.combinedItems) {
            let bomFlat: BomItem[] = [];
            if (
              order.bom?.combinedBoms &&
              order.bom.combinedBoms[poItem.Kode_Barang]
            ) {
              bomFlat = order.bom.combinedBoms[poItem.Kode_Barang].flat;
            } else {
              bomFlat = order.bom?.flat || [];
            }
            processBom(bomFlat, poItem.QTY, poItem);
          }
        } else {
          processBom(order.bom.flat, order.order.QTY, order.order);
        }
      }

      // Tambahkan informasi di akhir sheet BOM
      bomData.push({});
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": "📦 = Produk Level 1 (Barang Jadi)",
      });
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": "└─ = Sub-komponen Level 2",
      });
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": "  └─ = Sub-komponen Level 3",
      });
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": "    └─ = Sub-komponen Level 4",
      });
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": "↳ = Duplikat item (kode sama tapi level berbeda)",
      });
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": `* Komponen dengan departemen INJECTION tidak ditampilkan (${totalINJECTIONRemoved} item dihapus)`,
      });
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": `* Qty ditampilkan dalam format terformat (contoh: 1,000)`,
      });

      // Buat worksheet BOM
      const wsBOM = XLSX.utils.json_to_sheet(bomData);
      wsBOM["!cols"] = [
        { wch: 15 }, // No SPK
        { wch: 18 }, // Kode Barang Jadi
        { wch: 35 }, // Nama Barang Jadi
        { wch: 12 }, // QTY PO
        { wch: 10 }, // Level
        { wch: 18 }, // Kode Komponen
        { wch: 55 }, // Nama Komponen (lebih lebar untuk indentasi)
        { wch: 35 }, // Nama Komponen China
        { wch: 18 }, // Qty per Unit (BOM)
        { wch: 18 }, // Accumulated Qty
        { wch: 18 }, // Total Kebutuhan
        { wch: 15 }, // Stok
        { wch: 12 }, // Status
        { wch: 60 }, // Keterangan Perhitungan Accumulated
      ];

      // Styling untuk header (opsional - membuat baris header lebih menonjol)
      const range = XLSX.utils.decode_range(wsBOM["!ref"] || "A1:N1");
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_cell({ r: 0, c: C });
        if (!wsBOM[address]) continue;
        wsBOM[address].s = {
          font: { bold: true },
          fill: { fgColor: { rgb: "D3D3D3" } },
        };
      }

      XLSX.utils.book_append_sheet(wb, wsBOM, "BOM");

      // ==================== SHEET PER DEPARTEMEN ====================
      const materialsByDept = new Map<string, Map<string, any[]>>();

      for (const row of materialDataForDatabase) {
        const dept = row["Departemen"] || "UNKNOWN";
        const materialCode = row["Kode Material"];
        if (!materialsByDept.has(dept)) materialsByDept.set(dept, new Map());
        const deptMap = materialsByDept.get(dept)!;

        const rowArray = [
          row["Barang Jadi"],
          "", // QTY PO Dipesan (akan diisi nanti)
          row["Kode Material"],
          row["Nama Material"],
          row["Nama China"],
          row["Spesifikasi"],
          row["Warna"],
          row["Bahan"],
          row["Departemen"],
          parseInt(row["Total Kebutuhan"].replace(/,/g, "")) || 0,
          parseInt(row["Qty Reserved (PO Lain)"].replace(/,/g, "")) || 0,
          parseInt(row["Total Dibutuhkan"].replace(/,/g, "")) || 0,
          parseInt(row["Stok Wincp (Real)"].replace(/,/g, "")) || 0,
          parseInt(row["Saldo Akhir"].replace(/,/g, "")) || 0,
          parseInt(row["Qty Available"].replace(/,/g, "")) || 0,
          row["Reserved Oleh SPK"],
          row["Status Stock"],
          row["Keterangan Variant"],
        ];

        if (!deptMap.has(materialCode)) {
          deptMap.set(materialCode, rowArray);
        } else {
          const existing = deptMap.get(materialCode)!;
          existing[9] = (existing[9] || 0) + (rowArray[9] || 0);
          existing[11] = (existing[11] || 0) + (rowArray[11] || 0);
          existing[14] = (existing[13] || 0) - (existing[11] || 0);
          existing[16] =
            existing[14] > 0
              ? "KELEBIHAN"
              : existing[14] < 0
                ? "KURANG"
                : "CUKUP";

          const existingBarangJadi = existing[0] || "";
          const newBarangJadi = rowArray[0] || "";
          if (
            newBarangJadi &&
            !existingBarangJadi.includes(newBarangJadi.split("\n")[0])
          ) {
            existing[0] =
              existingBarangJadi +
              (existingBarangJadi ? "\n" : "") +
              newBarangJadi;
          }

          const existingReserved = existing[15] || "";
          const newReserved = rowArray[15] || "";
          if (
            newReserved !== "-" &&
            newReserved &&
            !existingReserved.includes(newReserved)
          ) {
            existing[15] =
              existingReserved +
              (existingReserved !== "-" && existingReserved ? "\n" : "") +
              newReserved;
          }

          const existingVariant = existing[17] || "";
          const newVariant = rowArray[17] || "";
          if (
            newVariant !== "-" &&
            newVariant !== existingVariant &&
            !existingVariant.includes(newVariant)
          ) {
            existing[17] =
              existingVariant +
              (existingVariant !== "-" && existingVariant ? " / " : "") +
              newVariant;
          }

          deptMap.set(materialCode, existing);
        }
      }

      const finalMaterialsByDept = new Map<string, any[][]>();
      for (const [dept, materialMap] of materialsByDept) {
        const rows: any[][] = [];
        for (const row of materialMap.values()) rows.push(row);
        rows.sort((a, b) => a[2].localeCompare(b[2]));
        finalMaterialsByDept.set(dept, rows);
      }

      const sortedDepartments = Array.from(finalMaterialsByDept.keys()).sort();
      const deptColWidths = [
        { wch: 50 },
        { wch: 20 },
        { wch: 15 },
        { wch: 40 },
        { wch: 35 },
        { wch: 30 },
        { wch: 20 },
        { wch: 25 },
        { wch: 20 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 50 },
        { wch: 15 },
        { wch: 25 },
      ];

      const headersDept = [
        "Barang Jadi",
        "QTY PO Dipesan",
        "Kode Material",
        "Nama Material",
        "Nama China",
        "Spesifikasi",
        "Warna",
        "Bahan",
        "Departemen",
        "Total Kebutuhan",
        "Reserved (Qty PO Lain)",
        "Total Dibutuhkan",
        "Stok Wincp (Real)",
        "Stok Akhir",
        "Sisa Stok",
        "Reserved Oleh SPK",
        "Status",
        "Keterangan Variant",
      ];

      for (const dept of sortedDepartments) {
        const deptMaterials = finalMaterialsByDept.get(dept) || [];
        const totalNeeded = deptMaterials.reduce(
          (sum, row) => sum + (row[9] || 0),
          0,
        );
        const totalSisa = deptMaterials.reduce(
          (sum, row) => sum + (row[14] || 0),
          0,
        );

        const variantItems = deptMaterials
          .filter((row) => row[17] && row[17] !== "-")
          .map((row) => `${row[2]} (${row[3]})`);
        const variantNote =
          variantItems.length > 0
            ? `Catatan: Item dengan variant (Grade A,B,C): ${variantItems.join(", ")}`
            : "";

        const wsData = [
          [`LAPORAN KEBUTUHAN MATERIAL - DEPARTEMEN ${dept.toUpperCase()}`],
          [
            `Tanggal Export: ${new Date().toLocaleDateString("id-ID")} ${new Date().toLocaleTimeString("id-ID")}`,
          ],
          [`Tanggal Stok: ${today}`],
          [
            `Catatan: Material dengan kode yang sama telah dijumlahkan total kebutuhannya`,
          ],
          variantNote ? [`${variantNote}`] : [],
          [],
          ["DETAIL MATERIAL"],
          headersDept,
          ...deptMaterials,
          [],
          [
            `Total Keseluruhan: ${deptMaterials.length} material unik, ` +
              `Total Kebutuhan: ${totalNeeded.toLocaleString()}, ` +
              `Total Sisa Stok: ${totalSisa.toLocaleString()}`,
          ],
        ];

        const finalWsData = variantNote
          ? wsData
          : wsData.filter((_, idx) => idx !== 4);

        const wsDept = XLSX.utils.aoa_to_sheet(finalWsData);
        wsDept["!cols"] = deptColWidths;

        wsDept["!merges"] = [
          { s: { r: 0, c: 0 }, e: { r: 0, c: headersDept.length - 1 } },
          { s: { r: 1, c: 0 }, e: { r: 1, c: headersDept.length - 1 } },
          { s: { r: 2, c: 0 }, e: { r: 2, c: headersDept.length - 1 } },
          { s: { r: 3, c: 0 }, e: { r: 3, c: headersDept.length - 1 } },
        ];

        if (variantNote) {
          wsDept["!merges"].push({
            s: { r: 4, c: 0 },
            e: { r: 4, c: headersDept.length - 1 },
          });
        }

        let sheetName = sanitizeSheetName(dept.toUpperCase());
        sheetName = getUniqueSheetName(wb, sheetName);

        XLSX.utils.book_append_sheet(wb, wsDept, sheetName);
      }

      // ==================== SHEET REKAP PER DEPARTEMEN ====================
      const allDeptSummary: any[][] = [
        ["REKAP KEBUTUHAN MATERIAL PER DEPARTEMEN"],
        [
          `Tanggal Export: ${new Date().toLocaleDateString("id-ID")} ${new Date().toLocaleTimeString("id-ID")}`,
        ],
        [`Tanggal Stok: ${today}`],
        [],
        [
          "Departemen",
          "Jumlah Material",
          "Total Kebutuhan",
          "Total Sisa Stok",
          "Status",
        ],
      ];

      for (const dept of sortedDepartments) {
        const deptMaterials = finalMaterialsByDept.get(dept) || [];
        const totalNeeded = deptMaterials.reduce(
          (sum, row) => sum + (row[9] || 0),
          0,
        );
        const totalSisa = deptMaterials.reduce(
          (sum, row) => sum + (row[14] || 0),
          0,
        );
        const status =
          totalSisa > 0 ? "KELEBIHAN" : totalSisa < 0 ? "KEKURANGAN" : "CUKUP";
        allDeptSummary.push([
          dept,
          deptMaterials.length,
          totalNeeded.toLocaleString(),
          totalSisa.toLocaleString(),
          status,
        ]);
      }

      const totalAllMaterials = materialDataForDatabase.length;
      const totalAllNeeded = materialDataForDatabase.reduce(
        (sum, row) =>
          sum + (parseInt(row["Total Kebutuhan"].replace(/,/g, "")) || 0),
        0,
      );
      const totalAllSisa = materialDataForDatabase.reduce(
        (sum, row) =>
          sum + (parseInt(row["Qty Available"].replace(/,/g, "")) || 0),
        0,
      );

      allDeptSummary.push(
        [],
        [
          "TOTAL KESELURUHAN",
          totalAllMaterials,
          totalAllNeeded.toLocaleString(),
          totalAllSisa.toLocaleString(),
          totalAllSisa > 0
            ? "KELEBIHAN"
            : totalAllSisa < 0
              ? "KEKURANGAN"
              : "CUKUP",
        ],
      );

      const wsSummary = XLSX.utils.aoa_to_sheet(allDeptSummary);
      wsSummary["!cols"] = [
        { wch: 25 },
        { wch: 18 },
        { wch: 20 },
        { wch: 20 },
        { wch: 20 },
      ];
      wsSummary["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 4 } },
      ];

      XLSX.utils.book_append_sheet(wb, wsSummary, "REKAP_PER_DEPARTEMEN");

      // ==================== SHEET KETERANGAN ====================
      const keteranganData: any[][] = [
        ["📋 PETUNJUK MEMBACA LAPORAN KEBUTUHAN MATERIAL"],
        [""],
        ["A. INFORMASI UMUM"],
        ["No", "Item", "Keterangan"],
        ["1", "Tanggal Export", "Tanggal saat laporan diekspor"],
        [
          "2",
          "Tanggal Stok",
          "Tanggal data stok yang digunakan (berdasarkan order date)",
        ],
        [""],
        ["B. PENJELASAN KOLOM - SHEET PO"],
        ["No", "Kolom", "Keterangan", "Contoh"],
        ["1", "No SPK", "Nomor Surat Perintah Kerja / Production Order", "AS-"],
        ["2", "Tanggal Order", "Tanggal order diterima", "01/05/2025"],
        [
          "3",
          "Tanggal Stok",
          "Tanggal acuan data stok yang digunakan",
          "01/05/2025",
        ],
        [
          "4",
          "Nama PO",
          "Nama lengkap Production Order",
          "Produksi T-12345 ABC",
        ],
        [
          "5",
          "Kode Barang Jadi",
          "Kode produk/jadi yang akan diproduksi",
          "FG-001",
        ],
        [
          "6",
          "QTY PO",
          "Quantity yang dipesan (jumlah yang harus diproduksi)",
          "1,000",
        ],
        [""],
        ["C. PENJELASAN KOLOM - SHEET BOM"],
        ["No", "Kolom", "Keterangan", "Contoh"],
        [
          "1",
          "No SPK",
          "Nomor SPK (hanya di baris header)",
          "SPK-001/PPIC/V/2025",
        ],
        [
          "2",
          "Kode Barang Jadi",
          "Kode produk yang diproduksi (hanya di baris header)",
          "FG-001",
        ],
        [
          "3",
          "Nama Barang Jadi",
          "Nama produk yang diproduksi (hanya di baris header)",
          "Produk ABC",
        ],
        [
          "4",
          "QTY PO",
          "Quantity yang diproduksi (hanya di baris header)",
          "1,000",
        ],
        [
          "5",
          "Level",
          "Tingkat komponen (1=produk jadi, 2=sub-komponen, 3=sub-sub komponen, dll)",
          "1, 2, 3, ...",
        ],
        [
          "6",
          "Kode Komponen",
          "Kode material/komponen yang dibutuhkan",
          "RAW-001",
        ],
        [
          "7",
          "Nama Komponen",
          "Nama material/komponen (dengan indentasi)",
          "📦 Material A / └─ Sub Material",
        ],
        [
          "8",
          "Nama Komponen China",
          "Nama komponen dalam bahasa Mandarin",
          "材料A",
        ],
        [
          "9",
          "Qty per Unit (BOM)",
          "Jumlah material yang dibutuhkan per unit produk",
          "2",
        ],
        [
          "10",
          "Accumulated Qty",
          "Jumlah material per unit produk setelah dikalikan parent",
          "4",
        ],
        [
          "11",
          "Total Kebutuhan",
          "Total material yang dibutuhkan (Accumulated Qty × QTY PO)",
          "4,000",
        ],
        [
          "12",
          "Stok",
          "Stok yang tersedia saat ini (SaldoAkhirFisik - Reserved untuk PO aktif)",
          "5,000",
        ],
        ["13", "Status", "Status ketersediaan stok", "CUKUP / KURANG"],
        [""],
        ["D. PENJELASAN HIERARKI NAMA KOMPONEN"],
        ["No", "Tanda", "Level", "Keterangan", "Contoh"],
        ["1", "📦", "Level 1", "Produk / Komponen utama", "📦 Material A"],
        ["2", "└─", "Level 2", "Sub-komponen level 2", "└─ Material B"],
        ["3", "  └─", "Level 3", "Sub-komponen level 3", "  └─ Material C"],
        [
          "4",
          "    └─",
          "Level 4",
          "Sub-komponen level 4 (dan seterusnya)",
          "    └─ Material D",
        ],
        [""],
        ["E. PENJELASAN KOLOM - SHEET PER DEPARTEMEN"],
        ["No", "Kolom", "Keterangan", "Contoh"],
        [
          "1",
          "Barang Jadi",
          "Kode-kode produk jadi yang menggunakan material ini",
          "FG-001\nFG-002",
        ],
        [
          "2",
          "QTY PO Dipesan",
          "Quantity PO untuk masing-masing produk jadi",
          "1,000\n500",
        ],
        ["3", "Kode Material", "Kode material/komponen", "RAW-001"],
        ["4", "Nama Material", "Nama material/komponen", "Material A"],
        ["5", "Nama China", "Nama material dalam bahasa Mandarin", "材料A"],
        ["6", "Spesifikasi", "Spesifikasi teknis material", "Diameter 10mm"],
        ["7", "Warna", "Warna material", "Merah"],
        ["8", "Bahan", "Jenis bahan material", "Plastik ABS"],
        ["9", "Departemen", "Departemen yang bertanggung jawab", "INJEKSI"],
        [
          "10",
          "Total Kebutuhan",
          "Total kebutuhan material untuk semua PO",
          "5,000",
        ],
        [
          "11",
          "Reserved (Qty PO Lain)",
          "Quantity material yang sudah di-reserved untuk PO lain",
          "1,000",
        ],
        ["12", "Total Dibutuhkan", "Total Kebutuhan + Reserved", "6,000"],
        [
          "13",
          "Stok Wincp",
          "STOK REAL DI GUDANG WINCP (SaldoAkhirFisik)",
          "8,000",
        ],
        ["14", "Stok Akhir", "STOK BERSIH (Stok Fisik - Reserved)", "7,000"],
        [
          "15",
          "Sisa Stok",
          "Stok Akhir - Total Dibutuhkan (+ = kelebihan, - = kekurangan)",
          "1,000",
        ],
        [
          "16",
          "Reserved Oleh SPK",
          "Daftar SPK/PO yang mereserve material ini",
          "SPK-001 (500)\nSPK-002 (500)",
        ],
        ["17", "Status", "Status ketersediaan", "CUKUP / KURANG / HABIS"],
        [
          "18",
          "Keterangan Variant",
          "Informasi variant material (Grade A, B, C)",
          "Grade A / Grade B",
        ],
        [""],
        ["F. PENJELASAN KOLOM - SHEET REKAP PER DEPARTEMEN"],
        ["No", "Kolom", "Keterangan", "Contoh"],
        ["1", "Departemen", "Nama departemen", "INJEKSI"],
        [
          "2",
          "Jumlah Material",
          "Jumlah material unik di departemen ini",
          "25",
        ],
        [
          "3",
          "Total Kebutuhan",
          "Total kebutuhan material di departemen ini",
          "50,000",
        ],
        [
          "4",
          "Total Sisa Stok",
          "Total sisa stok di departemen ini (+ = kelebihan, - = kekurangan)",
          "5,000",
        ],
        [
          "5",
          "Status",
          "Status keseluruhan departemen",
          "KELEBIHAN / KEKURANGAN / CUKUP",
        ],
        [""],
        ["G. PENJELASAN STATUS"],
        ["No", "Status", "Keterangan", "Kondisi"],
        [
          "1",
          "CUKUP",
          "Stok mencukupi untuk memenuhi kebutuhan",
          "Sisa Stok ≥ 0",
        ],
        [
          "2",
          "KURANG",
          "Stok tidak mencukupi, perlu pembelian / produksi",
          "Sisa Stok < 0",
        ],
        ["3", "HABIS", "Stok habis tepat", "Sisa Stok = 0"],
        [
          "4",
          "KELEBIHAN",
          "Stok berlebih (khusus sheet Rekap)",
          "Total Sisa Stok > 0",
        ],
        [
          "5",
          "KEKURANGAN",
          "Stok kurang (khusus sheet Rekap)",
          "Total Sisa Stok < 0",
        ],
        [""],
        ["H. RUMUS PERHITUNGAN"],
        ["No", "Rumus", "Keterangan", "Contoh Perhitungan"],
        [
          "1",
          "Accumulated Qty",
          "Qty per Unit × Parent Accumulated / Parent Qty",
          "2 × 4 / 2 = 4",
        ],
        [
          "2",
          "Total Kebutuhan",
          "Accumulated Qty × QTY PO",
          "4 × 1,000 = 4,000",
        ],
        [
          "3",
          "Total Dibutuhkan",
          "Total Kebutuhan + Reserved Qty",
          "4,000 + 1,000 = 5,000",
        ],
        [
          "4",
          "Sisa Stok",
          "Stok Akhir - Total Dibutuhkan",
          "7,000 - 5,000 = 2,000",
        ],
        [
          "5",
          "SaldoAkhirFisik",
          "Stok fisik tanpa pengurangan reserved",
          "8,000",
        ],
        [
          "6",
          "SaldoAkhir",
          "SaldoAkhirFisik - TotalReserved (untuk PO aktif)",
          "8,000 - 1,000 = 7,000",
        ],
        [""],
        ["I. CATATAN PENTING"],
        ["No", "Catatan", "Keterangan"],
        [
          "1",
          "Aggregasi Material",
          "Material dengan kode yang sama dijumlahkan total kebutuhannya",
        ],
        [
          "2",
          "Filter INJECTION",
          "Komponen dengan departemen INJECTION tidak ditampilkan dalam perhitungan",
        ],
        [
          "3",
          "Filter Completed PO",
          "Reserved stok hanya diperhitungkan dari PO yang statusnya BELUM COMPLETED (Completed = 0)",
        ],
        [
          "4",
          "Stok Akhir",
          "Stok Akhir yang ditampilkan adalah stok setelah dikurangi reserved untuk PO aktif",
        ],
        [
          "5",
          "Hierarki",
          "Tanda 📦 dan └─ pada kolom Nama Komponen menunjukkan tingkatan hierarki",
        ],
        [
          "6",
          "Tanggal Stok",
          "Data stok diambil per tanggal order (masing-masing PO beda tanggal stok)",
        ],
        [
          "7",
          "Variant",
          "Variant (Grade A,B,C) pada material ditampilkan di kolom Keterangan Variant",
        ],
        [
          "8",
          "Reserved",
          "Reserved adalah stok yang sudah direserve untuk PO lain tapi belum di-commit",
        ],
        [
          "9",
          "Commit",
          "Commit adalah stok yang sudah direserve dan dikunci untuk PO tertentu",
        ],
        [""],
        ["J. INFORMASI FILE"],
        ["No", "Informasi", "Nilai"],
        ["1", "Nama File", fileName],
        ["2", "Jumlah PO Diexport", selectedOrders.length],
        ["3", "Total Material Unik", materialDataForDatabase.length],
        ["4", "Tanggal Export", new Date().toLocaleDateString("id-ID")],
        ["5", "Waktu Export", new Date().toLocaleTimeString("id-ID")],
        ["6", "User", "PPIC Department"],
        ["7", "Aplikasi", "Production Planning System"],
      ];

      const wsKeterangan = XLSX.utils.aoa_to_sheet(keteranganData);
      wsKeterangan["!cols"] = [
        { wch: 8 },
        { wch: 30 },
        { wch: 50 },
        { wch: 35 },
      ];
      wsKeterangan["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
        { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } },
        { s: { r: 8, c: 0 }, e: { r: 8, c: 3 } },
        { s: { r: 16, c: 0 }, e: { r: 16, c: 3 } },
        { s: { r: 27, c: 0 }, e: { r: 27, c: 4 } },
        { s: { r: 32, c: 0 }, e: { r: 32, c: 3 } },
        { s: { r: 52, c: 0 }, e: { r: 52, c: 3 } },
        { s: { r: 59, c: 0 }, e: { r: 59, c: 3 } },
        { s: { r: 66, c: 0 }, e: { r: 66, c: 3 } },
        { s: { r: 74, c: 0 }, e: { r: 74, c: 2 } },
        { s: { r: 85, c: 0 }, e: { r: 85, c: 2 } },
      ];

      XLSX.utils.book_append_sheet(wb, wsKeterangan, "KETERANGAN");

      XLSX.writeFile(wb, fileName);

      setExportProgress({ visible: false, current: 0, total: 0, message: "" });
      showToast(
        `✅ Export berhasil! ${selectedOrders.length} PO diexport & tersimpan di history`,
        "success",
      );
      alert(
        `✅ Export berhasil!\nFile: ${fileName}\n\n📦 Total PO: ${selectedOrders.length}\n📅 Stok per tanggal: ${today}\n💾 Data juga tersimpan di history database!`,
      );
    } catch (error) {
      console.error("Error export:", error);
      showToast(
        `❌ Gagal export: ${error instanceof Error ? error.message : "Unknown error"}`,
        "error",
      );
      alert(
        "Gagal mengekspor data: " +
          (error instanceof Error ? error.message : "Unknown error"),
      );
    } finally {
      setExportLoading(false);
      setExportProgress({ visible: false, current: 0, total: 0, message: "" });
    }
  };

  // ==================== FUNGSI SELECT ====================
  const toggleSelection = (index: number) => {
    const plan = paginatedOrders[index];
    if (!plan) return;

    const originalIndex = orders.findIndex(
      (o) => o.order.No_SPK === plan.order.No_SPK,
    );
    if (originalIndex === -1) return;

    setOrders((prev) =>
      prev.map((item, i) =>
        i === originalIndex
          ? { ...item, selected: !item.selected } // Izinkan semua PO dipilih
          : item,
      ),
    );
  };

  const toggleSelectAll = () => {
    // Select all PO yang ada di halaman saat ini (paginatedOrders)
    const currentPageOrders = paginatedOrders;
    const allSelected = currentPageOrders.every(
      (order) => order.selected || order.committed,
    );

    setOrders((prev) =>
      prev.map((order) => {
        // Cek apakah order ini ada di halaman saat ini
        const isInCurrentPage = currentPageOrders.some(
          (p) => p.order.No_SPK === order.order.No_SPK,
        );
        if (isInCurrentPage && !order.committed) {
          return { ...order, selected: !allSelected };
        }
        return order;
      }),
    );
  };

  const toggleSelectAllGlobal = () => {
    // Dapatkan semua PO yang aktif (belum di-commit) dari filteredOrders (bukan orders)
    const activeOrders = filteredOrders.filter((order) => !order.committed);
    const allSelected =
      activeOrders.length > 0 && activeOrders.every((order) => order.selected);

    setOrders((prev) =>
      prev.map((order) => {
        // Hanya toggle untuk PO yang aktif (belum di-commit)
        if (!order.committed) {
          return { ...order, selected: !allSelected };
        }
        return order;
      }),
    );
  };

  const handleDateFilterChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setDateFilter((prev) => ({ ...prev, [name]: value }));
  };

  const applyDateFilter = () => {
    fetchOrders(
      dateFilter.startDate,
      dateFilter.endDate,
      undefined,
      showCommitted,
    );
  };

  const resetDateFilter = () => {
    setDateFilter({ startDate: "", endDate: "" });
    fetchOrders(undefined, undefined, undefined, showCommitted);
  };

  const handleSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setSearchQuery(e.target.value);
    setCurrentPage(1);
  };

  const clearSearch = () => {
    setSearchQuery("");
    setCurrentPage(1);
  };

  // Fungsi untuk menyimpan material requirements ke database
  // Fungsi untuk menyimpan material requirements ke database
  const saveMaterialRequirementsToDatabase = async (materialData: any[]) => {
    try {
      const selectedOrders = filteredOrders.filter((order) => order.selected);

      // 🔥 BUAT NAMA BERDASARKAN PO (bukan CALC)
      let calculationName = "";
      if (selectedOrders.length === 1) {
        const singlePO = selectedOrders[0];
        calculationName = singlePO.order.Nama_PO || singlePO.order.No_SPK;
        // Bersihkan nama
        calculationName = calculationName
          .replace(/[\\/*?:"<>|]/g, "")
          .replace(/\s+/g, "_")
          .substring(0, 50);
      } else {
        const firstPO = selectedOrders[0];
        const firstPOName = firstPO.order.Nama_PO || firstPO.order.No_SPK;
        const cleanFirstPO = firstPOName
          .replace(/[\\/*?:"<>|]/g, "")
          .replace(/\s+/g, "_")
          .substring(0, 40);
        calculationName = `${cleanFirstPO}_dan_${selectedOrders.length - 1}_lainnya`;
      }

      // Tambahkan timestamp untuk uniqueness
      const timestamp = new Date()
        .toISOString()
        .replace(/[-:]/g, "")
        .split(".")[0];
      const calculationId = `${calculationName}_${timestamp}`;

      // Hitung summary
      const totalKebutuhan = materialData.reduce(
        (sum, item) =>
          sum +
          (parseInt(String(item["Total Kebutuhan"]).replace(/,/g, "")) || 0),
        0,
      );

      const totalKekurangan = materialData.reduce((sum, item) => {
        const kekurangan =
          parseInt(String(item["Kekurangan"]).replace(/,/g, "")) || 0;
        return sum + kekurangan;
      }, 0);

      const materialAman = materialData.filter(
        (item) => item["Status Stock"] === "AMAN",
      ).length;
      const materialKurang = materialData.filter(
        (item) => item["Status Stock"] === "KURANG",
      ).length;
      const materialHabis = materialData.filter(
        (item) => item["Status Stock"] === "HABIS",
      ).length;

      // Data PO yang dipilih
      const poList = selectedOrders.map((order) => ({
        no_spk: order.order.No_SPK,
        nama_po: order.order.Nama_PO,
        kode_barang: order.order.Kode_Barang,
        qty: order.order.QTY,
        tanggal_order: order.order.Tanggal_Order,
        is_combined:
          order.order.combinedItems && order.order.combinedItems.length > 1,
      }));

      // Siapkan data untuk disimpan
      const saveData = {
        calculation_id: calculationId,
        calculation_name: calculationName, // 🔥 TAMBAHKAN NAMA PERHITUNGAN
        user_id: "current_user",
        po_list: poList,
        total_po: selectedOrders.length,
        material_data: materialData,
        total_materials: materialData.length,
        total_kebutuhan: totalKebutuhan,
        total_kekurangan: totalKekurangan,
        material_aman: materialAman,
        material_kurang: materialKurang,
        material_habis: materialHabis,
        stock_date: new Date().toISOString().split("T")[0],
        notes: `Calculate ${selectedOrders.length} PO(s) - ${calculationName}`,
      };

      console.log("Saving to database:", saveData);

      const response = await fetch("/api/ppic/save-material-requirements", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(saveData),
      });

      const result = await response.json();

      if (result.success) {
        showToast(
          `✅ Data berhasil disimpan! Nama: ${calculationName}`,
          "success",
        );
        await loadSavedCalculations();
        return result;
      } else {
        throw new Error(result.error || "Failed to save");
      }
    } catch (error) {
      console.error("Error saving to database:", error);
      showToast(`⚠️ Gagal menyimpan ke database: ${error}`, "error");
      return null;
    }
  };

  // Fungsi untuk memuat history perhitungan yang tersimpan
  const loadSavedCalculations = async () => {
    try {
      const response = await fetch(
        "/api/ppic/save-material-requirements?limit=20",
      );
      const result = await response.json();
      if (result.success) {
        setSavedCalculations(result.data);
      }
    } catch (error) {
      console.error("Error loading saved calculations:", error);
    }
  };

  // Fungsi handle save dari preview dialog
  const handleSaveToDatabase = async () => {
    if (!previewDialog.data || previewDialog.data.length === 0) {
      showToast("Tidak ada data untuk disimpan", "error");
      return;
    }

    setSavingToDatabase(true);
    try {
      await saveMaterialRequirementsToDatabase(previewDialog.data);
    } catch (error) {
      console.error("Error saving:", error);
    } finally {
      setSavingToDatabase(false);
    }
  };

  const previewExport = async () => {
    try {
      setExportLoading(true);
      setExportProgress({
        visible: true,
        current: 0,
        total: 100,
        message: "Mempersiapkan preview data...",
      });
      showToast("Mempersiapkan preview data dengan cek stok...", "loading");

      const selectedOrders = filteredOrders.filter((order) => order.selected);

      if (selectedOrders.length === 0) {
        alert("Tidak ada PO yang dipilih untuk di-preview!");
        setExportLoading(false);
        setExportProgress({
          visible: false,
          current: 0,
          total: 0,
          message: "",
        });
        return;
      }

      setExportProgress({
        visible: true,
        current: 20,
        total: 100,
        message: `Memuat BOM & stok untuk ${selectedOrders.length} PO...`,
      });
      const ordersWithBom = await loadBomAndStockForExport(selectedOrders);

      setExportProgress({
        visible: true,
        current: 50,
        total: 100,
        message: "Mengambil data master...",
      });

      const allMaterialIds: string[] = [];
      for (const order of ordersWithBom) {
        if (order.bom?.flat) {
          order.bom.flat.forEach((item: BomItem) => {
            if (item.ItemID && Number(item.Level) > 0)
              allMaterialIds.push(normalizeItemId(item.ItemID));
          });
        }
      }
      const masterDataMap = await fetchMasterDataForItems(allMaterialIds);

      setExportProgress({
        visible: true,
        current: 70,
        total: 100,
        message: "Menghitung kebutuhan & reserved stock...",
      });

      // BUAT SET UNTUK MENAMPUNG SPK YANG SEDANG DIPREVIEW
      const exportingSPKs = new Set(
        selectedOrders.map((order) => order.order.No_SPK),
      );
      console.log(
        "Preview SPKs (excluding self-reservation):",
        Array.from(exportingSPKs),
      );

      // RESERVATIONS BY ITEM (dengan filter self-reservation)
      const reservationsByItem = new Map<
        string,
        {
          totalQty: number;
          spkList: Set<{
            namaPO: string;
            qtyReserved: number;
          }>;
          itemName: string;
        }
      >();

      for (const reservation of stockReservations) {
        if (
          reservation.status !== "RESERVED" ||
          reservation.reservedQty <= 0 ||
          !reservation.noSPK
        )
          continue;

        // LEWATKAN RESERVATION UNTUK PO YANG SEDANG DIPREVIEW
        if (exportingSPKs.has(reservation.noSPK)) {
          console.log(
            `Skipping self-reservation for PO: ${reservation.noSPK}, Qty: ${reservation.reservedQty}`,
          );
          continue;
        }

        const itemId = normalizeItemId(reservation.itemID);
        if (!reservationsByItem.has(itemId)) {
          reservationsByItem.set(itemId, {
            totalQty: 0,
            spkList: new Set(),
            itemName: reservation.itemName || itemId,
          });
        }

        const itemData = reservationsByItem.get(itemId)!;
        itemData.totalQty += reservation.reservedQty;

        const namaPO = reservation.namaPO || reservation.noSPK;

        let existing: { namaPO: string; qtyReserved: number } | undefined;
        for (const item of itemData.spkList) {
          if (item.namaPO === namaPO) {
            existing = item;
            break;
          }
        }

        if (existing) {
          existing.qtyReserved += reservation.reservedQty;
        } else {
          itemData.spkList.add({
            namaPO: namaPO,
            qtyReserved: reservation.reservedQty,
          });
        }
      }

      // FUNGSI CALCULATE ACCUMULATED QTY
      const calculateAccumulatedQtyForMaterial = (
        flatBom: BomItem[],
      ): Map<string, number> => {
        const cache = new Map<string, number>();
        const itemMap = new Map<string, BomItem>();

        for (const item of flatBom) {
          itemMap.set(normalizeItemId(item.ItemID), item);
        }

        for (const item of flatBom) {
          const itemId = normalizeItemId(item.ItemID);
          const level = Number(item.Level);

          if (level === 1) {
            cache.set(itemId, item.Qty);
          } else {
            let parent: BomItem | undefined = undefined;

            if (item.ParentItemID) {
              parent = itemMap.get(normalizeItemId(item.ParentItemID));
            }

            if (!parent) {
              parent = flatBom.find((p) => Number(p.Level) === level - 1);
            }

            if (parent) {
              const parentId = normalizeItemId(parent.ItemID);
              const parentAccumulated = cache.get(parentId);
              if (parentAccumulated !== undefined) {
                cache.set(itemId, item.Qty * parentAccumulated);
              } else {
                cache.set(itemId, item.Qty);
              }
            } else {
              cache.set(itemId, item.Qty);
            }
          }
        }
        return cache;
      };

      // MATERIAL AGGREGATION MAP
      const materialAggMap = new Map<string, any>();

      for (const order of ordersWithBom) {
        if (!order.bom || !order.stock) continue;

        const isCombined =
          order.order.combinedItems && order.order.combinedItems.length > 1;
        const barangJadiItems: Array<{
          kode: string;
          qty: number;
          nama: string;
        }> = [];

        if (isCombined && order.order.combinedItems) {
          order.order.combinedItems.forEach((item) =>
            barangJadiItems.push({
              kode: item.Kode_Barang,
              qty: item.QTY,
              nama: item.Nama_PO,
            }),
          );
        } else {
          barangJadiItems.push({
            kode: order.order.Kode_Barang,
            qty: order.order.QTY,
            nama: order.order.Nama_PO,
          });
        }

        for (const barangJadi of barangJadiItems) {
          let bomFlat: BomItem[] = [];
          if (isCombined && order.bom?.combinedBoms) {
            const bomItem = order.bom.combinedBoms[barangJadi.kode];
            if (bomItem) bomFlat = bomItem.flat;
          } else {
            bomFlat = order.bom.flat;
          }

          if (bomFlat.length === 0) continue;

          const components = bomFlat.filter(
            (b) => Number(b.Level) > 0 && !isINJECTIONDepartment(b.Departemen),
          );

          if (components.length === 0) continue;

          const accumulatedMap = calculateAccumulatedQtyForMaterial(bomFlat);
          const tempNeeds = new Map<string, number>();

          for (const component of components) {
            const materialId = normalizeItemId(component.ItemID);
            const accumulatedQty =
              accumulatedMap.get(materialId) || component.Qty;
            const needed = accumulatedQty * barangJadi.qty;
            tempNeeds.set(
              materialId,
              (tempNeeds.get(materialId) || 0) + needed,
            );
          }

          for (const [materialId, needed] of tempNeeds) {
            const masterInfo = masterDataMap.get(materialId) || {
              spec: "-",
              warna: "-",
              bahan: "-",
            };

            const stockItem = order.stock.find(
              (s) => normalizeItemId(s.itemid) === materialId,
            );

            const stockWincp = stockItem?.physicalStock || 0;
            const stockAkhir = stockItem?.stockAkhir || 0;

            const reservedData = reservationsByItem.get(materialId);
            const qtyReservedFromOtherPO = reservedData?.totalQty || 0;

            const component = components.find(
              (c) => normalizeItemId(c.ItemID) === materialId,
            );

            if (!materialAggMap.has(materialId)) {
              materialAggMap.set(materialId, {
                kode: materialId,
                nama: component?.ItemName || materialId,
                nama_china: component?.ItemName2 || "-",
                spec: masterInfo.spec,
                warna: masterInfo.warna,
                bahan: masterInfo.bahan,
                departemen: component?.Departemen || "UNKNOWN",
                totalNeeded: 0,
                stockWincp: stockWincp,
                stockAkhir: stockAkhir,
                qtyReserved: qtyReservedFromOtherPO,
                barangJadiSet: new Map(),
              });
            }

            const agg = materialAggMap.get(materialId);
            agg.totalNeeded += needed;

            if (!agg.barangJadiSet.has(barangJadi.kode)) {
              agg.barangJadiSet.set(barangJadi.kode, {
                qty: barangJadi.qty,
                nama: barangJadi.nama,
                kode: barangJadi.kode,
              });
            }
          }
        }
      }

      // ==================== HITUNG DETAIL PENGAMBILAN PER SPK ====================
      const spkReservationDetails = new Map<
        string,
        Map<
          string,
          {
            needed: number;
            taken: number;
            shortage: number;
            stockBefore: number;
            stockAfter: number;
          }
        >
      >();

      for (const order of ordersWithBom) {
        if (!order.bom || !order.stock) continue;

        const spk = order.order.No_SPK;
        if (!spkReservationDetails.has(spk)) {
          spkReservationDetails.set(spk, new Map());
        }

        const spkDetails = spkReservationDetails.get(spk)!;

        const isCombined =
          order.order.combinedItems && order.order.combinedItems.length > 1;
        const barangJadiItems: Array<{
          kode: string;
          qty: number;
          nama: string;
        }> = [];

        if (isCombined && order.order.combinedItems) {
          order.order.combinedItems.forEach((item) =>
            barangJadiItems.push({
              kode: item.Kode_Barang,
              qty: item.QTY,
              nama: item.Nama_PO,
            }),
          );
        } else {
          barangJadiItems.push({
            kode: order.order.Kode_Barang,
            qty: order.order.QTY,
            nama: order.order.Nama_PO,
          });
        }

        for (const barangJadi of barangJadiItems) {
          let bomFlat: BomItem[] = [];
          if (isCombined && order.bom?.combinedBoms) {
            const bomItem = order.bom.combinedBoms[barangJadi.kode];
            if (bomItem) bomFlat = bomItem.flat;
          } else {
            bomFlat = order.bom.flat;
          }

          if (bomFlat.length === 0) continue;

          const components = bomFlat.filter(
            (b) => Number(b.Level) > 0 && !isINJECTIONDepartment(b.Departemen),
          );

          if (components.length === 0) continue;

          const accumulatedMap = calculateAccumulatedQtyForMaterial(bomFlat);

          for (const component of components) {
            const materialId = normalizeItemId(component.ItemID);
            const accumulatedQty =
              accumulatedMap.get(materialId) || component.Qty;
            const needed = accumulatedQty * barangJadi.qty;

            // 🔥 PERBAIKAN: PAKAI physicalStock (Stok Wincp / Stok Real)
            const stockItem = order.stock.find(
              (s) => normalizeItemId(s.itemid) === materialId,
            );
            const stockReal = stockItem?.physicalStock || 0; // Stok Wincp (stok fisik real)

            // Reserved dari PO LAIN (bukan PO ini)
            const reservedFromOthers =
              reservationsByItem.get(materialId)?.totalQty || 0;

            // Stok yang benar-benar tersedia = Stok Real - Reserved PO Lain
            const availableStock = Math.max(0, stockReal - reservedFromOthers);

            // Berapa yang bisa diambil (tidak melebihi kebutuhan dan stok tersedia)
            const taken = Math.min(needed, availableStock);

            // Kekurangan = kebutuhan - yang bisa diambil
            const shortage = Math.max(0, needed - taken);

            if (spkDetails.has(materialId)) {
              const existing = spkDetails.get(materialId)!;
              existing.needed += needed;
              existing.taken += taken;
              existing.shortage += shortage;
            } else {
              spkDetails.set(materialId, {
                needed,
                taken,
                shortage,
                stockBefore: stockReal,
                stockAfter: stockReal - taken,
              });
            }
          }
        }
      }

      // ==================== BUAT PREVIEW DATA DENGAN FORMAT BARU ====================
      // ==================== BUAT PREVIEW DATA DENGAN FORMAT BARU ====================
      const previewMaterialData: any[] = [];

      // BUAT SET UNTUK MENGECEK SPK YANG SEDANG DIEXPORT
      const exportingSPKsSet = new Set(
        selectedOrders.map((order) => order.order.No_SPK),
      );

      for (const agg of materialAggMap.values()) {
        const barangJadiDetails: string[] = [];
        for (const [kode, info] of agg.barangJadiSet) {
          barangJadiDetails.push(`${kode} (${info.qty.toLocaleString()})`);
        }

        // 🔥 PERBAIKAN: Ambil reserved dari PO LAIN dari reservationsByItem
        const reservationDetailsList: string[] = [];

        // Cek apakah material ini memiliki reserved dari PO lain
        const reservedData = reservationsByItem.get(agg.kode);

        if (reservedData && reservedData.spkList.size > 0) {
          // Loop melalui semua SPK yang mereserved material ini
          for (const spkReservation of reservedData.spkList) {
            // namaPO di spkList bisa berupa No_SPK atau Nama_PO
            // Kita perlu mengecek apakah SPK ini SEDANG DIPREVIEW
            let isExportingSPK = false;

            // Cek apakah spkReservation.namaPO matches dengan salah satu No_SPK yang dipreview
            for (const order of selectedOrders) {
              if (
                order.order.No_SPK === spkReservation.namaPO ||
                order.order.Nama_PO === spkReservation.namaPO
              ) {
                isExportingSPK = true;
                break;
              }
            }

            // Hanya tampilkan jika SPK ini TIDAK sedang dipreview
            if (!isExportingSPK) {
              reservationDetailsList.push(
                `${spkReservation.namaPO} (${spkReservation.qtyReserved.toLocaleString()})`,
              );
            }
          }
        }

        const reservedByText =
          reservationDetailsList.length > 0
            ? reservationDetailsList.join("\n")
            : "-";

        const variantInfo = getVariantInfo(agg.kode);

        const totalDibutuhkan = agg.totalNeeded + agg.qtyReserved;
        const available = agg.stockWincp - totalDibutuhkan;
        const kekurangan = totalDibutuhkan - agg.stockWincp;

        let status = "";
        if (agg.stockWincp >= totalDibutuhkan) {
          status = "AMAN";
        } else if (agg.stockWincp > 0) {
          status = "KURANG";
        } else {
          status = "HABIS";
        }

        previewMaterialData.push({
          "Kode Material": agg.kode,
          "Nama Material": agg.nama,
          "Nama China": agg.nama_china,
          Spesifikasi: agg.spec,
          Warna: agg.warna,
          Bahan: agg.bahan,
          Departemen: agg.departemen,
          "Barang Jadi": barangJadiDetails.join("\n"),
          "Total Kebutuhan": agg.totalNeeded.toLocaleString(),
          "Stok Wincp (Real)": agg.stockWincp.toLocaleString(),
          "Saldo Akhir": agg.stockAkhir.toLocaleString(),
          "Qty Reserved (PO Lain)": agg.qtyReserved.toLocaleString(),
          "Total Dibutuhkan": totalDibutuhkan.toLocaleString(),
          "Qty Available": available.toLocaleString(),
          "Reserved Oleh SPK": reservedByText,
          "Keterangan Variant": variantInfo,
          "Status Stock": status,
          Kekurangan: kekurangan > 0 ? kekurangan.toLocaleString() : "0",
        });
      }

      previewMaterialData.sort((a, b) =>
        a["Kode Material"].localeCompare(b["Kode Material"]),
      );

      let fileName = "";
      if (selectedOrders.length === 1) {
        const singlePO = selectedOrders[0];
        const poName = singlePO.order.Nama_PO || singlePO.order.No_SPK;
        fileName = poName
          .replace(/[\\/*?:"<>|]/g, "")
          .replace(/\s+/g, "_")
          .substring(0, 50);
      } else {
        const firstPO = selectedOrders[0];
        const firstPOName = firstPO.order.Nama_PO || firstPO.order.No_SPK;
        fileName = firstPOName
          .replace(/[\\/*?:"<>|]/g, "")
          .replace(/\s+/g, "_")
          .substring(0, 40);
      }

      setPreviewDialog({
        open: true,
        data: previewMaterialData,
        fileName: fileName,
        totalPO: selectedOrders.length,
        totalMaterial: previewMaterialData.length,
      });

      setExportProgress({ visible: false, current: 0, total: 0, message: "" });
      showToast("Preview siap dengan informasi stok lengkap", "success");
    } catch (error) {
      console.error("Error preview:", error);
      showToast(
        `❌ Gagal preview: ${error instanceof Error ? error.message : "Unknown error"}`,
        "error",
      );
    } finally {
      setExportLoading(false);
      setExportProgress({ visible: false, current: 0, total: 0, message: "" });
    }
  };

  // PreviewDialog Component - VERSION WITH SAVE TO DATABASE BUTTON
  const PreviewDialog: React.FC<{
    open: boolean;
    data: any[] | null;
    fileName: string;
    totalPO: number;
    totalMaterial: number;
    onClose: () => void;
    onExport: () => void;
    onSaveToDatabase?: () => void; // TAMBAHKAN PROPS INI
    saving?: boolean; // TAMBAHKAN PROPS INI
  }> = ({
    open,
    data,
    fileName,
    totalPO,
    totalMaterial,
    onClose,
    onExport,
    onSaveToDatabase, // TAMBAHKAN
    saving = false, // TAMBAHKAN
  }) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [sortField, setSortField] = useState<string>("Kode Material");
    const [sortDirection, setSortDirection] = useState<"asc" | "desc">("asc");
    const [stockFilter, setStockFilter] = useState<
      "all" | "aman" | "kurang" | "habis"
    >("all");

    const filteredAndSortedData = useMemo(() => {
      if (!data) return [];

      let filtered = data;

      if (searchTerm) {
        const term = searchTerm.toLowerCase();
        filtered = filtered.filter(
          (item) =>
            item["Kode Material"]?.toLowerCase().includes(term) ||
            item["Nama Material"]?.toLowerCase().includes(term) ||
            item["Departemen"]?.toLowerCase().includes(term) ||
            item["Reserved Oleh SPK"]?.toLowerCase().includes(term),
        );
      }

      if (stockFilter !== "all") {
        filtered = filtered.filter((item) => {
          if (stockFilter === "aman") return item["Status Stock"] === "AMAN";
          if (stockFilter === "kurang")
            return item["Status Stock"] === "KURANG";
          if (stockFilter === "habis") return item["Status Stock"] === "HABIS";
          return true;
        });
      }

      filtered = [...filtered].sort((a, b) => {
        let aVal = a[sortField];
        let bVal = b[sortField];

        if (
          sortField === "Total Kebutuhan" ||
          sortField === "Stok Wincp (Real)" ||
          sortField === "Saldo Akhir" ||
          sortField === "Qty Reserved (PO Lain)" ||
          sortField === "Qty Available" ||
          sortField === "Kekurangan"
        ) {
          aVal = parseInt(String(aVal).replace(/,/g, "")) || 0;
          bVal = parseInt(String(bVal).replace(/,/g, "")) || 0;
          return sortDirection === "asc" ? aVal - bVal : bVal - aVal;
        }

        const comparison = String(aVal).localeCompare(String(bVal));
        return sortDirection === "asc" ? comparison : -comparison;
      });

      return filtered;
    }, [data, searchTerm, sortField, sortDirection, stockFilter]);

    const handleSort = (field: string) => {
      if (sortField === field) {
        setSortDirection(sortDirection === "asc" ? "desc" : "asc");
      } else {
        setSortField(field);
        setSortDirection("asc");
      }
    };

    const getSortIcon = (field: string) => {
      if (sortField !== field)
        return <ChevronsUpDown className="h-3 w-3 ml-1" />;
      return sortDirection === "asc" ? (
        <ChevronUp className="h-3 w-3 ml-1" />
      ) : (
        <ChevronDown className="h-3 w-3 ml-1" />
      );
    };

    if (!open || !data) return null;

    const totalKebutuhan = data.reduce(
      (sum, item) =>
        sum +
        (parseInt(String(item["Total Kebutuhan"]).replace(/,/g, "")) || 0),
      0,
    );

    const totalKekurangan = data.reduce((sum, item) => {
      const kekurangan =
        parseInt(String(item["Kekurangan"]).replace(/,/g, "")) || 0;
      return sum + kekurangan;
    }, 0);

    const materialAman = data.filter(
      (item) => item["Status Stock"] === "AMAN",
    ).length;
    const materialKurang = data.filter(
      (item) => item["Status Stock"] === "KURANG",
    ).length;
    const materialHabis = data.filter(
      (item) => item["Status Stock"] === "HABIS",
    ).length;

    return (
      <Sheet open={open} onOpenChange={onClose}>
        <SheetContent side="bottom" className="h-[90vh] p-0 flex flex-col">
          {/* Header - tidak scroll */}
          <div className="flex-shrink-0 border-b px-6 py-4 bg-white">
            <SheetHeader className="text-left">
              <SheetTitle className="flex items-center gap-2">
                <Eye className="h-5 w-5" />
                Preview Export Data - Material Requirements (Lengkap dengan Cek
                Stok)
              </SheetTitle>
              <SheetDescription className="mt-2">
                {totalPO} PO dipilih | {totalMaterial} material unik | Total
                Kebutuhan: {totalKebutuhan.toLocaleString()}
                {fileName && ` | Nama file: ${fileName}.xlsx`}
              </SheetDescription>
              <div className="flex gap-4 mt-3 text-sm flex-wrap">
                <span className="text-green-600 font-medium">
                  ✅ AMAN: {materialAman}
                </span>
                <span className="text-orange-600 font-medium">
                  ⚠️ KURANG: {materialKurang}
                </span>
                <span className="text-red-600 font-medium">
                  ❌ HABIS: {materialHabis}
                </span>
                {totalKekurangan > 0 && (
                  <span className="text-red-600 font-bold">
                    Total Kekurangan: {totalKekurangan.toLocaleString()}
                  </span>
                )}
              </div>
            </SheetHeader>
          </div>

          {/* Filter Section - tidak scroll */}
          <div className="flex-shrink-0 border-b px-6 py-3 bg-white">
            <div className="flex gap-4 items-center flex-wrap">
              <div className="flex-1 relative min-w-[200px]">
                <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-muted-foreground" />
                <Input
                  placeholder="Cari Kode Material, Nama Material, Departemen, atau SPK..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="pl-8"
                />
              </div>
              <div className="flex gap-2">
                <Button
                  variant={stockFilter === "all" ? "default" : "outline"}
                  size="sm"
                  onClick={() => setStockFilter("all")}
                >
                  Semua
                </Button>
                <Button
                  variant={stockFilter === "aman" ? "default" : "outline"}
                  size="sm"
                  className="text-green-600"
                  onClick={() => setStockFilter("aman")}
                >
                  AMAN
                </Button>
                <Button
                  variant={stockFilter === "kurang" ? "default" : "outline"}
                  size="sm"
                  className="text-orange-600"
                  onClick={() => setStockFilter("kurang")}
                >
                  KURANG
                </Button>
                <Button
                  variant={stockFilter === "habis" ? "default" : "outline"}
                  size="sm"
                  className="text-red-600"
                  onClick={() => setStockFilter("habis")}
                >
                  HABIS
                </Button>
              </div>
              <div className="text-sm text-muted-foreground whitespace-nowrap">
                {filteredAndSortedData.length} / {data.length}
              </div>
            </div>
          </div>

          {/* Table - YANG SCROLL */}
          <div className="flex-1 overflow-auto px-6 py-4 min-h-0">
            <div className="border rounded-md h-full">
              <div
                className="overflow-auto"
                style={{ maxHeight: "calc(90vh - 280px)" }}
              >
                <Table>
                  <TableHeader className="sticky top-0 bg-gray-50 z-10">
                    <TableRow>
                      <TableHead
                        className="min-w-[150px] cursor-pointer"
                        onClick={() => handleSort("Kode Material")}
                      >
                        Kode Material {getSortIcon("Kode Material")}
                      </TableHead>
                      <TableHead
                        className="min-w-[200px] cursor-pointer"
                        onClick={() => handleSort("Nama Material")}
                      >
                        Nama Material {getSortIcon("Nama Material")}
                      </TableHead>
                      <TableHead
                        className="min-w-[150px] cursor-pointer"
                        onClick={() => handleSort("Departemen")}
                      >
                        Departemen {getSortIcon("Departemen")}
                      </TableHead>
                      <TableHead className="min-w-[200px]">
                        Barang Jadi
                      </TableHead>
                      <TableHead
                        className="text-right min-w-[120px] cursor-pointer"
                        onClick={() => handleSort("Total Kebutuhan")}
                      >
                        Kebutuhan {getSortIcon("Total Kebutuhan")}
                      </TableHead>
                      <TableHead
                        className="text-right min-w-[140px] cursor-pointer"
                        onClick={() => handleSort("Qty Reserved (PO Lain)")}
                      >
                        Reserved (PO Lain){" "}
                        {getSortIcon("Qty Reserved (PO Lain)")}
                      </TableHead>
                      <TableHead
                        className="text-right min-w-[120px] cursor-pointer"
                        onClick={() => handleSort("Total Dibutuhkan")}
                      >
                        Total Dibutuhkan {getSortIcon("Total Dibutuhkan")}
                      </TableHead>
                      <TableHead
                        className="text-right min-w-[120px] cursor-pointer"
                        onClick={() => handleSort("Stok Wincp (Real)")}
                      >
                        Stok Wincp {getSortIcon("Stok Wincp (Real)")}
                      </TableHead>
                      <TableHead
                        className="text-right min-w-[120px] cursor-pointer"
                        onClick={() => handleSort("Saldo Akhir")}
                      >
                        Saldo Akhir {getSortIcon("Saldo Akhir")}
                      </TableHead>
                      <TableHead
                        className="text-right min-w-[120px] cursor-pointer"
                        onClick={() => handleSort("Qty Available")}
                      >
                        Available {getSortIcon("Qty Available")}
                      </TableHead>
                      <TableHead
                        className="min-w-[250px] cursor-pointer"
                        onClick={() => handleSort("Reserved Oleh SPK")}
                      >
                        Reserved Oleh PO {getSortIcon("Reserved Oleh SPK")}
                      </TableHead>
                      <TableHead
                        className="min-w-[100px] cursor-pointer"
                        onClick={() => handleSort("Status Stock")}
                      >
                        Status {getSortIcon("Status Stock")}
                      </TableHead>
                      <TableHead
                        className="text-right min-w-[100px] cursor-pointer"
                        onClick={() => handleSort("Kekurangan")}
                      >
                        Kekurangan {getSortIcon("Kekurangan")}
                      </TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {filteredAndSortedData.map((item, idx) => (
                      <TableRow
                        key={idx}
                        className={
                          item["Status Stock"] === "HABIS"
                            ? "bg-red-50 hover:bg-red-100"
                            : item["Status Stock"] === "KURANG"
                              ? "bg-orange-50 hover:bg-orange-100"
                              : ""
                        }
                      >
                        <TableCell className="font-mono text-sm">
                          {item["Kode Material"]}
                        </TableCell>
                        <TableCell className="text-sm">
                          {item["Nama Material"]}
                        </TableCell>
                        <TableCell className="text-sm">
                          {item["Departemen"]}
                        </TableCell>
                        <TableCell className="text-sm whitespace-pre-wrap max-w-[250px] break-words">
                          {item["Barang Jadi"]}
                        </TableCell>
                        <TableCell className="text-right font-mono font-bold">
                          {item["Total Kebutuhan"]}
                        </TableCell>
                        <TableCell className="text-right font-mono text-orange-600">
                          {item["Qty Reserved (PO Lain)"]}
                        </TableCell>
                        <TableCell className="text-right font-mono font-bold text-purple-600">
                          {item["Total Dibutuhkan"]}
                        </TableCell>
                        <TableCell className="text-right font-mono">
                          {item["Stok Wincp (Real)"]}
                        </TableCell>
                        <TableCell className="text-right font-mono">
                          {item["Saldo Akhir"]}
                        </TableCell>
                        <TableCell className="text-right font-mono">
                          {item["Qty Available"]}
                        </TableCell>
                        <TableCell className="text-sm whitespace-pre-wrap max-w-[250px] break-words">
                          {item["Reserved Oleh SPK"] !== "-"
                            ? item["Reserved Oleh SPK"]
                            : "-"}
                        </TableCell>
                        <TableCell
                          className={`text-sm font-bold ${
                            item["Status Stock"] === "AMAN"
                              ? "text-green-600"
                              : item["Status Stock"] === "KURANG"
                                ? "text-orange-600"
                                : "text-red-600"
                          }`}
                        >
                          {item["Status Stock"]}
                        </TableCell>
                        <TableCell className="text-right font-mono text-red-600 font-bold">
                          {item["Kekurangan"] !== "-"
                            ? item["Kekurangan"]
                            : "0"}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            </div>
          </div>

          {/* Footer - PERBAIKI DENGAN MENAMBAHKAN TOMBOL SAVE KE DATABASE */}
          <div className="flex-shrink-0 border-t px-6 py-4 bg-white">
            <div className="flex justify-between items-center flex-wrap gap-2">
              <div className="text-sm text-muted-foreground">
                Total Kebutuhan:{" "}
                <span className="font-bold">
                  {totalKebutuhan.toLocaleString()}
                </span>
                {totalKekurangan > 0 && (
                  <span className="ml-4 text-red-600">
                    | Total Kekurangan:{" "}
                    <span className="font-bold">
                      {totalKekurangan.toLocaleString()}
                    </span>
                  </span>
                )}
                <span className="ml-2 text-blue-600">
                  | ✅ Dengan cek stok & reserved PO lain
                </span>
              </div>
              <div className="flex gap-2">
                <Button variant="outline" onClick={onClose}>
                  Tutup
                </Button>

                {/* TAMBAHKAN TOMBOL SAVE KE DATABASE */}
                {onSaveToDatabase && (
                  <Button
                    onClick={onSaveToDatabase}
                    disabled={saving}
                    variant="secondary"
                    className="gap-2"
                  >
                    {saving ? (
                      <Loader2 className="h-4 w-4 animate-spin" />
                    ) : (
                      <Database className="h-4 w-4" />
                    )}
                    {saving ? "Menyimpan..." : "Simpan ke Database"}
                  </Button>
                )}

                <Button onClick={onExport} className="gap-2">
                  <Download className="h-4 w-4" />
                  Export ke Excel
                </Button>
              </div>
            </div>
          </div>
        </SheetContent>
      </Sheet>
    );
  };

  const paginatedOrders = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return filteredOrders.slice(startIndex, startIndex + itemsPerPage);
  }, [filteredOrders, currentPage, itemsPerPage]);

  const totalPages = Math.ceil(filteredOrders.length / itemsPerPage);
  const selectedCount = orders.filter((p) => p.selected).length;
  const totalActiveOrders = orders.filter((p) => !p.committed).length;
  const committedCount = orders.filter((p) => p.committed).length;

  useEffect(() => {
    (refreshAllData(), loadSavedCalculations());
  }, []);
  useEffect(() => {
    if (isMounted) {
      refreshAllData();
    }
  }, [showCommitted]);
  // Prevent hydration mismatch
  if (!isMounted) {
    return (
      <div className="container mx-auto py-6">
        <Card>
          <CardContent className="pt-6">
            <div className="flex justify-center p-8">Loading...</div>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="container mx-auto py-6 space-y-6">
      {/* Toast Notification */}
      {toast && (
        <ToastNotification
          message={toast.message}
          type={toast.type}
          onClose={hideToast}
        />
      )}

      {/* Export Progress Dialog */}
      <Dialog open={exportProgress.visible}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <RefreshCw className="h-5 w-5 animate-spin" /> Sedang Mengekspor
              Data...
            </DialogTitle>
            <DialogDescription>
              Mohon tunggu, proses export sedang berjalan
            </DialogDescription>
          </DialogHeader>
          <div className="space-y-4">
            <p className="text-sm text-muted-foreground">
              {exportProgress.message}
            </p>
            <Progress
              value={(exportProgress.current / exportProgress.total) * 100}
              className="h-2"
            />
          </div>
        </DialogContent>

        {/* Preview Dialog */}
        <PreviewDialog
          open={previewDialog.open}
          data={previewDialog.data}
          fileName={previewDialog.fileName}
          totalPO={previewDialog.totalPO}
          totalMaterial={previewDialog.totalMaterial}
          onClose={() =>
            setPreviewDialog({
              open: false,
              data: null,
              fileName: "",
              totalPO: 0,
              totalMaterial: 0,
            })
          }
          onExport={() => {
            setPreviewDialog({
              open: false,
              data: null,
              fileName: "",
              totalPO: 0,
              totalMaterial: 0,
            });
            exportSelectedToExcel();
          }}
          onSaveToDatabase={handleSaveToDatabase} // TAMBAHKAN INI
          saving={savingToDatabase} // TAMBAHKAN INI
        />
      </Dialog>

      {/* Header */}
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-bold flex items-center gap-2">
            <Factory className="h-8 w-8" /> Production Planning
          </h1>
          <p className="text-muted-foreground text-sm mt-1">
            Commit PO akan mereserve stok | Uncommit akan mengembalikan stok
          </p>
        </div>
        {/* Header buttons - update teks */}
        <div className="flex gap-2">
          <Button onClick={refreshAllData} disabled={loading} variant="outline">
            <RefreshCw className={`h-4 w-4 ${loading ? "animate-spin" : ""}`} />{" "}
            Refresh
          </Button>
          <Button
            onClick={previewExport}
            disabled={exportLoading || selectedCount === 0}
            variant="outline"
            className="gap-2"
          >
            <Eye className="h-4 w-4" />
            Preview ({selectedCount} PO)
          </Button>
          <Button
            onClick={exportSelectedToExcel}
            disabled={exportLoading || selectedCount === 0}
            className="gap-2"
          >
            <Download className="h-4 w-4" />
            Export ({selectedCount} PO)
          </Button>
          <Button
            onClick={() =>
              setBomOverrideDialog({
                open: true,
                originalItemId: "",
                originalItemName: "",
                currentReplacement: "",
              })
            }
            variant="outline"
            className="gap-2"
          >
            <TreePine className="h-4 w-4" />
            Perubahan BOM ({bomOverrides.filter((o) => o.isActive).length})
          </Button>

          {/* Tambahkan dialog di akhir komponen (sebelum closing div) */}
          {bomOverrideDialog.open && (
            <BomOverrideManager
              overrides={bomOverrides}
              onAdd={addBomOverride}
              onUpdate={updateBomOverride}
              onDelete={deleteBomOverride}
              onToggle={toggleBomOverride}
              onClose={() =>
                setBomOverrideDialog({
                  open: false,
                  originalItemId: "",
                  originalItemName: "",
                  currentReplacement: "",
                })
              }
              loading={loadingOverrides}
            />
          )}
        </div>
      </div>

      {/* Committed PO Panel */}
      {/* <CommittedPOsPanel
        committedPOs={committedPOs}
        stockReservations={stockReservations}
        onRefresh={refreshAllData}
        onUncommit={uncommitPO}
      /> */}
      <HistoryPanel
        calculations={savedCalculations}
        onLoadCalculation={loadCalculationToPreview}
        onRefresh={loadSavedCalculations}
        onDeleteCalculation={deleteCalculation}
      />
      {/* Search */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Search className="h-5 w-5" /> Pencarian PO
          </CardTitle>
        </CardHeader>
        <CardContent>
          <Input
            value={searchQuery}
            onChange={handleSearchChange}
            placeholder="Cari No SPK, Nama PO, atau Kode Barang..."
          />
          {searchQuery && (
            <Button variant="link" onClick={clearSearch} className="mt-2">
              Clear
            </Button>
          )}
        </CardContent>
      </Card>

      {/* Date Filter */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Calendar className="h-5 w-5" /> Filter Tanggal
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="flex flex-wrap gap-4">
            <div className="flex-1">
              <Label>Tanggal Mulai</Label>
              <Input
                type="date"
                name="startDate"
                value={dateFilter.startDate}
                onChange={handleDateFilterChange}
              />
            </div>
            <div className="flex-1">
              <Label>Tanggal Akhir</Label>
              <Input
                type="date"
                name="endDate"
                value={dateFilter.endDate}
                onChange={handleDateFilterChange}
              />
            </div>
            <div className="flex gap-2 items-end">
              <Button onClick={applyDateFilter} disabled={loading}>
                Filter
              </Button>
              <Button variant="outline" onClick={resetDateFilter}>
                Reset
              </Button>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Statistics */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <Card>
          <CardContent className="pt-6">
            <div className="text-sm text-muted-foreground">Total Order</div>
            <div className="text-2xl font-bold">{filteredOrders.length}</div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="pt-6">
            <div className="text-sm text-muted-foreground">Active PO</div>
            <div className="text-2xl font-bold">{totalActiveOrders}</div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="pt-6">
            <div className="text-sm font-medium text-muted-foreground">
              Terpilih
            </div>
            <div className="text-2xl font-bold text-primary">
              {selectedCount} / {totalActiveOrders}
            </div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="pt-6">
            <div className="text-sm text-muted-foreground">Committed</div>
            <div className="text-2xl font-bold text-green-600">
              {committedCount}
            </div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="pt-6">
            <div className="text-sm text-muted-foreground">PO Digabung</div>
            <div className="text-2xl font-bold text-purple-600">
              {
                orders.filter(
                  (o) =>
                    o.order.combinedItems && o.order.combinedItems.length > 1,
                ).length
              }
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Selection Controls */}
      {/* Di dalam JSX, sekitar selection controls */}
      <Card>
        <CardHeader className="flex flex-row justify-between items-center flex-wrap gap-2">
          <CardTitle className="flex items-center gap-2">
            <Package className="h-5 w-5" /> Daftar Production Order
          </CardTitle>
          <div className="flex gap-2">
            {/* 🔥 TOMBOL TOGGLE SHOW COMMITTED */}
            <Button
              variant={showCommitted ? "default" : "outline"}
              size="sm"
              onClick={() => {
                console.log(
                  "Toggling showCommitted from",
                  showCommitted,
                  "to",
                  !showCommitted,
                );
                setShowCommitted(!showCommitted);
                // Tidak perlu panggil refreshAllData karena useEffect akan handle
                // Tapi jika useEffect tidak jalan, panggil manual:
                // refreshAllData();
              }}
              className="gap-2"
            >
              <Eye className="h-4 w-4" />
              {showCommitted ? "Sembunyikan Committed" : "Tampilkan Committed"}
            </Button>
            <Button variant="outline" onClick={toggleSelectAll}>
              Select Page
            </Button>
            <Button onClick={toggleSelectAllGlobal}>
              Select All ({totalActiveOrders})
            </Button>
          </div>
        </CardHeader>
      </Card>

      {/* Error */}
      {error && (
        <Alert variant="destructive">
          <AlertCircle className="h-4 w-4" />
          <AlertTitle>Error</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      )}

      {/* Loading */}
      {loading && (
        <Card>
          <CardContent className="pt-6">
            <div className="flex justify-center p-8">
              <RefreshCw className="h-8 w-8 animate-spin" />
            </div>
          </CardContent>
        </Card>
      )}

      {/* Table - HANYA MENAMPILKAN PO YANG BELUM DI-COMMIT */}
      {!loading && filteredOrders.length > 0 && (
        <>
          <Card>
            <div className="w-full overflow-x-auto">
              <Table className="min-w-200 lg:min-w-full">
                <TableHeader>
                  <TableRow>
                    <TableHead className="w-12.5">Pilih</TableHead>
                    <TableHead className="w-32.5">No SPK</TableHead>
                    <TableHead className="w-25">Tanggal Order</TableHead>
                    <TableHead className="w-27.5">Nama PO</TableHead>

                    <TableHead className="w-30 text-center">Aksi</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {paginatedOrders.map((plan, idx) => {
                    const isCombined =
                      plan.order.combinedItems &&
                      plan.order.combinedItems.length > 1;
                    const hasBom = !!plan.bom;
                    const isCommitting = committing === plan.order.No_SPK;

                    return (
                      <TableRow
                        key={`${plan.order.No_SPK}-${idx}`}
                        className={`
    ${isCombined ? "bg-purple-50" : ""}
    ${plan.committed ? "bg-green-50 opacity-70" : ""}
  `}
                      >
                        <TableCell className="w-12.5 align-top">
                          <Checkbox
                            checked={plan.selected}
                            onCheckedChange={() => toggleSelection(idx)}
                            // disabled={plan.committed}
                          />
                        </TableCell>

                        <TableCell className="w-32.5 align-top font-medium">
                          <div className="whitespace-nowrap">
                            {plan.order.No_SPK}
                            {plan.committed && (
                              <Badge
                                variant="default"
                                className="ml-2 bg-green-500 text-white"
                              >
                                COMMITTED
                              </Badge>
                            )}
                          </div>
                          <div className="flex flex-wrap gap-1 mt-1">
                            {isCombined && (
                              <Badge
                                variant="secondary"
                                className="whitespace-nowrap"
                              >
                                {plan.order.combinedItems?.length} Item
                              </Badge>
                            )}
                            {hasBom && (
                              <Badge
                                variant="outline"
                                className="bg-blue-100 text-blue-800 whitespace-nowrap"
                              >
                                BOM Ready
                              </Badge>
                            )}
                            {plan.committed && plan.CommitID && (
                              <Badge
                                variant="outline"
                                className="bg-green-100 text-green-800 whitespace-nowrap"
                              >
                                ID: {plan.CommitID}
                              </Badge>
                            )}
                          </div>
                        </TableCell>

                        <TableCell className="w-25 align-top whitespace-nowrap">
                          {plan.order.Tanggal_Order}
                        </TableCell>

                        <TableCell className="align-top">
                          <div
                            className="font-medium truncate max-w-62.5"
                            title={plan.order.Nama_PO}
                          >
                            {plan.order.Nama_PO}
                          </div>
                          {isCombined && plan.order.combinedItems && (
                            <div className="text-xs text-muted-foreground mt-1 space-y-0.5">
                              {plan.order.combinedItems.map((item, i) => (
                                <div
                                  key={i}
                                  className="truncate max-w-50"
                                  title={`${item.Kode_Barang} (QTY: ${item.QTY})`}
                                >
                                  • {item.Kode_Barang} (QTY: {item.QTY})
                                </div>
                              ))}
                            </div>
                          )}
                        </TableCell>
                        <TableCell className="w-30 align-top text-center">
                          {plan.committed ? (
                            <div className="flex gap-2 justify-center">
                              <Button
                                variant="secondary"
                                size="sm"
                                disabled
                                className="gap-1 whitespace-nowrap"
                              >
                                <Lock className="h-3 w-3" />
                                Committed
                              </Button>
                              <Button
                                variant="destructive"
                                size="sm"
                                onClick={() => uncommitPO(plan.order.No_SPK)}
                                disabled={
                                  isCommitting &&
                                  committing === plan.order.No_SPK
                                }
                                className="gap-1 whitespace-nowrap"
                              >
                                {isCommitting &&
                                committing === plan.order.No_SPK ? (
                                  <Loader2 className="h-3 w-3 animate-spin" />
                                ) : (
                                  <Unlock className="h-3 w-3" />
                                )}
                                {isCommitting &&
                                committing === plan.order.No_SPK
                                  ? "Uncommitting..."
                                  : "Uncommit"}
                              </Button>
                            </div>
                          ) : (
                            <Button
                              variant="default"
                              size="sm"
                              onClick={() => commitPO(idx)}
                              disabled={
                                isCommitting && committing === plan.order.No_SPK
                              }
                              className="gap-1 whitespace-nowrap"
                            >
                              {isCommitting &&
                              committing === plan.order.No_SPK ? (
                                <Loader2 className="h-3 w-3 animate-spin" />
                              ) : (
                                <Lock className="h-3 w-3" />
                              )}
                              {isCommitting && committing === plan.order.No_SPK
                                ? "Committing..."
                                : "Commit PO"}
                            </Button>
                          )}
                        </TableCell>
                      </TableRow>
                    );
                  })}
                </TableBody>
              </Table>
            </div>
          </Card>
          {/* Pagination */}
          {totalPages > 1 && (
            <div className="flex justify-between items-center">
              <div className="text-sm text-muted-foreground">
                Page {currentPage} of {totalPages}
              </div>
              <div className="flex gap-2">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                >
                  Previous
                </Button>
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() =>
                    setCurrentPage((p) => Math.min(totalPages, p + 1))
                  }
                  disabled={currentPage === totalPages}
                >
                  Next
                </Button>
              </div>
            </div>
          )}
        </>
      )}

      {!loading && filteredOrders.length === 0 && (
        <Card>
          <CardContent className="pt-6">
            <div className="text-center p-8">Tidak ada data produksi aktif</div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

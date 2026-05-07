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
import { Progress }from "@/components/ui/progress"
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
  ChevronUp, ChevronDown, ChevronsUpDown
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
}

interface StockItem {
  itemid: string;
  itemname: string;
  stockAkhir: number;      // Untuk SaldoAkhir
  physicalStock: number;   // Untuk SaldoAkhirFisik (Stok Wincp)
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
  namaPO?:string;
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
      let saldoAkhir = 0;      // Untuk kolom "Stok Akhir"
      let saldoAkhirFisik = 0; // Untuk kolom "Stok Wincp"

      // Ambil SaldoAkhir
      if (typeof itemData.SaldoAkhir === "number" && itemData.SaldoAkhir !== undefined) {
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

      console.log(`Stock untuk ${itemId}: SaldoAkhir=${saldoAkhir}, SaldoAkhirFisik=${saldoAkhirFisik}`);

      return {
        itemid: normalizeItemId(
          itemData.KodeBarang || itemData.itemid || itemData.ItemID || itemId,
        ),
        itemname:
          itemData.NamaBarang ||
          itemData.itemname ||
          itemData.ItemName ||
          itemId,
        stockAkhir: saldoAkhir,        // Untuk kolom "Stok Akhir"
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
        showCommitted, // Tambahkan parameter showCommitted
      );
      showToast("Data berhasil dimuat ulang", "success");
    } catch (error) {
      console.error("Error refreshing data:", error);
      showToast("Gagal memuat ulang data", "error");
    } finally {
      setLoading(false);
    }
  }, [dateFilter.startDate, dateFilter.endDate, showCommitted]); // Tambahkan showCommitted ke dependensi

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
        const treeStructure = buildTreeStructure(bomResponse.data.flat);
        combinedBoms[kb] = { flat: bomResponse.data.flat, tree: treeStructure };
        allItemIds.push(
          ...bomResponse.data.flat.map((item: BomItem) =>
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

      // Ambil data BOM untuk committed PO
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
          const treeStructure = buildTreeStructure(bomResponse.data.flat);
          combinedBoms[kb] = {
            flat: bomResponse.data.flat,
            tree: treeStructure,
          };
          allItemIds.push(
            ...bomResponse.data.flat.map((item: BomItem) =>
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

      console.log(
        `Processing order ${i + 1}/${totalOrders}: ${order.order.No_SPK}, committed: ${order.committed}`,
      );

      // Jika sudah punya BOM dan stock, skip
      if (order.bom && order.stock) {
        setExportProgress((prev) => ({
          ...prev,
          current: Math.floor(((i + 1) / totalOrders) * 50),
          message: `Menggunakan BOM yang sudah ada untuk ${order.order.No_SPK} (${i + 1}/${totalOrders})...`,
        }));
        console.log(`Using existing BOM for ${order.order.No_SPK}`);
        continue;
      }

      // 🔥 TAMBAHKAN: Jika PO sudah di-commit, coba load dari database terlebih dahulu
      if (order.committed) {
        setExportProgress({
          visible: true,
          current: Math.floor((i / totalOrders) * 50),
          total: 100,
          message: `Memuat data untuk committed PO ${order.order.No_SPK} (${i + 1}/${totalOrders})...`,
        });
        console.log(`Loading committed order data for ${order.order.No_SPK}`);

        const loadedOrder = await loadCommittedOrderData(order);
        updatedOrders[i] = loadedOrder;
        console.log(
          `Loaded committed order data for ${order.order.No_SPK}, hasBom: ${!!loadedOrder.bom}`,
        );
        continue;
      }

      // Untuk PO yang belum di-commit, load seperti biasa
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
          const treeStructure = buildTreeStructure(bomResponse.data.flat);
          combinedBoms[kb] = {
            flat: bomResponse.data.flat,
            tree: treeStructure,
          };
          allItemIds.push(
            ...bomResponse.data.flat.map((item: BomItem) =>
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

    console.log(`Finished loading BOM for ${updatedOrders.length} orders`);
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

  // ==================== FUNGSI EXPORT (DENGAN NAMA FILE BERDASARKAN PO) ====================

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
      const selectedOrders = filteredOrders.filter(
        (order) => order.selected,
        // && !order.committed,
      );
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
        // Jika hanya 1 PO, gunakan nama PO tersebut
        const singlePO = selectedOrders[0];
        const poName = singlePO.order.Nama_PO || singlePO.order.No_SPK;
        // Bersihkan karakter yang tidak valid untuk nama file
        const cleanFileName = poName
          .replace(/[\\/*?:"<>|]/g, "") // Hapus karakter tidak valid
          .replace(/\s+/g, "_") // Ganti spasi dengan underscore
          .substring(0, 50); // Batasi panjang nama
        fileName = `${cleanFileName}_${today}.xlsx`;
      } else {
        // Jika lebih dari 1 PO, gunakan nama PO pertama + dll
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

      // ==================== SHEET 2: BOM (FINAL - DENGAN PARENT TRACKING) ====================
      const bomData: any[] = [];
      let totalINJECTIONRemoved = 0;

      // ==================== FUNGSI CALCULATE ACCUMULATED QTY ====================
      const calculateAccumulatedQty = (
        flatBom: BomItem[],
      ): Map<string, number> => {
        const cache = new Map<string, number>();

        // Buat map untuk akses cepat
        const itemMap = new Map<string, BomItem>();
        for (const item of flatBom) {
          itemMap.set(normalizeItemId(item.ItemID), item);
        }

        // Hitung accumulated untuk setiap item
        for (const item of flatBom) {
          const itemId = normalizeItemId(item.ItemID);
          const level = Number(item.Level);

          if (level === 1) {
            cache.set(itemId, item.Qty);
          } else {
            // Cari parent berdasarkan ParentItemID atau level
            let parent: BomItem | undefined = undefined;

            // Coba cari berdasarkan ParentItemID dulu
            if (item.ParentItemID) {
              parent = itemMap.get(normalizeItemId(item.ParentItemID));
            }

            // Jika tidak ada ParentItemID, cari berdasarkan level
            if (!parent) {
              parent = flatBom.find((p) => Number(p.Level) === level - 1);
            }

            if (parent) {
              const parentId = normalizeItemId(parent.ItemID);
              const parentAccumulated = cache.get(parentId);
              if (parentAccumulated !== undefined) {
                const accumulated = item.Qty * parentAccumulated;
                cache.set(itemId, accumulated);
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

      // Build tree structure
      const buildTree = (flatBom: BomItem[]): BomItem[] => {
        if (!flatBom || flatBom.length === 0) return [];

        // Buat map untuk node
        const nodeMap = new Map<string, BomItem>();
        const rootItems: BomItem[] = [];

        // Buat node untuk setiap item
        for (const item of flatBom) {
          const node: BomItem = { ...item, children: [] };
          nodeMap.set(normalizeItemId(item.ItemID), node);
        }

        // Bangun parent-child relationship
        for (const item of flatBom) {
          const node = nodeMap.get(normalizeItemId(item.ItemID));
          if (!node) continue;

          const level = Number(item.Level);

          if (level === 1) {
            rootItems.push(node);
          } else {
            // Cari parent
            let parent: BomItem | undefined = undefined;

            // Coba cari berdasarkan ParentItemID
            if (item.ParentItemID) {
              parent = nodeMap.get(normalizeItemId(item.ParentItemID));
            }

            // Jika tidak ada, cari berdasarkan level
            if (!parent) {
              parent = flatBom.find((p) => Number(p.Level) === level - 1);
              if (parent) {
                parent = nodeMap.get(normalizeItemId(parent.ItemID));
              }
            }

            if (parent) {
              if (!parent.children) parent.children = [];
              parent.children.push(node);
            } else {
              rootItems.push(node);
            }
          }
        }

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

      // Format nama dengan indentasi
      const formatIndentedName = (itemName: string, level: number): string => {
        if (level === 1) return `📦 ${itemName}`;
        const indent = "  ".repeat(level - 1);
        return `${indent}└─ ${itemName}`;
      };

      // ==================== PROSES SETIAP ORDER ====================
      for (const order of ordersWithBom) {
        if (!order.bom) continue;
        const isCombined =
          order.order.combinedItems && order.order.combinedItems.length > 1;

        const processBom = (bomFlat: BomItem[], poQty: number, poItem: any) => {
          // Filter komponen
          const filteredBom = bomFlat.filter(
            (b) => Number(b.Level) > 0 && !isINJECTIONDepartment(b.Departemen),
          );
          const removedCount = bomFlat.filter(
            (b) => Number(b.Level) > 0 && isINJECTIONDepartment(b.Departemen),
          ).length;
          totalINJECTIONRemoved += removedCount;

          if (filteredBom.length === 0) return;

          // HITUNG ACCUMULATED QTY
          const accumulatedMap = calculateAccumulatedQty(bomFlat);

          // HEADER
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

          // Build tree
          const treeStructure = buildTree(filteredBom);

          // TRAVERSE TREE
          const traverseTree = (nodes: BomItem[]) => {
            for (const node of nodes) {
              const nodeLevel = Number(node.Level);
              const nodeId = normalizeItemId(node.ItemID);
              const stockItem = order.stock?.find(
                (s) => normalizeItemId(s.itemid) === nodeId,
              );

              const accumulatedQty = accumulatedMap.get(nodeId) || node.Qty;
              const totalNeeded = accumulatedQty * poQty;
              const stock = stockItem?.stockAkhir || 0;
              const shortage = totalNeeded > stock;

              // Keterangan
              let calculationNote = "";
              if (nodeLevel === 1) {
                calculationNote = `Qty per Unit × QTY PO = ${node.Qty} × ${poQty} = ${totalNeeded}`;
              } else {
                calculationNote = `Qty per Unit × Accumulated Parent × QTY PO = ${node.Qty} × ${accumulatedQty / node.Qty} × ${poQty} = ${totalNeeded}`;
              }

              bomData.push({
                "No SPK": "",
                "Kode Barang Jadi": "",
                "Nama Barang Jadi": "",
                "QTY PO": "",
                Level: node.Level,
                "Kode Komponen": node.ItemID,
                "Nama Komponen": formatIndentedName(node.ItemName, nodeLevel),
                "Nama Komponen China": node.ItemName2 || "",
                "Qty per Unit (BOM)": node.Qty,
                "Accumulated Qty": accumulatedQty,
                "Total Kebutuhan": totalNeeded,
                Stok: stock,
                Status: shortage ? "KURANG" : "CUKUP",
                "Keterangan Perhitungan Accumulated": calculationNote,
              });

              if (node.children && node.children.length > 0) {
                traverseTree(node.children);
              }
            }
          };

          traverseTree(treeStructure);
          bomData.push({});
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

      // INFORMASI TAMBAHAN
      bomData.push({});
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": "📦 = Produk Level 1",
      });
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": "  └─ = Sub-komponen Level 2",
      });
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": "    └─ = Sub-komponen Level 3",
      });
      bomData.push({
        "No SPK": "INFORMASI",
        "Nama Komponen": `* Komponen dengan departemen INJECTION tidak ditampilkan (${totalINJECTIONRemoved} item dihapus)`,
      });

      const wsBOM = XLSX.utils.json_to_sheet(bomData);
      wsBOM["!cols"] = [
        { wch: 12 },
        { wch: 15 },
        { wch: 30 },
        { wch: 10 },
        { wch: 8 },
        { wch: 15 },
        { wch: 50 },
        { wch: 35 },
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 12 },
        { wch: 10 },
        { wch: 60 },
      ];
      XLSX.utils.book_append_sheet(wb, wsBOM, "BOM");
      // ==================== SHEET 3: TOTAL KEBUTUHAN MATERIAL (DIPERBAIKI) ====================

      // ==================== DI DALAM FUNGSI exportSelectedToExcel ====================

      // ==================== SHEET 3: TOTAL KEBUTUHAN MATERIAL (DIPERBAIKI) ====================

      // Buat Map untuk mencari Nama PO berdasarkan CommitID (lebih cepat)
      const poNameByCommitID = new Map<number, string>();
      const poNoSPKByCommitID = new Map<number, string>();

      for (const committedPO of committedPOs) {
        if (committedPO.status === "COMMITTED") {
          poNameByCommitID.set(committedPO.CommitID, committedPO.namaPO);
          poNoSPKByCommitID.set(committedPO.CommitID, committedPO.noSPK);
        }
      }

      // Kemudian reservationsByItem
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

        // 🔥 Langsung pakai namaPO dari reservation (sudah ada dari JOIN)
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

      // ==================== FUNGSI CALCULATE ACCUMULATED QTY UNTUK MATERIAL (LEVEL-BASED) ====================
      const calculateAccumulatedQtyForMaterial = (
        flatBom: BomItem[],
      ): Map<string, number> => {
        const cache = new Map<string, number>();

        // Kelompokkan item berdasarkan level
        const itemsByLevel = new Map<number, BomItem[]>();
        for (const item of flatBom) {
          const level = Number(item.Level);
          if (!itemsByLevel.has(level)) itemsByLevel.set(level, []);
          itemsByLevel.get(level)!.push(item);
        }

        // Hitung accumulated qty untuk setiap item
        for (const item of flatBom) {
          const level = Number(item.Level);
          const itemId = normalizeItemId(item.ItemID);

          if (level === 1) {
            cache.set(itemId, item.Qty);
          } else {
            // Kalikan dengan semua Qty dari level 1 sampai level-1
            let accumulated = item.Qty;
            for (let l = level - 1; l >= 1; l--) {
              const parents = itemsByLevel.get(l);
              if (parents && parents.length > 0) {
                accumulated = accumulated * parents[0].Qty;
              }
            }
            cache.set(itemId, accumulated);
          }
        }

        return cache;
      };

      const materialDataRows: any[][] = [];
      const headers = [
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

          // Filter komponen (Level > 0) dan exclude INJECTION
          const components = bomFlat.filter(
            (b) => Number(b.Level) > 0 && !isINJECTIONDepartment(b.Departemen),
          );

          if (components.length === 0) continue;

          // 🔥 HITUNG ACCUMULATED QTY UNTUK BOM INI (LEVEL-BASED)
          const accumulatedCache = calculateAccumulatedQtyForMaterial(bomFlat);
          const accumulatedMap = calculateAccumulatedQty(bomFlat);
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

          // Di dalam loop material aggregation
          for (const [materialId, needed] of tempNeeds) {
            const stockItem = order.stock?.find(
              (s) => normalizeItemId(s.itemid) === materialId,
            );

            // 🔥 BEDAKAN: stockWincp pakai physicalStock (SaldoAkhirFisik)
            // 🔥 stockAkhir pakai stockAkhir (SaldoAkhir)
            const stockWincp = stockItem?.physicalStock || 0; // Untuk kolom Stok Wincp
            const stockAkhir = stockItem?.stockAkhir || 0; // Untuk kolom Stok Akhir

            const masterInfo = masterDataMap.get(materialId) || {
              spec: "-",
              warna: "-",
              bahan: "-",
            };

            const reservedData = reservationsByItem.get(materialId);
            const reservedQty = reservedData?.totalQty || 0;
            const reservedByText = reservedData
              ? Array.from(reservedData.spkList)
                  .map(
                    (item) =>
                      `${item.namaPO} (QTY : ${item.qtyReserved.toLocaleString()})`,
                  )
                  .join("\n")
              : "-";
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
                stockWincp: stockWincp, // SaldoAkhirFisik
                stockAkhir: stockAkhir, // SaldoAkhir
                reserved: reservedQty,
                reservedBy: reservedByText,
                totalNeeded: 0,
                barangJadiSet: new Map(),
              });
            }
            const agg = materialAggMap.get(materialId);
            agg.totalNeeded += needed;

            // Update stock values (ambil yang terbaru jika ada multiple)
            if (stockWincp > agg.stockWincp) agg.stockWincp = stockWincp;
            if (stockAkhir > agg.stockAkhir) agg.stockAkhir = stockAkhir;

            agg.reserved = reservedQty;
            agg.reservedBy = reservedByText;

            if (!agg.barangJadiSet.has(barangJadi.kode))
              agg.barangJadiSet.set(barangJadi.kode, {
                qty: barangJadi.qty,
                nama: barangJadi.nama,
              });
          }
        }
      }

      // Buat material data rows
      // Buat material data rows
      for (const agg of materialAggMap.values()) {
        const barangJadiDetails: string[] = [];
        const qtyPODetails: string[] = [];

        for (const [kode, info] of agg.barangJadiSet) {
          barangJadiDetails.push(`${kode}`);
          qtyPODetails.push(info.qty.toLocaleString());
        }

        const totalDibutuhkan = agg.totalNeeded + agg.reserved;
        const sisaStok = agg.stockWincp - totalDibutuhkan; // 🔥 Pakai stockAkhir ()
        let status = sisaStok > 0 ? "CUKUP" : sisaStok < 0 ? "KURANG" : "HABIS";

        const variantInfo = getVariantInfo(agg.kode);

        materialDataRows.push([
          barangJadiDetails.join("\n"),
          qtyPODetails.join("\n"),
          agg.kode,
          agg.nama,
          agg.nama_china,
          agg.spec,
          agg.warna,
          agg.bahan,
          agg.departemen,
          agg.totalNeeded, // Total Kebutuhan
          agg.reserved, // Reserved (Qty PO Lain)
          totalDibutuhkan, // Total Dibutuhkan
          agg.stockWincp, // 🔥 Stok Wincp (SaldoAkhirFisik)
          agg.stockAkhir, // 🔥 Stok Akhir (SaldoAkhir)
          sisaStok, // Sisa Stok
          agg.reservedBy, // Reserved Oleh SPK
          status, // Status
          variantInfo, // Keterangan Variant
        ]);
      }

      // Urutkan berdasarkan kode material
      materialDataRows.sort((a, b) => a[2].localeCompare(b[2]));

      // ==================== SHEET PER DEPARTEMEN ====================
      const materialsByDept = new Map<string, Map<string, any[]>>();

      for (const row of materialDataRows) {
        const dept = row[8] || "UNKNOWN";
        const materialCode = row[2];
        if (!materialsByDept.has(dept)) materialsByDept.set(dept, new Map());
        const deptMap = materialsByDept.get(dept)!;

        if (!deptMap.has(materialCode)) {
          deptMap.set(materialCode, [...row]);
        } else {
          const existing = deptMap.get(materialCode)!;
          // Total Kebutuhan (indeks 9) - dijumlah
          existing[9] = (existing[9] || 0) + (row[9] || 0);
          // Total Dibutuhkan (indeks 11) - dijumlah
          existing[11] = (existing[11] || 0) + (row[11] || 0);
          // Sisa Stok (indeks 14) - dihitung ulang
          existing[14] = (existing[13] || 0) - (existing[11] || 0);
          // Status (indeks 16) - update
          existing[16] =
            existing[14] > 0
              ? "KELEBIHAN"
              : existing[14] < 0
                ? "KURANG"
                : "CUKUP";

          // Gabungkan Barang Jadi (indeks 0)
          const existingBarangJadi = existing[0] || "";
          const newBarangJadi = row[0] || "";
          if (
            newBarangJadi &&
            !existingBarangJadi.includes(newBarangJadi.split("\n")[0])
          ) {
            existing[0] =
              existingBarangJadi +
              (existingBarangJadi ? "\n" : "") +
              newBarangJadi;
          }

          // Gabungkan Reserved Oleh SPK (indeks 15)
          const existingReserved = existing[15] || "";
          const newReserved = row[15] || "";
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

          // Gabungkan Keterangan Variant (indeks 17)
          const existingVariant = existing[17] || "";
          const newVariant = row[17] || "";
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

      // Konversi ke array dan urutkan
      const finalMaterialsByDept = new Map<string, any[][]>();
      for (const [dept, materialMap] of materialsByDept) {
        const rows: any[][] = [];
        for (const row of materialMap.values()) rows.push(row);
        rows.sort((a, b) => a[2].localeCompare(b[2]));
        finalMaterialsByDept.set(dept, rows);
      }

      const sortedDepartments = Array.from(finalMaterialsByDept.keys()).sort();
      const deptColWidths = [
        { wch: 50 }, // Barang Jadi
        { wch: 20 }, // QTY PO Dipesan
        { wch: 15 }, // Kode Material
        { wch: 40 }, // Nama Material
        { wch: 35 }, // Nama China
        { wch: 30 }, // Spesifikasi
        { wch: 20 }, // Warna
        { wch: 25 }, // Bahan
        { wch: 20 }, // Departemen
        { wch: 15 }, // Total Kebutuhan
        { wch: 15 }, // Reserved
        { wch: 15 }, // Total Dibutuhkan
        { wch: 15 }, // Stok Wincp
        { wch: 15 }, // Stok Akhir
        { wch: 15 }, // Sisa Stok
        { wch: 50 }, // Reserved Oleh SPK
        { wch: 15 }, // Status
        { wch: 25 }, // Keterangan Variant
      ];

      // Buat sheet per departemen
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

        // Daftar item yang memiliki variant di departemen ini
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
          headers,
          ...deptMaterials,
          [],
          [
            `Total Keseluruhan: ${deptMaterials.length} material unik, ` +
              `Total Kebutuhan: ${totalNeeded.toLocaleString()}, ` +
              `Total Sisa Stok: ${totalSisa.toLocaleString()}`,
          ],
        ];

        // Filter baris kosong jika variantNote tidak ada
        const finalWsData = variantNote
          ? wsData
          : wsData.filter((_, idx) => idx !== 4);

        const wsDept = XLSX.utils.aoa_to_sheet(finalWsData);
        wsDept["!cols"] = deptColWidths;

        // Merge cells untuk header
        wsDept["!merges"] = [
          { s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } },
          { s: { r: 1, c: 0 }, e: { r: 1, c: headers.length - 1 } },
          { s: { r: 2, c: 0 }, e: { r: 2, c: headers.length - 1 } },
          { s: { r: 3, c: 0 }, e: { r: 3, c: headers.length - 1 } },
        ];

        if (variantNote) {
          wsDept["!merges"].push({
            s: { r: 4, c: 0 },
            e: { r: 4, c: headers.length - 1 },
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

      const totalAllMaterials = materialDataRows.length;
      const totalAllNeeded = materialDataRows.reduce(
        (sum, row) => sum + (row[9] || 0),
        0,
      );
      const totalAllSisa = materialDataRows.reduce(
        (sum, row) => sum + (row[14] || 0),
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
        { wch: 25 }, // Departemen
        { wch: 18 }, // Jumlah Material
        { wch: 20 }, // Total Kebutuhan
        { wch: 20 }, // Total Sisa Stok
        { wch: 20 }, // Status
      ];
      wsSummary["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 4 } },
      ];

      XLSX.utils.book_append_sheet(wb, wsSummary, "REKAP_PER_DEPARTEMEN");
      // ==================== SHEET KETERANGAN (DI PALING AKHIR) ====================
      // Buat sheet keterangan untuk menjelaskan setiap kolom dalam format tabel
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
        ["3", "Total Material Unik", materialDataRows.length],
        ["4", "Tanggal Export", new Date().toLocaleDateString("id-ID")],
        ["5", "Waktu Export", new Date().toLocaleTimeString("id-ID")],
        ["6", "User", "PPIC Department"],
        ["7", "Aplikasi", "Production Planning System"],
      ];

      const wsKeterangan = XLSX.utils.aoa_to_sheet(keteranganData);

      // Set column widths untuk sheet keterangan
      wsKeterangan["!cols"] = [
        { wch: 8 }, // Kolom A (No)
        { wch: 30 }, // Kolom B (Item/Kolom/Rumus)
        { wch: 50 }, // Kolom C (Keterangan)
        { wch: 35 }, // Kolom D (Contoh/Nilai)
      ];

      // Merge cells untuk judul
      wsKeterangan["!merges"] = [
        // Judul utama
        { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
        // Header section A
        { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } },
        // Header section B
        { s: { r: 8, c: 0 }, e: { r: 8, c: 3 } },
        // Header section C
        { s: { r: 16, c: 0 }, e: { r: 16, c: 3 } },
        // Header section D
        { s: { r: 27, c: 0 }, e: { r: 27, c: 4 } },
        // Header section E
        { s: { r: 32, c: 0 }, e: { r: 32, c: 3 } },
        // Header section F
        { s: { r: 52, c: 0 }, e: { r: 52, c: 3 } },
        // Header section G
        { s: { r: 59, c: 0 }, e: { r: 59, c: 3 } },
        // Header section H
        { s: { r: 66, c: 0 }, e: { r: 66, c: 3 } },
        // Header section I
        { s: { r: 74, c: 0 }, e: { r: 74, c: 2 } },
        // Header section J
        { s: { r: 85, c: 0 }, e: { r: 85, c: 2 } },
      ];

      // Terapkan border style untuk tabel (opsional, menggunakan karakter untuk membuat tabel sederhana)
      // Atau bisa menggunakan format default Excel

      // 🔥 TAMBAHKAN SHEET KETERANGAN DI PALING AKHIR
      XLSX.utils.book_append_sheet(wb, wsKeterangan, "KETERANGAN");
      // Gunakan fileName yang sudah dibuat
      XLSX.writeFile(wb, fileName);

      setExportProgress({ visible: false, current: 0, total: 0, message: "" });
      showToast(
        `✅ Export berhasil! ${selectedOrders.length} PO diexport`,
        "success",
      );
      alert(
        `✅ Export berhasil!\nFile: ${fileName}\n\n📦 Total PO: ${selectedOrders.length}\n📅 Stok per tanggal: ${today}`,
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

  // ==================== FUNGSI PREVIEW EXPORT (DIPERBAIKI) ====================
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

     // 🔥 BUAT SET UNTUK MENAMPUNG SPK YANG SEDANG DIPREVIEW
     const exportingSPKs = new Set(
       selectedOrders.map((order) => order.order.No_SPK),
     );
     console.log(
       "Preview SPKs (excluding self-reservation):",
       Array.from(exportingSPKs),
     );

     // 🔥 HANYA SATU DEKLARASI reservationsByItem (dengan filter)
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

       // 🔥 LEWATKAN RESERVATION UNTUK PO YANG SEDANG DIPREVIEW
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
           tempNeeds.set(materialId, (tempNeeds.get(materialId) || 0) + needed);
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

           const reservedByText = reservedData
             ? Array.from(reservedData.spkList)
                 .map(
                   (item) =>
                     `${item.namaPO} (${item.qtyReserved.toLocaleString()})`,
                 )
                 .join("\n")
             : "-";

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
               reservedBy: reservedByText,
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

     // 🔥 BUAT PREVIEW DATA
     const previewMaterialData: any[] = [];

     for (const agg of materialAggMap.values()) {
       const barangJadiDetails: string[] = [];
       for (const [kode, info] of agg.barangJadiSet) {
         barangJadiDetails.push(`${kode} (${info.qty.toLocaleString()})`);
       }

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
         "Reserved Oleh SPK": agg.reservedBy,
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

  const PreviewDialog: React.FC<{
    open: boolean;
    data: any[] | null;
    fileName: string;
    totalPO: number;
    totalMaterial: number;
    onClose: () => void;
    onExport: () => void;
  }> = ({
    open,
    data,
    fileName,
    totalPO,
    totalMaterial,
    onClose,
    onExport,
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
                        Reserved Oleh SPK {getSortIcon("Reserved Oleh SPK")}
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

          {/* Footer - tidak scroll */}
          <div className="flex-shrink-0 border-t px-6 py-4 bg-white">
            <div className="flex justify-between items-center">
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
    refreshAllData();
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
        </div>
      </div>

      {/* Committed PO Panel */}
      <CommittedPOsPanel
        committedPOs={committedPOs}
        stockReservations={stockReservations}
        onRefresh={refreshAllData}
        onUncommit={uncommitPO}
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
              <Table className="min-w-[800px] lg:min-w-full">
                <TableHeader>
                  <TableRow>
                    <TableHead className="w-[50px]">Pilih</TableHead>
                    <TableHead className="w-[130px]">No SPK</TableHead>
                    <TableHead className="w-[100px]">Tanggal Order</TableHead>
                    <TableHead className="w-[110px]">Nama PO</TableHead>

                    <TableHead className="w-[120px] text-center">
                      Aksi
                    </TableHead>
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
                        <TableCell className="w-[50px] align-top">
                          <Checkbox
                            checked={plan.selected}
                            onCheckedChange={() => toggleSelection(idx)}
                            // disabled={plan.committed}
                          />
                        </TableCell>

                        <TableCell className="w-[130px] align-top font-medium">
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

                        <TableCell className="w-[100px] align-top whitespace-nowrap">
                          {plan.order.Tanggal_Order}
                        </TableCell>

                        <TableCell className="align-top">
                          <div
                            className="font-medium truncate max-w-[250px]"
                            title={plan.order.Nama_PO}
                          >
                            {plan.order.Nama_PO}
                          </div>
                          {isCombined && plan.order.combinedItems && (
                            <div className="text-xs text-muted-foreground mt-1 space-y-0.5">
                              {plan.order.combinedItems.map((item, i) => (
                                <div
                                  key={i}
                                  className="truncate max-w-[200px]"
                                  title={`${item.Kode_Barang} (QTY: ${item.QTY})`}
                                >
                                  • {item.Kode_Barang} (QTY: {item.QTY})
                                </div>
                              ))}
                            </div>
                          )}
                        </TableCell>

                        <TableCell className="w-[120px] align-top text-center">
                          {plan.committed ? (
                            <Button
                              variant="secondary"
                              size="sm"
                              disabled
                              className="gap-1 whitespace-nowrap"
                            >
                              <Lock className="h-3 w-3" />
                              Committed
                            </Button>
                          ) : (
                            <Button
                              variant="default"
                              size="sm"
                              onClick={() => commitPO(idx)}
                              disabled={isCommitting}
                              className="gap-1 whitespace-nowrap"
                            >
                              {isCommitting ? (
                                <Loader2 className="h-3 w-3 animate-spin" />
                              ) : (
                                <Lock className="h-3 w-3" />
                              )}
                              {isCommitting ? "Committing..." : "Commit PO"}
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

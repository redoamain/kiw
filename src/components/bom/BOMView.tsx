/* eslint-disable @typescript-eslint/no-explicit-any */
"use client";

import { useState, useCallback, useRef, useMemo, useEffect } from "react";
import axios from "axios";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import {
  Search,
  RefreshCw,
  ChevronRight,
  Package,
  Layers,
  Hash,
  FileText,
  Filter,
  Download,
  Box,
  ListTree,
  Copy,
  Eye,
  EyeOff,
  X,
  CheckCircle2,
  AlertTriangle,
  Info,
  Edit,
  Trash2,
  Save,
  Plus,
  Loader2,
  CheckCircle,
  Printer,
  Palette,
  Ruler,
} from "lucide-react";
import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Alert, AlertDescription } from "@/components/ui/alert";
import {
  Tooltip,
  TooltipContent,
  TooltipProvider,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogFooter,
  DialogDescription,
} from "@/components/ui/dialog";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import * as XLSX from "xlsx";
import { toast } from "sonner";

interface BomData {
  TransID: number;
  itemidHD: string;
  itemnamehd: string;
  itemnamehd2?: string;
  ItemID: string;
  ItemName: string;
  ItemName2: string;
  BahanQty: number;
  BahanPackSatuan?: string;
  Departemen?: string;
  NamaJenis?: string;
  Spec?: string;
  warna?: string;
  bahan?: string;
}

interface MasterItem {
  ItemID: string;
  ItemName: string;
  ItemName2?: string;
  SatuanKecil?: string;
  warna?: string;
  Spec?: string;
}

interface Notification {
  show: boolean;
  message: string;
  type: "success" | "error" | "info";
}

const satuanOptions = [
  { value: "pcs", label: "PCS" },
  { value: "kg", label: "KG" },
  { value: "gram", label: "GRAM" },
  { value: "liter", label: "LITER" },
  { value: "meter", label: "METER" },
  { value: "cm", label: "CM" },
  { value: "roll", label: "ROLL" },
  { value: "set", label: "SET" },
  { value: "box", label: "BOX" },
  { value: "pack", label: "PACK" },
];

export default function BOMView() {
  const [mounted, setMounted] = useState(false);
  const [itemid, setItemid] = useState("");
  const [data, setData] = useState<BomData[]>([]);
  const [loading, setLoading] = useState(false);
  const [searchType, setSearchType] = useState<"itemid" | "itemname">("itemid");
  const [isExporting, setIsExporting] = useState(false);
  const [isPrinting, setIsPrinting] = useState(false);
  const [viewMode, setViewMode] = useState<"all" | "filtered">("all");
  const [expandedAll, setExpandedAll] = useState<string[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);
  const [loadProgress, setLoadProgress] = useState({
    current: 0,
    total: 100,
    visible: false,
  });
  const [notification, setNotification] = useState<Notification>({
    show: false,
    message: "",
    type: "success",
  });

  // State untuk checkbox pilihan produk
  const [selectedProducts, setSelectedProducts] = useState<Set<string>>(
    new Set(),
  );

  // Edit state
  const [editDialogOpen, setEditDialogOpen] = useState(false);
  const [editingComponent, setEditingComponent] = useState<BomData | null>(
    null,
  );
  const [editItemID, setEditItemID] = useState<string>("");
  const [editQuantity, setEditQuantity] = useState<number>(0);
  const [editSatuan, setEditSatuan] = useState<string>("");
  const [isUpdating, setIsUpdating] = useState(false);

  // Add component state
  const [addDialogOpen, setAddDialogOpen] = useState(false);
  const [selectedProduct, setSelectedProduct] = useState<{
    itemidHD: string;
    itemnamehd: string;
    TransID: number;
  } | null>(null);
  const [newItemID, setNewItemID] = useState<string>("");
  const [newQuantity, setNewQuantity] = useState<number>(1);
  const [newSatuan, setNewSatuan] = useState<string>("pcs");
  const [isAdding, setIsAdding] = useState(false);

  // Suggest state untuk add component
  const [itemSuggestions, setItemSuggestions] = useState<MasterItem[]>([]);
  const [showItemSuggestions, setShowItemSuggestions] = useState(false);
  const [isCheckingItem, setIsCheckingItem] = useState(false);
  const [selectedItemData, setSelectedItemData] = useState<MasterItem | null>(
    null,
  );
  const [suggestLoading, setSuggestLoading] = useState(false);

  // Delete confirmation
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
  const [deletingComponent, setDeletingComponent] = useState<BomData | null>(
    null,
  );

  // Pagination for display
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(20);
  const [isAllDataMode, setIsAllDataMode] = useState(false);

  const searchTimeoutRef = useRef<NodeJS.Timeout | null>(null);
  const suggestTimeoutRef = useRef<NodeJS.Timeout | null>(null);

  // Fix hydration mismatch
  useEffect(() => {
    setMounted(true);
  }, []);

  const showNotification = (
    message: string,
    type: "success" | "error" | "info" = "success",
  ) => {
    setNotification({ show: true, message, type });
    setTimeout(() => {
      setNotification({ show: false, message: "", type: "success" });
    }, 3000);
  };

  // API Calls
  const updateBomComponent = async (
    TransID: string,
    ItemID: string,
    newItemID: string,
    BahanQty: number,
    BahanPackSatuan: string,
  ) => {
    const response = await axios.put("/api/bom", {
      TransID,
      ItemID,
      newItemID,
      BahanQty,
      BahanPackSatuan,
      action: "update",
    });
    return response.data;
  };

  const deleteBomComponent = async (TransID: string, ItemID: string) => {
    const response = await axios.put("/api/bom", {
      TransID,
      ItemID,
      action: "delete",
    });
    return response.data;
  };

  const addBomComponent = async (
    TransID: string,
    ItemID: string,
    BahanQty: number,
    BahanPackSatuan: string,
  ) => {
    const response = await axios.patch("/api/bom", {
      action: "add",
      TransID,
      components: [{ ItemID, BahanQty, BahanPackSatuan }],
    });
    return response.data;
  };

  // Group data
  const grouped = useMemo(() => {
    if (data.length === 0) return {};

    const result: Record<
      string,
      {
        headerName: string;
        headerName2: string;
        children: BomData[];
        totalComponents: number;
      }
    > = {};

    for (const row of data) {
      if (!result[row.itemidHD]) {
        result[row.itemidHD] = {
          headerName: row.itemnamehd,
          headerName2: row.itemnamehd2 || "",
          children: [],
          totalComponents: 0,
        };
      }

      const exists = result[row.itemidHD].children.some(
        (child) => child.ItemID === row.ItemID,
      );

      if (!exists) {
        result[row.itemidHD].children.push(row);
        result[row.itemidHD].totalComponents++;
      }
    }

    return result;
  }, [data]);

  const paginatedGrouped = useMemo(() => {
    const entries = Object.entries(grouped);
    const startIndex = (currentPage - 1) * itemsPerPage;
    const paginatedEntries = entries.slice(
      startIndex,
      startIndex + itemsPerPage,
    );
    return Object.fromEntries(paginatedEntries);
  }, [grouped, currentPage, itemsPerPage]);

  const totalHeaders = useMemo(() => Object.keys(grouped).length, [grouped]);
  const totalPages = Math.ceil(totalHeaders / itemsPerPage);
  const totalComponentsAll = useMemo(() => data.length, [data]);

  // Hitung jumlah produk yang dipilih di halaman ini
  const selectedCountInPage = useMemo(() => {
    const pageProductIds = Object.keys(paginatedGrouped);
    return pageProductIds.filter((id) => selectedProducts.has(id)).length;
  }, [paginatedGrouped, selectedProducts]);

  // Hitung total semua produk yang dipilih
  const totalSelectedCount = selectedProducts.size;

  // Fetch BOM data
  const fetchBOM = useCallback(
    async (searchItem?: string, fetchAllData: boolean = false) => {
      if (fetchAllData) {
        setLoadProgress({ current: 0, total: 100, visible: true });
        setIsAllDataMode(true);
      } else {
        setIsAllDataMode(false);
      }

      setLoading(true);
      setCurrentPage(1);
      setSelectedProducts(new Set());

      const progressInterval = fetchAllData
        ? setInterval(() => {
            setLoadProgress((prev) => ({
              ...prev,
              current: Math.min(prev.current + 5, 90),
            }));
          }, 300)
        : null;

      try {
        const params: any = {};

        if (searchItem && searchItem.trim() !== "") {
          params.itemid = searchItem;
          params.searchType = searchType;
        }

        if (fetchAllData) {
          params.fetchAll = true;
        }

        const res = await axios.get(`/api/bom`, { params });

        let bomData: BomData[] = [];

        if (res.data.data) {
          bomData = res.data.data;
        } else {
          bomData = res.data;
        }

        const uniqueMap = new Map<string, BomData>();
        for (const item of bomData) {
          const uniqueKey = `${item.itemidHD}|${item.ItemID}`;
          if (!uniqueMap.has(uniqueKey)) {
            uniqueMap.set(uniqueKey, item);
          }
        }

        bomData = Array.from(uniqueMap.values());
        setData(bomData);
        setViewMode(
          searchItem && searchItem.trim() !== "" ? "filtered" : "all",
        );
        setHasSearched(true);

        if (progressInterval) {
          clearInterval(progressInterval);
          setLoadProgress({ current: 100, total: 100, visible: true });
          setTimeout(
            () => setLoadProgress((prev) => ({ ...prev, visible: false })),
            1000,
          );
        }

        const headers = Array.from(
          new Set(bomData.map((item) => item.itemidHD)),
        );
        if (headers.length > 10) {
          setExpandedAll(headers.slice(0, 5));
        } else {
          setExpandedAll(headers);
        }

        showNotification(
          `Berhasil memuat ${bomData.length} komponen dari ${headers.length} produk${fetchAllData ? " (semua data)" : ""}`,
          "success",
        );
      } catch (err) {
        console.error("Gagal ambil BOM:", err);
        setData([]);
        setHasSearched(true);
        showNotification("Gagal memuat data BOM", "error");
        if (progressInterval) clearInterval(progressInterval);
        setLoadProgress({ current: 0, total: 100, visible: false });
      } finally {
        setLoading(false);
        setIsSearching(false);
      }
    },
    [searchType],
  );

  // Handle select product
  const handleSelectProduct = useCallback(
    (productId: string, checked: boolean) => {
      setSelectedProducts((prev) => {
        const newSet = new Set(prev);
        if (checked) {
          newSet.add(productId);
        } else {
          newSet.delete(productId);
        }
        return newSet;
      });
    },
    [],
  );

  // Handle select all in current page
  const handleSelectAllInPage = useCallback(() => {
    const pageProductIds = Object.keys(paginatedGrouped);
    if (selectedCountInPage === pageProductIds.length) {
      setSelectedProducts((prev) => {
        const newSet = new Set(prev);
        pageProductIds.forEach((id) => newSet.delete(id));
        return newSet;
      });
    } else {
      setSelectedProducts((prev) => {
        const newSet = new Set(prev);
        pageProductIds.forEach((id) => newSet.add(id));
        return newSet;
      });
    }
  }, [paginatedGrouped, selectedCountInPage]);

  // Handle select all products
  const handleSelectAllProducts = useCallback(() => {
    const allProductIds = Object.keys(grouped);
    if (totalSelectedCount === allProductIds.length) {
      setSelectedProducts(new Set());
    } else {
      setSelectedProducts(new Set(allProductIds));
    }
  }, [grouped, totalSelectedCount]);

  // Export to Excel
  const exportSelectedToExcel = useCallback(async () => {
    if (selectedProducts.size === 0) {
      toast.error("Pilih minimal 1 produk untuk diexport");
      return;
    }

    setIsExporting(true);
    setLoadProgress({ current: 0, total: 100, visible: true });

    await new Promise((resolve) => setTimeout(resolve, 50));

    try {
      const today = new Date();
      const timestamp = `${today.toISOString().split("T")[0]}_${today.toLocaleTimeString("id-ID").replace(/:/g, "-")}`;
      const fileName = `BOM_Selected_${selectedProducts.size}Products_${timestamp}.xlsx`;

      const worksheetData: any[][] = [];

      worksheetData.push(["LAPORAN BILL OF MATERIALS (BOM)"]);
      worksheetData.push([
        `Tanggal Export: ${today.toLocaleDateString("id-ID")} ${today.toLocaleTimeString("id-ID")}`,
      ]);
      worksheetData.push([
        `Jumlah Produk Dipilih: ${selectedProducts.size} produk`,
      ]);

      const totalComponentsSelected = Array.from(selectedProducts).reduce(
        (sum, productId) => sum + (grouped[productId]?.totalComponents || 0),
        0,
      );
      worksheetData.push([`Total Komponen: ${totalComponentsSelected}`]);
      worksheetData.push([]);

      const columnHeaders = [
        "No",
        "Komponen ID",
        "Nama Komponen",
        "Nama China",
        "Spec",
        "Warna",
        "Bahan",
        "Jumlah",
        "Satuan",
      ];

      const productIds = Array.from(selectedProducts);

      for (let i = 0; i < productIds.length; i++) {
        const productId = productIds[i];
        const group = grouped[productId];

        if (!group) continue;

        setLoadProgress({
          current: Math.round((i / productIds.length) * 100),
          total: 100,
          visible: true,
        });

        await new Promise((resolve) => setTimeout(resolve, 5));

        worksheetData.push([`PRODUK: ${productId}`]);
        worksheetData.push([`Nama Produk: ${group.headerName}`]);
        if (group.headerName2 && group.headerName2 !== group.headerName) {
          worksheetData.push([`Nama Produk (China): ${group.headerName2}`]);
        }
        worksheetData.push([`Total Komponen: ${group.totalComponents}`]);
        worksheetData.push([]);
        worksheetData.push(columnHeaders);

        for (let j = 0; j < group.children.length; j++) {
          const row = group.children[j];
          worksheetData.push([
            j + 1,
            row.ItemID,
            row.ItemName,
            row.ItemName2 || "-",
            row.Spec || "-",
            row.warna || "-",
            row.bahan || "-",
            row.BahanQty,
            row.BahanPackSatuan || "-",
          ]);
        }

        worksheetData.push([]);
        worksheetData.push([]);
      }

      worksheetData.push([]);
      worksheetData.push([
        `TOTAL KESELURUHAN: ${selectedProducts.size} produk, ${totalComponentsSelected} komponen`,
      ]);

      const ws = XLSX.utils.aoa_to_sheet(worksheetData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "BOM Selected");

      ws["!cols"] = [
        { wch: 8 },
        { wch: 20 },
        { wch: 45 },
        { wch: 45 },
        { wch: 30 },
        { wch: 20 },
        { wch: 25 },
        { wch: 12 },
        { wch: 10 },
      ];

      XLSX.writeFile(wb, fileName);

      toast.success(
        `Berhasil export ${selectedProducts.size} produk ke: ${fileName}`,
      );
    } catch (error) {
      console.error("❌ Gagal export Excel:", error);
      toast.error("Gagal mengexport data");
    } finally {
      setIsExporting(false);
      setLoadProgress({ current: 0, total: 100, visible: false });
    }
  }, [selectedProducts, grouped]);

  // Print selected products
  const printSelectedProducts = useCallback(async () => {
    if (selectedProducts.size === 0) {
      toast.error("Pilih minimal 1 produk untuk dicetak");
      return;
    }

    setIsPrinting(true);
    setLoadProgress({ current: 0, total: 100, visible: true });

    await new Promise((resolve) => setTimeout(resolve, 50));

    try {
      const today = new Date();
      const productIds = Array.from(selectedProducts);

      let totalComponentsPrint = 0;
      productIds.forEach((productId) => {
        const group = grouped[productId];
        if (group) totalComponentsPrint += group.totalComponents;
      });

      setLoadProgress({ current: 50, total: 100, visible: true });
      await new Promise((resolve) => setTimeout(resolve, 100));

      const printWindow = window.open("", "_blank");
      if (!printWindow) {
        toast.error(
          "Pop-up blocker mencegah pencetakan. Izinkan pop-up untuk website ini.",
        );
        setIsPrinting(false);
        setLoadProgress({ current: 0, total: 100, visible: false });
        return;
      }

      const printHtml = `<!DOCTYPE html>
<html>
<head>
  <title>BOM Report - ${today.toLocaleDateString("id-ID")}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { 
      font-family: 'Segoe UI', Arial, sans-serif; 
      padding: 8px; 
      background: white;
      font-size: 11px;
    }
    .header { 
      text-align: center; 
      margin-bottom: 10px; 
      padding-bottom: 5px; 
      border-bottom: 1px solid #ccc;
    }
    .header h1 { font-size: 16px; margin-bottom: 2px; }
    .header p { color: #666; font-size: 10px; }
    .product-section { 
      margin-bottom: 15px; 
      page-break-inside: avoid;
      border: 1px solid #ddd;
      border-radius: 4px;
      overflow: hidden;
    }
    .product-header {
      background: #f0f0f0;
      padding: 6px 10px;
      border-bottom: 1px solid #ddd;
    }
    .product-header h2 { 
      font-size: 13px; 
      margin-bottom: 2px;
      font-family: monospace;
    }
    .product-header p { color: #555; font-size: 10px; margin-top: 2px; }
    .product-stats {
      display: inline-block;
      background: #e0e0e0;
      padding: 1px 6px;
      border-radius: 3px;
      font-size: 9px;
      margin-left: 6px;
    }
    table { 
      width: 100%; 
      border-collapse: collapse; 
      font-size: 9px;
    }
    th, td { 
      border: 1px solid #ddd; 
      padding: 4px 4px; 
      text-align: left; 
    }
    th { 
      background: #f5f5f5; 
      font-weight: bold;
    }
    td:first-child, th:first-child { text-align: center; width: 30px; }
    .footer {
      margin-top: 15px;
      padding-top: 5px;
      border-top: 1px solid #ccc;
      text-align: center;
      font-size: 9px;
      color: #999;
    }
    @media print {
      body { padding: 0; margin: 0; }
      .product-section { break-inside: avoid; margin-bottom: 10px; }
      .no-break { break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>📋 BILL OF MATERIALS (BOM) REPORT</h1>
    <p>Tanggal Cetak: ${today.toLocaleDateString("id-ID")} ${today.toLocaleTimeString("id-ID")}</p>
    <p>Jumlah Produk: ${productIds.length} | Total Komponen: ${totalComponentsPrint}</p>
  </div>
  ${productIds
    .map((productId) => {
      const group = grouped[productId];
      if (!group) return "";
      return `
      <div class="product-section no-break">
        <div class="product-header">
          <h2>
            ${productId}
            <span class="product-stats">${group.totalComponents} komponen</span>
          </h2>
          <p><strong>${group.headerName}</strong></p>
          ${group.headerName2 && group.headerName2 !== group.headerName ? `<p style="color:#666;font-size:9px">${group.headerName2}</p>` : ""}
        </div>
        <table>
          <thead>
            <tr>
              <th>No</th>
              <th>ID Komponen</th>
              <th>Nama Komponen</th>
              <th>Nama China</th>
              <th>Spec</th>
              <th>Warna</th>
              <th>Bahan</th>
              <th>Jumlah</th>
              <th>Satuan</th>
            </tr>
          </thead>
          <tbody>
            ${group.children
              .map(
                (row, idx) => `
              <tr>
                <td>${idx + 1}</td>
                <td><strong>${row.ItemID}</strong></td>
                <td>${row.ItemName}</td>
                <td>${row.ItemName2 || "-"}</td>
                <td>${row.Spec || "-"}</td>
                <td>${row.warna || "-"}</td>
                <td>${row.bahan || "-"}</td>
                <td>${row.BahanQty.toLocaleString()}</td>
                <td>${row.BahanPackSatuan || "-"}</td>
              </tr>
            `,
              )
              .join("")}
          </tbody>
        </table>
      </div>
    `;
    })
    .join("")}
  <div class="footer">
    <p>Dicetak dari sistem BOM Management</p>
  </div>
  <script>
    window.onload = () => { window.print(); setTimeout(() => window.close(), 500); };
  </script>
</body>
</html>`;

      printWindow.document.write(printHtml);
      printWindow.document.close();

      setLoadProgress({ current: 100, total: 100, visible: true });
      setTimeout(
        () => setLoadProgress((prev) => ({ ...prev, visible: false })),
        1000,
      );

      toast.success(`Berhasil mencetak ${selectedProducts.size} produk`);
    } catch (error) {
      console.error("❌ Gagal cetak:", error);
      toast.error("Gagal mencetak data");
    } finally {
      setIsPrinting(false);
      setLoadProgress({ current: 0, total: 100, visible: false });
    }
  }, [selectedProducts, grouped]);

  // Handle edit component
  const handleEditClick = (component: BomData) => {
    setEditingComponent(component);
    setEditItemID(component.ItemID);
    setEditQuantity(component.BahanQty);
    setEditSatuan(component.BahanPackSatuan || "");
    setSelectedItemData(null);
    setEditDialogOpen(true);
  };

  const handleUpdateComponent = async () => {
    if (!editingComponent) return;

    if (!editItemID.trim()) {
      toast.error("ID Komponen harus diisi");
      return;
    }

    if (editQuantity <= 0) {
      toast.error("Jumlah harus lebih dari 0");
      return;
    }

    setIsUpdating(true);
    try {
      const response = await updateBomComponent(
        editingComponent.TransID.toString(),
        editingComponent.ItemID,
        editItemID,
        editQuantity,
        editSatuan,
      );

      if (response.success) {
        await fetchBOM(undefined, true);
        toast.success(`Berhasil update komponen`);
        setEditDialogOpen(false);
        setEditingComponent(null);
        setSelectedItemData(null);
      } else {
        toast.error(response.message || "Gagal update komponen");
      }
    } catch (error: any) {
      console.error("Error updating:", error);
      toast.error(error.response?.data?.error || "Gagal update komponen");
    } finally {
      setIsUpdating(false);
    }
  };

  // Handle add component
  const handleAddClick = (
    productId: string,
    productName: string,
    TransID: number,
  ) => {
    setSelectedProduct({
      itemidHD: productId,
      itemnamehd: productName,
      TransID,
    });
    setNewItemID("");
    setNewQuantity(1);
    setNewSatuan("pcs");
    setSelectedItemData(null);
    setItemSuggestions([]);
    setShowItemSuggestions(false);
    setAddDialogOpen(true);
  };

  const handleAddComponent = async () => {
    if (!selectedProduct) return;

    if (!newItemID.trim()) {
      toast.error("ID Komponen harus diisi");
      return;
    }

    if (newQuantity <= 0) {
      toast.error("Jumlah harus lebih dari 0");
      return;
    }

    setIsAdding(true);
    try {
      const response = await addBomComponent(
        selectedProduct.TransID.toString(),
        newItemID,
        newQuantity,
        newSatuan,
      );

      if (response.success) {
        await fetchBOM(undefined, true);
        toast.success(`Berhasil menambahkan komponen ${newItemID}`);
        setAddDialogOpen(false);
        setSelectedProduct(null);
        setNewItemID("");
        setNewQuantity(1);
        setNewSatuan("pcs");
        setSelectedItemData(null);
      } else {
        toast.error(response.message || "Gagal menambahkan komponen");
      }
    } catch (error: any) {
      console.error("Error adding component:", error);
      toast.error(error.response?.data?.error || "Gagal menambahkan komponen");
    } finally {
      setIsAdding(false);
    }
  };

  // Handle delete component
  const handleDeleteClick = (component: BomData) => {
    setDeletingComponent(component);
    setDeleteDialogOpen(true);
  };

  const handleDeleteComponent = async () => {
    if (!deletingComponent) return;

    setIsUpdating(true);
    try {
      const response = await deleteBomComponent(
        deletingComponent.TransID.toString(),
        deletingComponent.ItemID,
      );

      if (response.success) {
        await fetchBOM(undefined, true);
        toast.success(`Berhasil hapus ${deletingComponent.ItemName}`);
        setDeleteDialogOpen(false);
        setDeletingComponent(null);
      } else {
        toast.error(response.message || "Gagal hapus komponen");
      }
    } catch (error: any) {
      console.error("Error deleting:", error);
      toast.error(error.response?.data?.error || "Gagal hapus komponen");
    } finally {
      setIsUpdating(false);
    }
  };

  // Handle search
  const handleSearch = useCallback(() => {
    if (searchTimeoutRef.current) {
      clearTimeout(searchTimeoutRef.current);
      searchTimeoutRef.current = null;
    }

    setIsSearching(true);
    searchTimeoutRef.current = setTimeout(() => {
      if (itemid && itemid.trim() !== "") {
        fetchBOM(itemid, false);
      } else {
        fetchBOM(undefined, true);
      }
    }, 300);
  }, [itemid, fetchBOM]);

  const handleReset = useCallback(() => {
    setItemid("");
    fetchBOM(undefined, true);
  }, [fetchBOM]);

  const handleInputChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const val = e.target.value;
      setItemid(val);

      if (searchTimeoutRef.current) {
        clearTimeout(searchTimeoutRef.current);
        searchTimeoutRef.current = null;
      }
    },
    [],
  );

  // Fetch item suggestions from master
  const fetchItemSuggestions = async (keyword: string) => {
    if (keyword.length < 2) {
      setItemSuggestions([]);
      setShowItemSuggestions(false);
      return;
    }

    setSuggestLoading(true);
    try {
      const res = await axios.get(`/api/master`, {
        params: { keyword, limit: 15 },
      });

      if (res.data.success) {
        setItemSuggestions(res.data.data);
        setShowItemSuggestions(res.data.data.length > 0);
      } else {
        setItemSuggestions([]);
        setShowItemSuggestions(false);
      }
    } catch (error) {
      console.error("Error fetching suggestions:", error);
      setItemSuggestions([]);
      setShowItemSuggestions(false);
    } finally {
      setSuggestLoading(false);
    }
  };

  const checkItemExists = async (itemID: string) => {
    if (!itemID.trim()) return null;

    setIsCheckingItem(true);
    try {
      const res = await axios.post(`/api/master`, { ItemID: itemID });

      if (res.data.success && res.data.exists) {
        setSelectedItemData(res.data.data);
        return res.data.data;
      } else {
        setSelectedItemData(null);
        return null;
      }
    } catch (error) {
      console.error("Error checking item:", error);
      setSelectedItemData(null);
      return null;
    } finally {
      setIsCheckingItem(false);
    }
  };

  const selectSuggestion = (item: MasterItem) => {
    setNewItemID(item.ItemID);
    setSelectedItemData(item);
    if (item.SatuanKecil) {
      setNewSatuan(item.SatuanKecil);
    }
    setShowItemSuggestions(false);
    toast.success(`Item ditemukan: ${item.ItemName}`);
  };

  const handleNewItemIDChange = (value: string) => {
    setNewItemID(value);
    setSelectedItemData(null);

    if (suggestTimeoutRef.current) {
      clearTimeout(suggestTimeoutRef.current);
      suggestTimeoutRef.current = null;
    }

    if (value.length >= 2) {
      suggestTimeoutRef.current = setTimeout(() => {
        fetchItemSuggestions(value);
      }, 300);
    } else {
      setItemSuggestions([]);
      setShowItemSuggestions(false);
    }
  };

  const handleNewItemIDBlur = async () => {
    if (newItemID && newItemID.trim()) {
      const itemData = await checkItemExists(newItemID);
      if (!itemData) {
        toast.warning(
          `Item ID "${newItemID}" tidak ditemukan di database master`,
        );
      }
    }
  };

  const toggleAllAccordions = useCallback(() => {
    if (expandedAll.length > 0) {
      setExpandedAll([]);
      showNotification("Semua detail ditutup", "info");
    } else {
      const headers = Object.keys(paginatedGrouped);
      setExpandedAll(headers);
      showNotification(
        `Semua detail dibuka (${headers.length} produk)`,
        "info",
      );
    }
  }, [paginatedGrouped]);

  const goToPage = (page: number) => {
    setCurrentPage(page);
    setExpandedAll([]);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const nextPage = () => {
    if (currentPage < totalPages) {
      setCurrentPage(currentPage + 1);
      setExpandedAll([]);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  const prevPage = () => {
    if (currentPage > 1) {
      setCurrentPage(currentPage - 1);
      setExpandedAll([]);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  };

  // Prevent hydration mismatch - render hanya setelah mounted
  if (!mounted) {
    return (
      <div className="container mx-auto p-4 md:p-6 space-y-6">
        <Card>
          <CardContent className="pt-6">
            <div className="flex justify-center p-8">Loading...</div>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="container mx-auto p-4 md:p-6 space-y-6">
      {/* Progress Bar */}
      {loadProgress.visible && (
        <div className="fixed bottom-4 right-4 z-50 w-80">
          <Card>
            <CardContent className="pt-4">
              <div className="space-y-2">
                <div className="flex justify-between text-sm">
                  <span>
                    {isExporting
                      ? "Mengekspor data..."
                      : isPrinting
                        ? "Mempersiapkan cetak..."
                        : "Memproses..."}
                  </span>
                  <span>{loadProgress.current}%</span>
                </div>
                <div className="h-2 bg-gray-200 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-primary transition-all duration-300"
                    style={{ width: `${loadProgress.current}%` }}
                  />
                </div>
                <p className="text-xs text-muted-foreground">
                  Mohon tunggu, sedang memproses data...
                </p>
              </div>
            </CardContent>
          </Card>
        </div>
      )}

      {/* Notification */}
      {notification.show && (
        <div className="fixed top-4 right-4 z-50 animate-in slide-in-from-top-2 fade-in duration-300">
          <Alert
            className={`border ${notification.type === "success" ? "bg-green-50 border-green-500 text-green-800" : notification.type === "error" ? "bg-red-50 border-red-500 text-red-800" : "bg-blue-50 border-blue-500 text-blue-800"}`}
          >
            {notification.type === "success" && (
              <CheckCircle2 className="h-4 w-4" />
            )}
            {notification.type === "error" && (
              <AlertTriangle className="h-4 w-4" />
            )}
            {notification.type === "info" && <Info className="h-4 w-4" />}
            <AlertDescription>{notification.message}</AlertDescription>
          </Alert>
        </div>
      )}

      {/* Header */}
      <div className="space-y-2">
        <div className="flex items-center gap-3">
          <ListTree className="h-8 w-8 text-primary" />
          <div>
            <h1 className="text-2xl md:text-3xl font-bold tracking-tight">
              Bill of Materials (BOM)
            </h1>
            <p className="text-muted-foreground">
              Struktur bahan lengkap untuk semua produk dengan spesifikasi,
              warna, dan bahan
            </p>
          </div>
        </div>
      </div>

      {/* Kontrol Pencarian */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Search className="h-5 w-5" /> Pencarian & Filter
          </CardTitle>
          <CardDescription>
            Cari atau filter BOM berdasarkan kriteria tertentu
          </CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="space-y-2">
              <label className="text-sm font-medium">Tipe Pencarian</label>
              <div className="flex gap-2">
                <Button
                  variant={searchType === "itemid" ? "default" : "outline"}
                  size="sm"
                  onClick={() => {
                    setSearchType("itemid");
                    setItemid("");
                    setData([]);
                    setHasSearched(false);
                    setSelectedProducts(new Set());
                  }}
                  className="flex-1"
                >
                  <Hash className="h-4 w-4 mr-2" /> ID Item
                </Button>
                <Button
                  variant={searchType === "itemname" ? "default" : "outline"}
                  size="sm"
                  onClick={() => {
                    setSearchType("itemname");
                    setItemid("");
                    setData([]);
                    setHasSearched(false);
                    setSelectedProducts(new Set());
                  }}
                  className="flex-1"
                >
                  <FileText className="h-4 w-4 mr-2" /> Nama Item
                </Button>
              </div>
            </div>
            <div className="space-y-2">
              <label className="text-sm font-medium">Kata Kunci</label>
              <div className="relative">
                <Input
                  placeholder={`Cari ${searchType === "itemid" ? "ID Item" : "Nama Item"}...`}
                  value={itemid}
                  onChange={handleInputChange}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") handleSearch();
                  }}
                  className="pl-10 pr-10"
                />
                <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-muted-foreground" />
                {itemid && (
                  <button
                    onClick={() => {
                      setItemid("");
                    }}
                    className="absolute right-3 top-1/2 transform -translate-y-1/2"
                  >
                    <X className="h-4 w-4 text-muted-foreground hover:text-foreground" />
                  </button>
                )}
              </div>
            </div>
            <div className="flex items-end gap-2">
              <Button
                onClick={handleSearch}
                disabled={loading || isSearching}
                className="flex-1"
              >
                {loading || isSearching ? (
                  <RefreshCw className="h-4 w-4 mr-2 animate-spin" />
                ) : (
                  <Search className="h-4 w-4 mr-2" />
                )}
                {loading || isSearching ? "Memuat..." : "Cari"}
              </Button>
            </div>
          </div>

          <div className="flex flex-wrap gap-2 pt-2">
            <Button
              variant="outline"
              size="sm"
              onClick={handleReset}
              disabled={loading}
              className="flex items-center gap-2"
            >
              <RefreshCw className="h-4 w-4" /> Reset & Load Semua Data
            </Button>

            <Button
              variant="outline"
              size="sm"
              onClick={toggleAllAccordions}
              disabled={data.length === 0}
              className="flex items-center gap-2"
            >
              {expandedAll.length > 0 ? (
                <EyeOff className="h-4 w-4" />
              ) : (
                <Eye className="h-4 w-4" />
              )}
              {expandedAll.length > 0 ? "Tutup Semua" : "Buka Semua"}
            </Button>

            {data.length > 0 && mounted && (
              <>
                <div className="flex items-center gap-1">
                  <Checkbox
                    id="selectPage"
                    checked={
                      selectedCountInPage ===
                        Object.keys(paginatedGrouped).length &&
                      Object.keys(paginatedGrouped).length > 0
                    }
                    onCheckedChange={handleSelectAllInPage}
                    disabled={
                      loading || Object.keys(paginatedGrouped).length === 0
                    }
                  />
                  <Label
                    htmlFor="selectPage"
                    className="text-sm cursor-pointer"
                  >
                    Pilih Halaman ({selectedCountInPage}/
                    {Object.keys(paginatedGrouped).length})
                  </Label>
                </div>

                <div className="flex items-center gap-1">
                  <Checkbox
                    id="selectAll"
                    checked={
                      totalSelectedCount === totalHeaders && totalHeaders > 0
                    }
                    onCheckedChange={handleSelectAllProducts}
                    disabled={loading || totalHeaders === 0}
                  />
                  <Label htmlFor="selectAll" className="text-sm cursor-pointer">
                    Pilih Semua ({totalSelectedCount}/{totalHeaders})
                  </Label>
                </div>

                <Button
                  variant="default"
                  size="sm"
                  onClick={exportSelectedToExcel}
                  disabled={isExporting || selectedProducts.size === 0}
                  className="flex items-center gap-2 bg-green-600 hover:bg-green-700"
                >
                  {isExporting ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <Download className="h-4 w-4" />
                  )}
                  Export ({selectedProducts.size})
                </Button>

                <Button
                  variant="secondary"
                  size="sm"
                  onClick={printSelectedProducts}
                  disabled={isPrinting || selectedProducts.size === 0}
                  className="flex items-center gap-2"
                >
                  {isPrinting ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <Printer className="h-4 w-4" />
                  )}
                  Cetak ({selectedProducts.size})
                </Button>
              </>
            )}
          </div>
        </CardContent>
      </Card>

      {/* Statistik */}
      {hasSearched && !loading && data.length > 0 && mounted && (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <Card>
            <CardContent className="pt-6">
              <div className="flex flex-col space-y-2">
                <div className="flex items-center justify-between">
                  <p className="text-sm font-medium text-muted-foreground">
                    Total Produk
                  </p>
                  <Package className="h-4 w-4 text-muted-foreground" />
                </div>
                <div className="flex items-center justify-between">
                  <h3 className="text-2xl font-bold">{totalHeaders}</h3>
                  <Badge variant="secondary">Produk</Badge>
                </div>
              </div>
            </CardContent>
          </Card>
          <Card>
            <CardContent className="pt-6">
              <div className="flex flex-col space-y-2">
                <div className="flex items-center justify-between">
                  <p className="text-sm font-medium text-muted-foreground">
                    Total Komponen
                  </p>
                  <Layers className="h-4 w-4 text-muted-foreground" />
                </div>
                <div className="flex items-center justify-between">
                  <h3 className="text-2xl font-bold">{totalComponentsAll}</h3>
                  <Badge variant="secondary">Item</Badge>
                </div>
              </div>
            </CardContent>
          </Card>
          <Card>
            <CardContent className="pt-6">
              <div className="flex flex-col space-y-2">
                <div className="flex items-center justify-between">
                  <p className="text-sm font-medium text-muted-foreground">
                    Rata-rata Komponen
                  </p>
                  <Box className="h-4 w-4 text-muted-foreground" />
                </div>
                <div className="flex items-center justify-between">
                  <h3 className="text-2xl font-bold">
                    {totalHeaders > 0
                      ? (totalComponentsAll / totalHeaders).toFixed(1)
                      : 0}
                  </h3>
                  <Badge variant="outline">per Produk</Badge>
                </div>
              </div>
            </CardContent>
          </Card>
          <Card>
            <CardContent className="pt-6">
              <div className="flex flex-col space-y-2">
                <div className="flex items-center justify-between">
                  <p className="text-sm font-medium text-muted-foreground">
                    Dipilih
                  </p>
                  <CheckCircle2 className="h-4 w-4 text-green-600" />
                </div>
                <div className="flex items-center justify-between">
                  <h3 className="text-2xl font-bold text-green-600">
                    {totalSelectedCount}
                  </h3>
                  <Badge variant="outline">Produk</Badge>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>
      )}

      {/* Hasil BOM */}
      <Card>
        <CardHeader className="flex flex-row items-center justify-between">
          <div>
            <CardTitle>Struktur BOM Lengkap</CardTitle>
            <CardDescription>
              {!hasSearched
                ? "Klik tombol Cari untuk menampilkan data"
                : loading
                  ? "Sedang memuat data..."
                  : data.length === 0
                    ? "Tidak ada data ditemukan"
                    : `Menampilkan ${Object.keys(paginatedGrouped).length} dari ${totalHeaders} produk`}
            </CardDescription>
          </div>
          {isAllDataMode && (
            <Badge variant="outline" className="bg-blue-50 text-blue-700">
              Mode: Semua Data
            </Badge>
          )}
        </CardHeader>
        <CardContent>
          {loading ? (
            <div className="space-y-4">
              {[...Array(5)].map((_, i) => (
                <Card key={i} className="overflow-hidden">
                  <CardContent className="pt-6">
                    <Skeleton className="h-8 w-3/4 mb-4" />
                    <div className="space-y-2 ml-6">
                      <Skeleton className="h-4 w-full" />
                      <Skeleton className="h-4 w-5/6" />
                      <Skeleton className="h-4 w-4/6" />
                      <Skeleton className="h-4 w-3/6" />
                    </div>
                  </CardContent>
                </Card>
              ))}
            </div>
          ) : !hasSearched ? (
            <div className="text-center py-12">
              <div className="mx-auto w-16 h-16 rounded-full bg-muted flex items-center justify-center mb-4">
                <Search className="h-8 w-8 text-muted-foreground" />
              </div>
              <h3 className="text-lg font-semibold mb-2">
                Belum Ada Pencarian
              </h3>
              <p className="text-muted-foreground mb-6">
                Masukkan kata kunci dan klik tombol Cari untuk menampilkan data
                BOM
              </p>
              <Button
                onClick={handleReset}
                variant="outline"
                className="flex items-center gap-2"
              >
                <RefreshCw className="h-4 w-4" /> Load Semua Data
              </Button>
            </div>
          ) : data.length === 0 ? (
            <div className="text-center py-12">
              <div className="mx-auto w-16 h-16 rounded-full bg-muted flex items-center justify-center mb-4">
                <Package className="h-8 w-8 text-muted-foreground" />
              </div>
              <h3 className="text-lg font-semibold mb-2">
                Data BOM tidak ditemukan
              </h3>
              <p className="text-muted-foreground mb-6">
                {itemid
                  ? `Tidak ada hasil untuk "${itemid}"`
                  : "Belum ada data BOM yang tersedia"}
              </p>
              {itemid && (
                <Button
                  variant="outline"
                  onClick={handleReset}
                  className="flex items-center gap-2"
                >
                  <RefreshCw className="h-4 w-4" /> Tampilkan Semua Data
                </Button>
              )}
            </div>
          ) : (
            <>
              <Accordion
                type="multiple"
                value={expandedAll}
                onValueChange={setExpandedAll}
                className="space-y-4"
              >
                {Object.entries(paginatedGrouped).map(([hd, group]) => (
                  <AccordionItem
                    key={hd}
                    value={hd}
                    className="border rounded-lg overflow-hidden"
                  >
                    <AccordionTrigger className="px-6 py-4 hover:no-underline hover:bg-accent/50">
                      <div className="flex flex-col items-start text-left w-full pr-4">
                        <div className="flex items-center gap-2 flex-wrap">
                          <Checkbox
                            checked={selectedProducts.has(hd)}
                            onCheckedChange={(checked) =>
                              handleSelectProduct(hd, checked as boolean)
                            }
                            onClick={(e) => e.stopPropagation()}
                          />
                          <span className="font-semibold text-lg font-mono bg-primary/10 px-2 py-1 rounded">
                            {hd}
                          </span>
                          <Badge variant="secondary" className="text-xs">
                            {group.totalComponents} komponen
                          </Badge>
                          <span
                            className="inline-flex items-center gap-1 h-6 px-2 text-xs rounded-md border border-input bg-background hover:bg-accent hover:text-accent-foreground cursor-pointer"
                            onClick={(e) => {
                              e.stopPropagation();
                              handleAddClick(
                                hd,
                                group.headerName,
                                parseInt(
                                  group.children[0]?.TransID.toString() || "0",
                                ),
                              );
                            }}
                          >
                            <Plus className="h-3 w-3" />
                            Tambah
                          </span>
                        </div>
                        <span className="text-sm text-foreground text-left mt-1 max-w-md">
                          {group.headerName}
                        </span>
                        {group.headerName2 &&
                          group.headerName2 !== group.headerName && (
                            <span className="text-xs text-muted-foreground text-left mt-0.5 max-w-md">
                              {group.headerName2}
                            </span>
                          )}
                      </div>
                    </AccordionTrigger>
                    <AccordionContent className="px-6 pb-6">
                      <div className="space-y-4">
                        <div className="bg-muted/30 p-4 rounded-lg">
                          <h4 className="font-medium mb-4 flex items-center gap-2">
                            <Layers className="h-4 w-4" /> Daftar Komponen
                          </h4>
                          <div className="overflow-x-auto">
                            <Table>
                              <TableHeader>
                                <TableRow>
                                  <TableHead className="w-12 text-center">
                                    No
                                  </TableHead>
                                  <TableHead className="min-w-[150px]">
                                    ID Komponen
                                  </TableHead>
                                  <TableHead className="min-w-[200px]">
                                    Nama Komponen
                                  </TableHead>
                                  <TableHead className="min-w-[200px]">
                                    Nama China
                                  </TableHead>
                                  <TableHead className="min-w-[150px]">
                                    <div className="flex items-center gap-1">
                                      <Ruler className="h-3 w-3" /> Spec
                                    </div>
                                  </TableHead>
                                  <TableHead className="min-w-[120px]">
                                    <div className="flex items-center gap-1">
                                      <Palette className="h-3 w-3" /> Warna
                                    </div>
                                  </TableHead>
                                  <TableHead className="min-w-[120px]">
                                    <div className="flex items-center gap-1">
                                      <Box className="h-3 w-3" /> Bahan
                                    </div>
                                  </TableHead>
                                  <TableHead className="w-24 text-right">
                                    Jumlah
                                  </TableHead>
                                  <TableHead className="w-24 text-center">
                                    Satuan
                                  </TableHead>
                                  <TableHead className="w-24 text-center">
                                    Aksi
                                  </TableHead>
                                </TableRow>
                              </TableHeader>
                              <TableBody>
                                {group.children.length === 0 ? (
                                  <TableRow>
                                    <td
                                      colSpan={10}
                                      className="px-4 py-8 text-center text-gray-500"
                                    >
                                      Belum ada komponen. Klik "Tambah" untuk
                                      menambahkan.
                                    </td>
                                  </TableRow>
                                ) : (
                                  group.children.map((row, index: number) => (
                                    <TableRow
                                      key={`${row.itemidHD}-${row.ItemID}`}
                                      className="hover:bg-muted/50"
                                    >
                                      <TableCell className="text-center font-medium">
                                        {index + 1}
                                      </TableCell>
                                      <TableCell>
                                        <div
                                          className="font-mono text-sm bg-secondary/20 px-2 py-1 rounded inline-block cursor-pointer hover:bg-secondary/40 transition-colors"
                                          onClick={() => {
                                            navigator.clipboard.writeText(
                                              row.ItemID,
                                            );
                                            showNotification(
                                              `Berhasil copy: ${row.ItemID}`,
                                              "success",
                                            );
                                          }}
                                        >
                                          {row.ItemID}
                                        </div>
                                      </TableCell>
                                      <TableCell>
                                        <div className="flex items-center gap-2">
                                          <ChevronRight className="h-3 w-3 text-muted-foreground flex-shrink-0" />
                                          <span title={row.ItemName}>
                                            {row.ItemName}
                                          </span>
                                        </div>
                                      </TableCell>
                                      <TableCell>
                                        <div className="flex items-center gap-2">
                                          <ChevronRight className="h-3 w-3 text-muted-foreground flex-shrink-0" />
                                          <span className="text-muted-foreground">
                                            {row.ItemName2 || "-"}
                                          </span>
                                        </div>
                                      </TableCell>
                                      <TableCell>
                                        <div
                                          className="text-sm truncate max-w-[150px]"
                                          title={row.Spec || "-"}
                                        >
                                          {row.Spec || "-"}
                                        </div>
                                      </TableCell>
                                      <TableCell>
                                        <div className="flex items-center gap-1">
                                          <div
                                            className="w-4 h-4 rounded border flex-shrink-0"
                                            style={{
                                              backgroundColor:
                                                row.warna || "#fff",
                                            }}
                                            title={row.warna || "-"}
                                          />
                                          <span
                                            className="text-sm truncate max-w-[80px]"
                                            title={row.warna || "-"}
                                          >
                                            {row.warna || "-"}
                                          </span>
                                        </div>
                                      </TableCell>
                                      <TableCell>
                                        <div
                                          className="text-sm truncate max-w-[120px]"
                                          title={row.bahan || "-"}
                                        >
                                          {row.bahan || "-"}
                                        </div>
                                      </TableCell>
                                      <TableCell className="text-right">
                                        <Badge
                                          variant="default"
                                          className="font-mono"
                                        >
                                          {row.BahanQty.toLocaleString()}
                                        </Badge>
                                      </TableCell>
                                      <TableCell className="text-center">
                                        <Badge
                                          variant="outline"
                                          className="text-xs"
                                        >
                                          {row.BahanPackSatuan || "-"}
                                        </Badge>
                                      </TableCell>
                                      <TableCell className="text-center">
                                        <div className="flex items-center justify-center gap-2">
                                          <Button
                                            variant="ghost"
                                            size="sm"
                                            className="h-8 w-8 p-0 text-blue-600"
                                            onClick={() => handleEditClick(row)}
                                          >
                                            <Edit className="h-4 w-4" />
                                          </Button>
                                          <Button
                                            variant="ghost"
                                            size="sm"
                                            className="h-8 w-8 p-0 text-red-600"
                                            onClick={() =>
                                              handleDeleteClick(row)
                                            }
                                          >
                                            <Trash2 className="h-4 w-4" />
                                          </Button>
                                        </div>
                                      </TableCell>
                                    </TableRow>
                                  ))
                                )}
                              </TableBody>
                            </Table>
                          </div>
                          <div className="mt-4 pt-4 border-t flex justify-between items-center text-sm text-muted-foreground">
                            <div>
                              Total:{" "}
                              <span className="font-semibold">
                                {group.totalComponents}
                              </span>{" "}
                              komponen
                            </div>
                            <div className="flex gap-2">
                              <Button
                                variant="ghost"
                                size="sm"
                                onClick={() => {
                                  navigator.clipboard.writeText(hd);
                                  showNotification(
                                    `Berhasil copy: ${hd}`,
                                    "success",
                                  );
                                }}
                              >
                                <Copy className="h-3 w-3" /> Copy ID
                              </Button>
                              <Button
                                variant="outline"
                                size="sm"
                                onClick={() =>
                                  handleAddClick(
                                    hd,
                                    group.headerName,
                                    parseInt(
                                      group.children[0]?.TransID.toString() ||
                                        "0",
                                    ),
                                  )
                                }
                              >
                                <Plus className="h-3 w-3 mr-1" /> Tambah
                                Komponen
                              </Button>
                            </div>
                          </div>
                        </div>
                      </div>
                    </AccordionContent>
                  </AccordionItem>
                ))}
              </Accordion>

              {!isAllDataMode && totalPages > 1 && mounted && (
                <div className="flex justify-between items-center mt-6 pt-4 border-t">
                  <div className="text-sm text-muted-foreground">
                    Halaman {currentPage} dari {totalPages}
                  </div>
                  <div className="flex gap-2">
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={prevPage}
                      disabled={currentPage === 1}
                    >
                      Sebelumnya
                    </Button>
                    <div className="flex gap-1">
                      {Array.from(
                        { length: Math.min(5, totalPages) },
                        (_, i) => {
                          let pageNum;
                          if (totalPages <= 5) pageNum = i + 1;
                          else if (currentPage <= 3) pageNum = i + 1;
                          else if (currentPage >= totalPages - 2)
                            pageNum = totalPages - 4 + i;
                          else pageNum = currentPage - 2 + i;
                          return (
                            <Button
                              key={pageNum}
                              variant={
                                currentPage === pageNum ? "default" : "outline"
                              }
                              size="sm"
                              className="w-9"
                              onClick={() => goToPage(pageNum)}
                            >
                              {pageNum}
                            </Button>
                          );
                        },
                      )}
                    </div>
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={nextPage}
                      disabled={currentPage === totalPages}
                    >
                      Selanjutnya
                    </Button>
                  </div>
                </div>
              )}

              {isAllDataMode && mounted && (
                <div className="text-center text-sm text-muted-foreground mt-4 pt-4 border-t">
                  Menampilkan semua {totalHeaders} produk ({totalComponentsAll}{" "}
                  komponen)
                </div>
              )}
            </>
          )}
        </CardContent>
      </Card>

      {/* Edit Dialog */}
      <Dialog open={editDialogOpen} onOpenChange={setEditDialogOpen}>
        <DialogContent className="sm:max-w-lg">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <Edit className="h-5 w-5" /> Edit Komponen BOM
            </DialogTitle>
            <DialogDescription>
              Ubah ID komponen, jumlah, atau satuan
            </DialogDescription>
          </DialogHeader>
          {editingComponent && (
            <div className="space-y-4 py-4">
              <div className="space-y-2">
                <Label>Produk</Label>
                <div className="p-2 bg-gray-100 dark:bg-gray-800 rounded text-sm">
                  {editingComponent.itemidHD} - {editingComponent.itemnamehd}
                </div>
              </div>
              <div className="space-y-2">
                <Label htmlFor="editItemID">ID Komponen *</Label>
                <Input
                  id="editItemID"
                  placeholder="Masukkan ID Komponen"
                  value={editItemID}
                  onChange={(e) => setEditItemID(e.target.value)}
                  className="font-mono"
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="editQuantity">Jumlah (Qty) *</Label>
                <Input
                  id="editQuantity"
                  type="number"
                  step="0.01"
                  value={editQuantity}
                  onChange={(e) =>
                    setEditQuantity(parseFloat(e.target.value) || 0)
                  }
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="editSatuan">Satuan</Label>
                <Select value={editSatuan} onValueChange={setEditSatuan}>
                  <SelectTrigger>
                    <SelectValue placeholder="Pilih satuan" />
                  </SelectTrigger>
                  <SelectContent>
                    {satuanOptions.map((option) => (
                      <SelectItem key={option.value} value={option.value}>
                        {option.label}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>
          )}
          <DialogFooter className="gap-2">
            <Button
              variant="outline"
              onClick={() => setEditDialogOpen(false)}
              disabled={isUpdating}
            >
              Batal
            </Button>
            <Button onClick={handleUpdateComponent} disabled={isUpdating}>
              {isUpdating ? (
                <RefreshCw className="h-4 w-4 mr-2 animate-spin" />
              ) : (
                <Save className="h-4 w-4 mr-2" />
              )}{" "}
              Simpan
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Add Component Dialog */}
      <Dialog open={addDialogOpen} onOpenChange={setAddDialogOpen}>
        <DialogContent className="sm:max-w-lg">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <Plus className="h-5 w-5" /> Tambah Komponen BOM
            </DialogTitle>
            <DialogDescription>
              Tambahkan komponen baru ke dalam BOM produk ini
            </DialogDescription>
          </DialogHeader>
          {selectedProduct && (
            <div className="space-y-4 py-4">
              <div className="space-y-2">
                <Label>Produk</Label>
                <div className="p-2 bg-gray-100 dark:bg-gray-800 rounded text-sm">
                  {selectedProduct.itemidHD} - {selectedProduct.itemnamehd}
                </div>
              </div>
              <div className="space-y-2">
                <Label htmlFor="newItemID">ID Komponen *</Label>
                <div className="relative">
                  <Input
                    id="newItemID"
                    placeholder="Masukkan ID Komponen"
                    value={newItemID}
                    onChange={(e) => handleNewItemIDChange(e.target.value)}
                    onBlur={handleNewItemIDBlur}
                    className={`font-mono ${selectedItemData ? "border-green-500" : newItemID ? "border-yellow-500" : ""}`}
                  />
                  {isCheckingItem && (
                    <div className="absolute right-3 top-1/2 transform -translate-y-1/2">
                      <Loader2 className="h-4 w-4 animate-spin text-gray-400" />
                    </div>
                  )}
                  {!isCheckingItem && newItemID && selectedItemData && (
                    <div className="absolute right-3 top-1/2 transform -translate-y-1/2">
                      <CheckCircle className="h-4 w-4 text-green-500" />
                    </div>
                  )}
                  {showItemSuggestions && itemSuggestions.length > 0 && (
                    <div className="absolute z-50 w-full mt-1 bg-white dark:bg-gray-800 border rounded-lg shadow-lg max-h-60 overflow-auto">
                      {itemSuggestions.map((item, idx) => (
                        <div
                          key={item.ItemID + idx}
                          className="px-4 py-2 hover:bg-gray-100 dark:hover:bg-gray-700 cursor-pointer border-b last:border-b-0"
                          onClick={() => selectSuggestion(item)}
                        >
                          <div className="font-mono text-sm font-medium">
                            {item.ItemID}
                          </div>
                          <div className="text-sm">{item.ItemName}</div>
                          {item.ItemName2 && (
                            <div className="text-xs text-gray-500">
                              {item.ItemName2}
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                  {suggestLoading && (
                    <div className="absolute z-50 w-full mt-1 bg-white dark:bg-gray-800 border rounded-lg shadow-lg p-4 text-center text-sm text-gray-500">
                      <Loader2 className="h-4 w-4 animate-spin inline mr-2" />{" "}
                      Mencari data...
                    </div>
                  )}
                </div>
                {selectedItemData && (
                  <div className="mt-2 p-2 bg-green-50 rounded text-sm">
                    <p className="text-green-700">
                      {selectedItemData.ItemName}
                    </p>
                    {selectedItemData.ItemName2 && (
                      <p className="text-xs text-green-600">
                        {selectedItemData.ItemName2}
                      </p>
                    )}
                    <div className="flex gap-4 mt-1 text-xs text-green-600">
                      {selectedItemData.Spec && (
                        <span>Spec: {selectedItemData.Spec}</span>
                      )}
                      {selectedItemData.warna && (
                        <span>Warna: {selectedItemData.warna}</span>
                      )}
                    </div>
                  </div>
                )}
              </div>
              <div className="space-y-2">
                <Label htmlFor="newQuantity">Jumlah (Qty) *</Label>
                <Input
                  id="newQuantity"
                  type="number"
                  step="0.01"
                  value={newQuantity}
                  onChange={(e) =>
                    setNewQuantity(parseFloat(e.target.value) || 0)
                  }
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="newSatuan">Satuan</Label>
                <Select value={newSatuan} onValueChange={setNewSatuan}>
                  <SelectTrigger>
                    <SelectValue placeholder="Pilih satuan" />
                  </SelectTrigger>
                  <SelectContent>
                    {satuanOptions.map((option) => (
                      <SelectItem key={option.value} value={option.value}>
                        {option.label}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>
          )}
          <DialogFooter className="gap-2">
            <Button
              variant="outline"
              onClick={() => setAddDialogOpen(false)}
              disabled={isAdding}
            >
              Batal
            </Button>
            <Button
              onClick={handleAddComponent}
              disabled={isAdding || !newItemID.trim() || newQuantity <= 0}
            >
              {isAdding ? (
                <RefreshCw className="h-4 w-4 mr-2 animate-spin" />
              ) : (
                <Plus className="h-4 w-4 mr-2" />
              )}{" "}
              Tambah Komponen
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <Dialog open={deleteDialogOpen} onOpenChange={setDeleteDialogOpen}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 text-red-600">
              <Trash2 className="h-5 w-5" /> Hapus Komponen
            </DialogTitle>
            <DialogDescription>
              Apakah Anda yakin ingin menghapus komponen ini? Tindakan ini tidak
              dapat dibatalkan.
            </DialogDescription>
          </DialogHeader>
          {deletingComponent && (
            <div className="space-y-4 py-4">
              <div className="p-4 bg-red-50 dark:bg-red-900/20 rounded-lg">
                <p className="text-sm font-medium">
                  Komponen yang akan dihapus:
                </p>
                <p className="text-sm mt-1">
                  <span className="font-mono">{deletingComponent.ItemID}</span>{" "}
                  - {deletingComponent.ItemName}
                </p>
                <p className="text-xs text-muted-foreground mt-1">
                  Dari produk: {deletingComponent.itemidHD} -{" "}
                  {deletingComponent.itemnamehd}
                </p>
              </div>
            </div>
          )}
          <DialogFooter className="gap-2">
            <Button
              variant="outline"
              onClick={() => setDeleteDialogOpen(false)}
              disabled={isUpdating}
            >
              Batal
            </Button>
            <Button
              variant="destructive"
              onClick={handleDeleteComponent}
              disabled={isUpdating}
            >
              {isUpdating ? (
                <RefreshCw className="h-4 w-4 mr-2 animate-spin" />
              ) : (
                <Trash2 className="h-4 w-4 mr-2" />
              )}{" "}
              Hapus
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

/* eslint-disable @typescript-eslint/no-explicit-any */
"use client";

import { useState, useEffect, useMemo } from "react";
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
  Package,
  Layers,
  Hash,
  FileText,
  Download,
  Box,
  ListTree,
  Eye,
  EyeOff,
  CheckCircle2,
  AlertTriangle,
  Printer,
  Palette,
  Ruler,
  ChevronLeft,
  ChevronRight,
  Edit,
  Trash2,
  Plus,
  Save,
  X,
  Loader2,
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
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogFooter,
  DialogDescription,
} from "@/components/ui/dialog";
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
  Spec?: string;
  warna?: string;
  bahan?: string;
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
  const [search, setSearch] = useState("");
  const [data, setData] = useState<BomData[]>([]);
  const [loading, setLoading] = useState(false);
  const [expandedAll, setExpandedAll] = useState<string[]>([]);
  const [selectedProducts, setSelectedProducts] = useState<Set<string>>(new Set());
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;
  
  // Edit state
  const [editDialogOpen, setEditDialogOpen] = useState(false);
  const [editingComponent, setEditingComponent] = useState<BomData | null>(null);
  const [editItemID, setEditItemID] = useState("");
  const [editQuantity, setEditQuantity] = useState(0);
  const [editSatuan, setEditSatuan] = useState("");
  const [isUpdating, setIsUpdating] = useState(false);
  
  // Delete state
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
  const [deletingComponent, setDeletingComponent] = useState<BomData | null>(null);

  useEffect(() => {
    setMounted(true);
    loadData();
  }, []);

  const loadData = async () => {
    const toastId = toast.loading("Mengambil data BOM...");
    setLoading(true);
    try {
      const res = await axios.get("/api/bom", { params: { fetchAll: true } });
      const bomData = res.data?.data || (Array.isArray(res.data) ? res.data : []);
      setData(bomData);
      toast.success(`Berhasil memuat ${bomData.length} komponen`, { id: toastId });
    } catch (err: any) {
      toast.error(err.message || "Gagal memuat data", { id: toastId });
    } finally {
      setLoading(false);
    }
  };

  const searchData = async () => {
    if (!search.trim()) {
      loadData();
      return;
    }
    const toastId = toast.loading("Mencari data...");
    setLoading(true);
    try {
      const res = await axios.get(`/api/bom?itemid=${search}`);
      const bomData = res.data?.data || (Array.isArray(res.data) ? res.data : []);
      setData(bomData);
      setCurrentPage(1);
      if (bomData.length === 0) {
        toast.info(`Tidak ditemukan hasil untuk "${search}"`, { id: toastId });
      } else {
        toast.success(`Ditemukan ${bomData.length} komponen`, { id: toastId });
      }
    } catch (err: any) {
      toast.error(err.message || "Gagal mencari", { id: toastId });
    } finally {
      setLoading(false);
    }
  };

  // Update komponen
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

    const toastId = toast.loading("Mengupdate komponen...");
    setIsUpdating(true);
    try {
      const response = await axios.put("/api/bom", {
        TransID: editingComponent.TransID.toString(),
        ItemID: editingComponent.ItemID,
        newItemID: editItemID,
        BahanQty: editQuantity,
        BahanPackSatuan: editSatuan,
        action: "update",
      });

      if (response.data.success) {
        toast.success("Komponen berhasil diupdate", { id: toastId });
        setEditDialogOpen(false);
        setEditingComponent(null);
        await loadData();
      } else {
        toast.error(response.data.message || "Gagal update komponen", { id: toastId });
      }
    } catch (error: any) {
      console.error("Error updating:", error);
      toast.error(error.response?.data?.error || "Gagal update komponen", { id: toastId });
    } finally {
      setIsUpdating(false);
    }
  };

  // Delete komponen
  const handleDeleteComponent = async () => {
    if (!deletingComponent) return;

    const toastId = toast.loading("Menghapus komponen...");
    setIsUpdating(true);
    try {
      const response = await axios.put("/api/bom", {
        TransID: deletingComponent.TransID.toString(),
        ItemID: deletingComponent.ItemID,
        action: "delete",
      });

      if (response.data.success) {
        toast.success("Komponen berhasil dihapus", { id: toastId });
        setDeleteDialogOpen(false);
        setDeletingComponent(null);
        await loadData();
      } else {
        toast.error(response.data.message || "Gagal hapus komponen", { id: toastId });
      }
    } catch (error: any) {
      console.error("Error deleting:", error);
      toast.error(error.response?.data?.error || "Gagal hapus komponen", { id: toastId });
    } finally {
      setIsUpdating(false);
    }
  };

  const grouped = useMemo(() => {
    const result: Record<string, any> = {};
    for (const row of data) {
      if (!row.itemidHD) continue;
      if (!result[row.itemidHD]) {
        result[row.itemidHD] = {
          name: row.itemnamehd,
          name2: row.itemnamehd2,
          components: []
        };
      }
      const exists = result[row.itemidHD].components.some((c: any) => c.ItemID === row.ItemID);
      if (!exists) {
        result[row.itemidHD].components.push(row);
      }
    }
    return result;
  }, [data]);

  const productIds = Object.keys(grouped);
  const totalProducts = productIds.length;
  const totalPages = Math.ceil(totalProducts / itemsPerPage);
  const paginatedIds = productIds.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);
  const selectedCount = selectedProducts.size;

  const handleSelectProduct = (id: string, checked: boolean) => {
    const newSet = new Set(selectedProducts);
    checked ? newSet.add(id) : newSet.delete(id);
    setSelectedProducts(newSet);
  };

  const handleSelectAll = () => {
    if (selectedCount === totalProducts) {
      setSelectedProducts(new Set());
    } else {
      setSelectedProducts(new Set(productIds));
    }
  };

  const toggleExpandAll = () => {
    if (expandedAll.length === paginatedIds.length) {
      setExpandedAll([]);
    } else {
      setExpandedAll(paginatedIds);
    }
  };

  const exportExcel = () => {
    if (selectedProducts.size === 0) {
      toast.error("Pilih minimal 1 produk");
      return;
    }
    
    const toastId = toast.loading("Mengexport data ke Excel...");
    
    try {
      const today = new Date();
      const fileName = `BOM_Report_${today.getFullYear()}_${(today.getMonth()+1).toString().padStart(2,'0')}_${today.getDate().toString().padStart(2,'0')}.xlsx`;
      
      const excelData: any[] = [];
      
      excelData.push(['BILL OF MATERIALS (BOM) REPORT']);
      excelData.push([`Tanggal: ${today.toLocaleDateString('id-ID')} | Waktu: ${today.toLocaleTimeString('id-ID')}`]);
      excelData.push([`Jumlah Produk: ${selectedProducts.size} | Total Komponen: ${Array.from(selectedProducts).reduce((sum, id) => sum + (grouped[id]?.components.length || 0), 0)}`]);
      excelData.push([]);
      
      let productNo = 1;
      for (const id of selectedProducts) {
        const product = grouped[id];
        if (product && product.components.length > 0) {
          excelData.push([`PRODUK #${productNo}`]);
          excelData.push([`Kode Produk: ${id}`]);
          excelData.push([`Nama Produk: ${product.name}`]);
          if (product.name2 && product.name2 !== product.name) {
            excelData.push([`Nama China: ${product.name2}`]);
          }
          excelData.push([`Jumlah Komponen: ${product.components.length}`]);
          excelData.push([]);
          
          excelData.push(['NO', 'KODE KOMPONEN', 'NAMA KOMPONEN', 'NAMA CHINA', 'SPESIFIKASI', 'WARNA', 'BAHAN', 'JUMLAH', 'SATUAN']);
          
          product.components.forEach((comp: BomData, idx: number) => {
            excelData.push([
              idx + 1,
              comp.ItemID,
              comp.ItemName || '-',
              comp.ItemName2 || '-',
              comp.Spec || '-',
              comp.warna || '-',
              comp.bahan || '-',
              comp.BahanQty,
              comp.BahanPackSatuan || '-'
            ]);
          });
          
          excelData.push([]);
          excelData.push([]);
          productNo++;
        }
      }
      
      excelData.push([`Dicetak dari Sistem BOM Management`]);
      excelData.push([`Total: ${selectedProducts.size} produk, ${Array.from(selectedProducts).reduce((sum, id) => sum + (grouped[id]?.components.length || 0), 0)} komponen`]);
      
      const ws = XLSX.utils.aoa_to_sheet(excelData);
      ws['!cols'] = [
        { wch: 6 }, { wch: 20 }, { wch: 35 }, { wch: 30 },
        { wch: 25 }, { wch: 15 }, { wch: 20 }, { wch: 12 }, { wch: 10 }
      ];
      
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "BOM Report");
      XLSX.writeFile(wb, fileName);
      
      toast.success(`Export ${selectedProducts.size} produk berhasil!`, { id: toastId });
    } catch (err) {
      console.error(err);
      toast.error("Gagal export data", { id: toastId });
    }
  };

  const printData = () => {
    if (selectedProducts.size === 0) {
      toast.error("Pilih minimal 1 produk");
      return;
    }
    
    const win = window.open("", "_blank");
    if (!win) return;
    
    let html = `
      <html>
      <head>
        <title>BOM Report</title>
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; font-size: 12px; }
          h1 { color: #2196F3; text-align: center; }
          h2 { color: #2196F3; margin-top: 20px; background: #e3f2fd; padding: 10px; }
          table { width: 100%; border-collapse: collapse; margin: 10px 0; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
          th { background: #2196F3; color: white; }
          .header { text-align: center; margin-bottom: 20px; }
          .footer { text-align: center; margin-top: 30px; font-size: 10px; color: #999; }
          .info { background: #f5f5f5; padding: 10px; margin: 10px 0; display: flex; justify-content: space-between; }
        </style>
      </head>
      <body>
      <div class="header">
        <h1>📋 BILL OF MATERIALS REPORT</h1>
        <p>Tanggal Cetak: ${new Date().toLocaleDateString('id-ID')} ${new Date().toLocaleTimeString('id-ID')}</p>
      </div>
      <div class="info">
        <span>📦 Produk: ${selectedProducts.size}</span>
        <span>🔧 Komponen: ${Array.from(selectedProducts).reduce((sum, id) => sum + (grouped[id]?.components.length || 0), 0)}</span>
      </div>
    `;
    
    for (const id of selectedProducts) {
      const product = grouped[id];
      if (product) {
        html += `<h2>📦 ${id} - ${product.name}</h2>`;
        if (product.name2 && product.name2 !== product.name) {
          html += `<p style="color:#666; font-size:11px;">${product.name2}</p>`;
        }
        html += `<p><strong>Total Komponen:</strong> ${product.components.length}</p>`;
        html += `<table><thead><tr><th>No</th><th>Kode</th><th>Nama Komponen</th><th>Spec</th><th>Warna</th><th>Bahan</th><th>Jumlah</th><th>Satuan</th></tr></thead><tbody>`;
        
        product.components.forEach((comp: BomData, idx: number) => {
          html += `<tr>
            <td>${idx + 1}</td>
            <td><strong>${comp.ItemID}</strong></td>
            <td>${comp.ItemName || '-'}</td>
            <td>${comp.Spec || '-'}</td>
            <td>${comp.warna || '-'}</td>
            <td>${comp.bahan || '-'}</td>
            <td style="text-align:right">${comp.BahanQty}</td>
            <td>${comp.BahanPackSatuan || '-'}</td>
          </tr>`;
        });
        html += `</tbody></table><br/><br/>`;
      }
    }
    
    html += `<div class="footer">Dicetak dari Sistem BOM Management</div></body></html>`;
    win.document.write(html);
    win.document.close();
    win.print();
  };

  if (!mounted) {
    return <div className="p-8 text-center">Loading...</div>;
  }

  return (
    <div className="container mx-auto p-4 space-y-4">
      {/* Header */}
      <div className="flex items-center gap-3">
        <ListTree className="h-8 w-8 text-primary" />
        <h1 className="text-2xl font-bold">Bill of Materials (BOM)</h1>
      </div>

      {/* Search */}
      <Card>
        <CardContent className="pt-6">
          <div className="flex gap-2">
            <Input
              placeholder="Cari ID Produk (contoh: LC-01B018BK)"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              onKeyPress={(e) => e.key === "Enter" && searchData()}
              className="flex-1 font-mono"
            />
            <Button onClick={searchData} disabled={loading}>
              {loading ? <Loader2 className="h-4 w-4 animate-spin" /> : <Search className="h-4 w-4" />}
              Cari
            </Button>
            <Button variant="outline" onClick={loadData} disabled={loading}>
              <RefreshCw className="h-4 w-4" />
              Reset
            </Button>
          </div>
        </CardContent>
      </Card>

      {/* Stats */}
      {data.length > 0 && (
        <div className="flex flex-wrap gap-4 items-center justify-between">
          <div className="flex gap-4">
            <Card className="w-32">
              <CardContent className="pt-4 text-center">
                <p className="text-xs text-muted-foreground">Produk</p>
                <p className="text-xl font-bold">{totalProducts}</p>
              </CardContent>
            </Card>
            <Card className="w-32">
              <CardContent className="pt-4 text-center">
                <p className="text-xs text-muted-foreground">Komponen</p>
                <p className="text-xl font-bold">{data.length}</p>
              </CardContent>
            </Card>
            <Card className="w-32">
              <CardContent className="pt-4 text-center">
                <p className="text-xs text-muted-foreground">Dipilih</p>
                <p className="text-xl font-bold text-green-600">{selectedCount}</p>
              </CardContent>
            </Card>
          </div>
          <div className="flex gap-2">
            <Button size="sm" variant="outline" onClick={toggleExpandAll}>
              {expandedAll.length > 0 ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
              {expandedAll.length > 0 ? "Tutup" : "Buka"} Semua
            </Button>
            <Button size="sm" variant="outline" onClick={handleSelectAll}>
              <Checkbox checked={selectedCount === totalProducts && totalProducts > 0} className="mr-2" />
              Pilih Semua ({selectedCount}/{totalProducts})
            </Button>
            <Button size="sm" onClick={exportExcel} disabled={selectedCount === 0} className="bg-green-600 hover:bg-green-700">
              <Download className="h-4 w-4 mr-1" /> Export Excel
            </Button>
            <Button size="sm" variant="secondary" onClick={printData} disabled={selectedCount === 0}>
              <Printer className="h-4 w-4 mr-1" /> Cetak
            </Button>
          </div>
        </div>
      )}

      {/* BOM Content */}
      <Card>
        <CardHeader>
          <CardTitle>Struktur BOM</CardTitle>
          <CardDescription>
            {loading ? "Memuat..." : data.length === 0 ? "Tidak ada data" : `Halaman ${currentPage} dari ${totalPages}`}
          </CardDescription>
        </CardHeader>
        <CardContent>
          {loading ? (
            <div className="space-y-3">{[...Array(3)].map((_, i) => <Skeleton key={i} className="h-24 w-full" />)}</div>
          ) : data.length === 0 ? (
            <div className="text-center py-12">
              <Package className="h-12 w-12 mx-auto text-muted-foreground mb-3" />
              <p>Tidak ada data</p>
              <Button variant="link" onClick={loadData}>Muat Data</Button>
            </div>
          ) : (
            <>
              <Accordion type="multiple" value={expandedAll} onValueChange={setExpandedAll} className="space-y-3">
                {paginatedIds.map((id) => {
                  const product = grouped[id];
                  if (!product) return null;
                  return (
                    <AccordionItem key={id} value={id} className="border rounded-lg">
                      <AccordionTrigger className="px-4 py-3">
                        <div className="flex items-center gap-3 text-left w-full">
                          <Checkbox
                            checked={selectedProducts.has(id)}
                            onCheckedChange={(c) => handleSelectProduct(id, c as boolean)}
                            onClick={(e) => e.stopPropagation()}
                          />
                          <div className="flex-1">
                            <div className="flex items-center gap-2 flex-wrap">
                              <span className="font-mono font-semibold bg-primary/10 px-2 py-0.5 rounded text-sm">{id}</span>
                              <Badge variant="secondary" className="text-xs">{product.components.length} komponen</Badge>
                            </div>
                            <p className="text-sm mt-1">{product.name}</p>
                            {product.name2 && product.name2 !== product.name && (
                              <p className="text-xs text-muted-foreground">{product.name2}</p>
                            )}
                          </div>
                        </div>
                      </AccordionTrigger>
                      <AccordionContent className="px-4 pb-4">
                        <div className="overflow-x-auto">
                          <Table>
                            <TableHeader>
                              <TableRow>
                                <TableHead className="w-10 text-center">No</TableHead>
                                <TableHead>Kode Komponen</TableHead>
                                <TableHead>Nama Komponen</TableHead>
                                <TableHead>Nama China</TableHead>
                                <TableHead>Spec</TableHead>
                                <TableHead>Warna</TableHead>
                                <TableHead>Bahan</TableHead>
                                <TableHead className="text-right">Jumlah</TableHead>
                                <TableHead>Satuan</TableHead>
                                <TableHead className="text-center">Aksi</TableHead>
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {product.components.map((comp: BomData, idx: number) => (
                                <TableRow key={comp.ItemID}>
                                  <TableCell className="text-center">{idx + 1}</TableCell>
                                  <TableCell className="font-mono text-sm">{comp.ItemID}</TableCell>
                                  <TableCell>{comp.ItemName || "-"}</TableCell>
                                  <TableCell className="text-muted-foreground">{comp.ItemName2 || "-"}</TableCell>
                                  <TableCell>{comp.Spec || "-"}</TableCell>
                                  <TableCell>
                                    <div className="flex items-center gap-1">
                                      <div className="w-3 h-3 rounded border" style={{ background: comp.warna || "#fff" }} />
                                      <span>{comp.warna || "-"}</span>
                                    </div>
                                  </TableCell>
                                  <TableCell>{comp.bahan || "-"}</TableCell>
                                  <TableCell className="text-right font-mono">{comp.BahanQty.toLocaleString()}</TableCell>
                                  <TableCell>{comp.BahanPackSatuan || "-"}</TableCell>
                                  <TableCell className="text-center">
                                    <div className="flex items-center justify-center gap-1">
                                      <Button
                                        variant="ghost"
                                        size="sm"
                                        className="h-7 w-7 p-0 text-blue-600"
                                        onClick={() => {
                                          setEditingComponent(comp);
                                          setEditItemID(comp.ItemID);
                                          setEditQuantity(comp.BahanQty);
                                          setEditSatuan(comp.BahanPackSatuan || "");
                                          setEditDialogOpen(true);
                                        }}
                                      >
                                        <Edit className="h-3.5 w-3.5" />
                                      </Button>
                                      <Button
                                        variant="ghost"
                                        size="sm"
                                        className="h-7 w-7 p-0 text-red-600"
                                        onClick={() => {
                                          setDeletingComponent(comp);
                                          setDeleteDialogOpen(true);
                                        }}
                                      >
                                        <Trash2 className="h-3.5 w-3.5" />
                                      </Button>
                                    </div>
                                  </TableCell>
                                </TableRow>
                              ))}
                            </TableBody>
                          </Table>
                        </div>
                      </AccordionContent>
                    </AccordionItem>
                  );
                })}
              </Accordion>

              {/* Pagination */}
              {totalPages > 1 && (
                <div className="flex justify-center gap-2 mt-6 pt-4 border-t">
                  <Button variant="outline" size="sm" onClick={() => setCurrentPage(p => Math.max(1, p - 1))} disabled={currentPage === 1}>
                    <ChevronLeft className="h-4 w-4" />
                  </Button>
                  <span className="px-4 py-2 text-sm">
                    Halaman {currentPage} / {totalPages}
                  </span>
                  <Button variant="outline" size="sm" onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))} disabled={currentPage === totalPages}>
                    <ChevronRight className="h-4 w-4" />
                  </Button>
                </div>
              )}
            </>
          )}
        </CardContent>
      </Card>

      {/* Edit Dialog */}
      <Dialog open={editDialogOpen} onOpenChange={setEditDialogOpen}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <Edit className="h-5 w-5" /> Edit Komponen BOM
            </DialogTitle>
            <DialogDescription>
              Ubah data komponen BOM
            </DialogDescription>
          </DialogHeader>
          {editingComponent && (
            <div className="space-y-4 py-4">
              <div className="space-y-2">
                <Label>Produk</Label>
                <div className="p-2 bg-muted rounded text-sm font-mono">
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
                  onChange={(e) => setEditQuantity(parseFloat(e.target.value) || 0)}
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="editSatuan">Satuan</Label>
                <select
                  id="editSatuan"
                  value={editSatuan}
                  onChange={(e) => setEditSatuan(e.target.value)}
                  className="w-full p-2 border rounded-md"
                >
                  {satuanOptions.map((opt) => (
                    <option key={opt.value} value={opt.value}>{opt.label}</option>
                  ))}
                </select>
              </div>
            </div>
          )}
          <DialogFooter className="gap-2">
            <Button variant="outline" onClick={() => setEditDialogOpen(false)} disabled={isUpdating}>
              Batal
            </Button>
            <Button onClick={handleUpdateComponent} disabled={isUpdating}>
              {isUpdating ? <Loader2 className="h-4 w-4 mr-2 animate-spin" /> : <Save className="h-4 w-4 mr-2" />}
              Simpan
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Delete Dialog */}
      <Dialog open={deleteDialogOpen} onOpenChange={setDeleteDialogOpen}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2 text-red-600">
              <Trash2 className="h-5 w-5" /> Hapus Komponen
            </DialogTitle>
            <DialogDescription>
              Apakah Anda yakin ingin menghapus komponen ini? Tindakan ini tidak dapat dibatalkan.
            </DialogDescription>
          </DialogHeader>
          {deletingComponent && (
            <div className="p-4 bg-red-50 rounded-lg">
              <p className="font-mono text-sm">{deletingComponent.ItemID}</p>
              <p className="text-sm mt-1">{deletingComponent.ItemName}</p>
              <p className="text-xs text-muted-foreground mt-2">
                Dari produk: {deletingComponent.itemidHD} - {deletingComponent.itemnamehd}
              </p>
            </div>
          )}
          <DialogFooter className="gap-2">
            <Button variant="outline" onClick={() => setDeleteDialogOpen(false)} disabled={isUpdating}>
              Batal
            </Button>
            <Button variant="destructive" onClick={handleDeleteComponent} disabled={isUpdating}>
              {isUpdating ? <Loader2 className="h-4 w-4 mr-2 animate-spin" /> : <Trash2 className="h-4 w-4 mr-2" />}
              Hapus
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
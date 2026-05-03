"use client";
import React, { useEffect, useState, useMemo, useCallback } from "react";
import { useDispatch, useSelector } from "react-redux";
import { fetchSupplierData } from "@/lib/features/supplierSlice";
import { RootState, AppDispatch } from "@/lib/store";
import Loading from "@/app/loading";
import { DataTable } from "../data-table";
import { columns } from "./columns";
import { saveAs } from "file-saver";
import { Input } from "../ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "../ui/card";
import { toast } from "sonner";
import * as XLSX from "xlsx";
import { Button } from "../ui/button";
import { SupplierType } from "@/lib/types";
import {
  FileUp,
  RefreshCcw,
  Search,
  Download,
  Users,
  Building2,
  FileSpreadsheet,
  Filter,
  X
} from "lucide-react";
import { Badge } from "../ui/badge";
import { Progress } from "../ui/progress";

const useAppDispatch = () => useDispatch<AppDispatch>();

const DataMasterSupplierPage: React.FC = () => {
  const dispatch = useAppDispatch();
  const { data, loading, error } = useSelector(
    (state: RootState) => state.supplier
  );

  const [searchTerm, setSearchTerm] = useState<string>("");
  const [selectedRows, setSelectedRows] = useState<SupplierType[]>([]);
  const [exportProgress, setExportProgress] = useState<number>(0);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [debouncedSearchTerm, setDebouncedSearchTerm] = useState<string>("");

  // Debounce search untuk performa lebih baik
  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedSearchTerm(searchTerm);
    }, 300);
    return () => clearTimeout(timer);
  }, [searchTerm]);

  // Fetch data dengan toast promise
  const fetchData = useCallback(() => {
    const myPromise = dispatch(fetchSupplierData());
    toast.promise(myPromise, {
      loading: "Memuat data supplier...",
      success: "Data supplier berhasil dimuat!",
      error: "Gagal memuat data supplier",
    });
  }, [dispatch]);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  // Filter data berdasarkan search term (menggunakan debounced)
  const filteredData = useMemo(() => {
    if (!debouncedSearchTerm.trim()) return data;

    const term = debouncedSearchTerm.toLowerCase();
    return data.filter((item) =>
      Object.values(item).some((value) =>
        String(value).toLowerCase().includes(term)
      )
    );
  }, [data, debouncedSearchTerm]);

  // Handle export dengan progress
  const handleExport = useCallback(async () => {
    const rowsToExport = selectedRows.length > 0 ? selectedRows : filteredData;

    if (rowsToExport.length === 0) {
      toast.error("Tidak ada data yang dapat diekspor", {
        description: "Silakan pilih data atau pastikan data tersedia",
      });
      return;
    }

    setIsExporting(true);
    setExportProgress(0);

    try {
      // Simulasi progress untuk UX yang lebih baik
      for (let i = 0; i <= 100; i += 20) {
        await new Promise(resolve => setTimeout(resolve, 50));
        setExportProgress(i);
      }

      const ws = XLSX.utils.json_to_sheet(rowsToExport);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data Supplier");

      // Auto-size columns
      const maxWidth = rowsToExport.reduce((widths, row) => {
        Object.values(row).forEach((value, idx) => {
          const length = String(value).length;
          widths[idx] = Math.max(widths[idx] || 10, Math.min(length, 50));
        });
        return widths;
      }, [] as number[]);

      ws['!cols'] = maxWidth.map(width => ({ wch: width + 2 }));

      const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });

      const fileName = `Data_Supplier_${new Date().toISOString().split('T')[0]}.xlsx`;
      saveAs(blob, fileName);

      toast.success(`Berhasil mengekspor ${rowsToExport.length} data supplier`, {
        description: `File ${fileName} telah disimpan`,
      });
    } catch (error) {
      console.error("Export error:", error);
      toast.error("Gagal mengekspor data", {
        description: "Terjadi kesalahan saat mengekspor data",
      });
    } finally {
      setIsExporting(false);
      setExportProgress(0);
    }
  }, [selectedRows, filteredData]);

  const handleRefresh = useCallback(() => {
    setSearchTerm("");
    setSelectedRows([]);
    fetchData();
  }, [fetchData]);

  const handleSearchChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setSearchTerm(event.target.value);
  };

  const clearSearch = () => {
    setSearchTerm("");
    setDebouncedSearchTerm("");
  };

  // Statistics
  const stats = useMemo(() => ({
    total: data.length,
    filtered: filteredData.length,
    selected: selectedRows.length,
  }), [data.length, filteredData.length, selectedRows.length]);

  if (loading) return <Loading />;
  if (error) return (
    <div className="flex flex-col items-center justify-center min-h-screen p-4">
      <Card className="max-w-md w-full border-red-200 bg-red-50">
        <CardContent className="pt-6">
          <div className="text-center">
            <div className="text-red-600 text-lg font-semibold mb-2">Error</div>
            <p className="text-red-500">{error}</p>
            <Button onClick={handleRefresh} className="mt-4">
              <RefreshCcw className="h-4 w-4 mr-2" />
              Coba Lagi
            </Button>
          </div>
        </CardContent>
      </Card>
    </div>
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 dark:from-slate-950 dark:to-slate-900">
      <div className="container mx-auto p-4 md:p-6 lg:p-8">
        {/* Header Section */}
        <div className="mb-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
            <div>
              <h1 className="text-3xl md:text-4xl lg:text-5xl font-bold bg-gradient-to-r from-slate-900 to-slate-600 dark:from-slate-100 dark:to-slate-400 bg-clip-text text-transparent">
                Data Supplier
              </h1>
              <p className="text-muted-foreground mt-2">
                Kelola dan monitor data supplier perusahaan
              </p>
            </div>

            
          </div>
        </div>

        {/* Main Card */}
        <Card className="border-0 shadow-xl bg-white dark:bg-slate-900 overflow-hidden">
          <CardHeader className="border-b bg-gradient-to-r from-slate-50 to-white dark:from-slate-900 dark:to-slate-950">
            <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4">
              <CardTitle className="text-xl font-semibold flex items-center gap-2">
                <FileSpreadsheet className="h-5 w-5 text-primary" />
                Daftar Supplier
              </CardTitle>

              <div className="flex flex-col sm:flex-row gap-3 w-full lg:w-auto">
                {/* Search Bar */}
                <div className="relative flex-1 min-w-[200px]">
                  <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-muted-foreground" />
                  <Input
                    type="search"
                    onChange={handleSearchChange}
                    value={searchTerm}
                    placeholder="Cari supplier..."
                    className="pl-9 pr-8 h-10 bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700"
                  />
                  {searchTerm && (
                    <button
                      onClick={clearSearch}
                      className="absolute right-3 top-1/2 transform -translate-y-1/2"
                    >
                      <X className="h-4 w-4 text-muted-foreground hover:text-foreground" />
                    </button>
                  )}
                </div>

                {/* Action Buttons */}
                <div className="flex gap-2">
                  <Button
                    onClick={handleRefresh}
                    variant="outline"
                    size="default"
                    className="h-10 px-4"
                  >
                    <RefreshCcw className="h-4 w-4 mr-2" />
                    Refresh
                  </Button>
                  <Button
                    onClick={handleExport}
                    disabled={isExporting || filteredData.length === 0}
                    className="h-10 px-6 bg-gradient-to-r from-green-600 to-green-500 hover:from-green-700 hover:to-green-600 text-white shadow-md hover:shadow-lg transition-all duration-200"
                  >
                    {isExporting ? (
                      <div className="flex items-center gap-2">
                        <div className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                        <span>Exporting...</span>
                      </div>
                    ) : (
                      <>
                        <Download className="h-4 w-4 mr-2" />
                        Export Excel
                      </>
                    )}
                  </Button>
                </div>
              </div>
            </div>

            {/* Export Progress */}
            {isExporting && (
              <div className="mt-4">
                <Progress value={exportProgress} className="h-2" />
                <p className="text-xs text-muted-foreground mt-1 text-center">
                  Sedang mengekspor data...
                </p>
              </div>
            )}

            {/* Info Badges */}
            <div className="flex flex-wrap gap-2 mt-4">
              {searchTerm && (
                <Badge variant="secondary" className="gap-1">
                  <Search className="h-3 w-3" />
                  Pencarian: {searchTerm}
                  <button onClick={clearSearch} className="ml-1 hover:text-foreground">
                    <X className="h-3 w-3" />
                  </button>
                </Badge>
              )}
              {selectedRows.length > 0 && (
                <Badge variant="default" className="bg-purple-500 gap-1">
                  <Users className="h-3 w-3" />
                  {selectedRows.length} supplier dipilih
                </Badge>
              )}
              {filteredData.length !== data.length && (
                <Badge variant="outline" className="gap-1">
                  <Filter className="h-3 w-3" />
                  Menampilkan {filteredData.length} dari {data.length} supplier
                </Badge>
              )}
            </div>
          </CardHeader>

          <CardContent className="p-0">
            <DataTable
              columns={columns(setSelectedRows, filteredData)}
              data={filteredData}
            />
          </CardContent>
        </Card>

        {/* Footer */}
        <div className="mt-4 text-center text-xs text-muted-foreground">
          <p>© {new Date().getFullYear()} - Data Supplier Management System</p>
        </div>
      </div>
    </div>
  );
};

export default DataMasterSupplierPage;
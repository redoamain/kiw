import { Loader2 } from "lucide-react";

export default function Loading() {
  return (
    <div className="flex items-center justify-center min-h-screen flex-col">
      <Loader2 className="h-32 w-32 animate-spin text-primary" />
    </div>
  );
}
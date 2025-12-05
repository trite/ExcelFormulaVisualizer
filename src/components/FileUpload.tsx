import { useCallback } from "react";
import { Button, Typography, Paper } from "@mui/material";
import CloudUploadIcon from "@mui/icons-material/CloudUpload";

interface FileUploadProps {
  onFileSelect: (file: File) => void;
  isLoading?: boolean;
}

export function FileUpload({ onFileSelect, isLoading }: FileUploadProps) {
  const handleDrop = useCallback(
    (event: React.DragEvent<HTMLDivElement>) => {
      event.preventDefault();
      const file = event.dataTransfer.files[0];
      if (file && isExcelFile(file)) {
        onFileSelect(file);
      }
    },
    [onFileSelect]
  );

  const handleDragOver = useCallback(
    (event: React.DragEvent<HTMLDivElement>) => {
      event.preventDefault();
    },
    []
  );

  const handleFileInput = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (file && isExcelFile(file)) {
        onFileSelect(file);
      }
    },
    [onFileSelect]
  );

  const isExcelFile = (file: File): boolean => {
    const validExtensions = [".xlsx", ".xls", ".xlsm", ".xlsb"];
    const extension = file.name.toLowerCase().slice(file.name.lastIndexOf("."));
    return validExtensions.includes(extension);
  };

  return (
    <Paper
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      sx={{
        p: 4,
        border: "2px dashed",
        borderColor: "primary.main",
        borderRadius: 2,
        textAlign: "center",
        cursor: "pointer",
        transition: "all 0.2s ease",
        "&:hover": {
          borderColor: "primary.light",
          bgcolor: "action.hover",
        },
      }}
    >
      <CloudUploadIcon sx={{ fontSize: 64, color: "primary.main", mb: 2 }} />
      <Typography variant="h6" gutterBottom>
        Drag & drop an Excel file here
      </Typography>
      <Typography variant="body2" color="text.secondary" gutterBottom>
        or
      </Typography>
      <Button
        variant="contained"
        component="label"
        disabled={isLoading}
        sx={{ mt: 1 }}
      >
        {isLoading ? "Processing..." : "Browse Files"}
        <input
          type="file"
          hidden
          accept=".xlsx,.xls,.xlsm,.xlsb"
          onChange={handleFileInput}
        />
      </Button>
      <Typography
        variant="caption"
        display="block"
        sx={{ mt: 2 }}
        color="text.secondary"
      >
        Supported formats: .xlsx, .xls, .xlsm, .xlsb
      </Typography>
    </Paper>
  );
}

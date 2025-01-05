import { useState } from "react";
import { read, utils } from "xlsx";
import { Button } from "@/components/ui/button";
import { useToast } from "@/components/ui/use-toast";
import { Card } from "@/components/ui/card";
import { Upload, Send, FileSpreadsheet } from "lucide-react";

const Index = () => {
  const [fileData, setFileData] = useState<any>(null);
  const { toast } = useToast();

  const extractCellValues = (worksheet: any) => {
    const cellAddresses = [
      "A6", "F6", "H6", "I6", "CN6", "CO6", "CP6", "CR6", "CS6", "CT6",
      "DA6", "DD6", "DE6", "DI6", "DQ6", "EG6", "EM6", "EN6", "EO6",
      "EP6", "FB6", "FL6", "FV6", "FY6", "FZ6", "GE6", "GF6"
    ];

    return cellAddresses.reduce((acc: any, address) => {
      acc[address] = worksheet[address]?.v || "";
      return acc;
    }, {});
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    
    if (!file) {
      toast({
        title: "Erro",
        description: "Nenhum arquivo selecionado",
        variant: "destructive"
      });
      return;
    }

    try {
      const data = await file.arrayBuffer();
      const workbook = read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const values = extractCellValues(worksheet);
      
      setFileData(values);
      toast({
        title: "Sucesso",
        description: "Dados extraídos com sucesso!",
      });
    } catch (error) {
      toast({
        title: "Erro",
        description: "Erro ao ler o arquivo",
        variant: "destructive"
      });
    }
  };

  const sendToGoogleSheets = async () => {
    if (!fileData) {
      toast({
        title: "Erro",
        description: "Nenhum dado para enviar",
        variant: "destructive"
      });
      return;
    }

    try {
      const GOOGLE_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxQSY46agkwGpgjGTPxRArL_UTpUaM0NUu4sc9172ovXc9igY8LqR55NV3aD1qmnuY_/exec';
      
      const response = await fetch(GOOGLE_APPS_SCRIPT_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(fileData)
      });

      if (!response.ok) {
        throw new Error('Erro ao enviar dados');
      }

      toast({
        title: "Sucesso",
        description: "Dados enviados para o Google Sheets com sucesso!",
      });
    } catch (error) {
      toast({
        title: "Erro",
        description: "Erro ao enviar dados para o Google Sheets",
        variant: "destructive"
      });
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-blue-50 to-white p-8">
      <div className="max-w-4xl mx-auto space-y-8">
        <Card className="p-8">
          <h1 className="text-3xl font-bold text-blue-900 mb-6">
            Importador de Planilhas
          </h1>
          
          <div className="space-y-6">
            <div className="flex flex-col items-center justify-center border-2 border-dashed border-blue-200 rounded-lg p-8 bg-blue-50 hover:bg-blue-100 transition-colors">
              <FileSpreadsheet className="w-12 h-12 text-blue-500 mb-4" />
              <Button
                variant="outline"
                className="gap-2"
                onClick={() => document.getElementById('file-upload')?.click()}
              >
                <Upload className="w-4 h-4" />
                Selecionar Arquivo XLSX
              </Button>
              <input
                id="file-upload"
                type="file"
                accept=".xlsx"
                className="hidden"
                onChange={handleFileUpload}
              />
            </div>

            {fileData && (
              <div className="space-y-4">
                <h2 className="text-xl font-semibold text-blue-900">
                  Dados Extraídos
                </h2>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  {Object.entries(fileData).map(([cell, value]) => (
                    <div key={cell} className="p-3 bg-white rounded-lg shadow-sm">
                      <span className="font-medium text-blue-600">{cell}:</span>
                      <span className="ml-2 text-gray-700">{String(value)}</span>
                    </div>
                  ))}
                </div>
                <Button
                  onClick={sendToGoogleSheets}
                  className="w-full mt-4 gap-2"
                >
                  <Send className="w-4 h-4" />
                  Enviar para Google Sheets
                </Button>
              </div>
            )}
          </div>
        </Card>
      </div>
    </div>
  );
};

export default Index;
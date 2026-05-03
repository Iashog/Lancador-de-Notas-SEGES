import React, { useState, useMemo } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { FileSpreadsheet, ClipboardCopy, Settings2, Layers, GraduationCap, CheckCircle2 } from "lucide-react";
import * as XLSX from 'xlsx';
import { showSuccess, showError } from "@/utils/toast";
import { MadeWithDyad } from "@/components/made-with-dyad";

interface WorkbookData {
  [sheetName: string]: any[];
}

const AREAS = [
  { id: 'CH', name: 'Ciências Humanas' },
  { id: 'CN', name: 'Ciências da Natureza' },
  { id: 'LI', name: 'Linguagens' },
  { id: 'MT', name: 'Matemática' },
];

const CATEGORIES = [
  { id: 'all', name: 'Todas as Categorias' },
  { id: '0', name: 'Atividade Discursiva' },
  { id: '1', name: 'Prova Interdisciplinar' },
  { id: '2', name: 'Produção Escrita' },
];

const FIELDS = [
  { id: 'both', name: 'Avaliação e Recuperação' },
  { id: 'av', name: 'Apenas Avaliação (Av)' },
  { id: 'rec', name: 'Apenas Recuperação (Rec)' },
];

const Index = () => {
  const [workbookData, setWorkbookData] = useState<WorkbookData>({});
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [studentNameCol, setStudentNameCol] = useState<string>('');
  
  const [selectedTurma, setSelectedTurma] = useState<string>('');
  const [selectedArea, setSelectedArea] = useState<string>('');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [selectedField, setSelectedField] = useState<string>('both');

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bData = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(bData, { type: 'array' });
        
        const newWorkbookData: WorkbookData = {};
        workbook.SheetNames.forEach(name => {
          const worksheet = workbook.Sheets[name];
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" }) as any[][];
          
          // 1. Encontrar a linha das Áreas (CH, CN...)
          let areaRowIndex = rows.findIndex(row => 
            row.some(cell => ['CH', 'CN', 'LI', 'MT'].includes(String(cell).trim().toUpperCase()))
          );

          if (areaRowIndex === -1) return;

          // 2. Processar cabeçalhos com "Fill Forward" para células mescladas
          const rawAreas = rows[areaRowIndex].map(h => String(h || '').trim().toUpperCase());
          const processedHeaders: string[] = [];
          let lastArea = "";

          rawAreas.forEach((area, i) => {
            if (['CH', 'CN', 'LI', 'MT'].includes(area)) {
              lastArea = area;
            }
            // Se a coluna começa com R (Recuperação), ex: RCH
            if (area.startsWith('R') && ['RCH', 'RCN', 'RLI', 'RMT'].includes(area)) {
                processedHeaders.push(area);
            } else {
                processedHeaders.push(lastArea || `COL_${i}`);
            }
          });

          // 3. Mapear dados
          const dataRows = rows.slice(areaRowIndex + 1);
          const formattedData = dataRows.map(row => {
            const studentObj: any = { _raw: row };
            
            // Agrupar notas por área para suportar múltiplas categorias
            AREAS.forEach(area => {
                studentObj[area.id] = row.filter((_, i) => processedHeaders[i] === area.id);
                studentObj[`R${area.id}`] = row.filter((_, i) => processedHeaders[i] === `R${area.id}`);
            });

            // Detectar nome do aluno (primeira coluna com texto longo)
            studentObj._name = String(row.find(cell => String(cell).length > 5 && !/\d/.test(String(cell))) || "");
            
            return studentObj;
          }).filter(row => row._name.length > 3 && !row._name.includes("TOTAL") && !row._name.includes("MÉDIA"));

          newWorkbookData[name] = formattedData;
        });

        setWorkbookData(newWorkbookData);
        setSheetNames(workbook.SheetNames);
        
        if (workbook.SheetNames.length > 0) {
          setSelectedTurma(workbook.SheetNames[0]);
          showSuccess("Planilha processada com sucesso!");
        }
      } catch (err) {
        showError("Erro ao processar Excel.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const filteredData = useMemo(() => {
    return workbookData[selectedTurma] || [];
  }, [workbookData, selectedTurma]);

  const generateScript = () => {
    if (!selectedTurma) return showError("Selecione a Turma.");
    if (!selectedArea) return showError("Selecione a Área.");

    const scriptData = filteredData.map(row => {
      const areaNotes = row[selectedArea] || [];
      const recNotes = row[`R${selectedArea}`] || [];
      
      return {
        name: row._name.trim().toUpperCase(),
        notes: areaNotes,
        recs: recNotes
      };
    });

    const script = `
(function() {
  const studentsData = ${JSON.stringify(scriptData)};
  const category = "${selectedCategory}";
  const field = "${selectedField}";
  
  let count = 0;
  studentsData.forEach(student => {
    const row = Array.from(document.querySelectorAll('tr')).find(tr => 
      tr.innerText.toUpperCase().includes(student.name)
    );

    if (row) {
      const inputs = Array.from(row.querySelectorAll('input[type="text"]'));
      // No SEGES: [Cat0_Av, Cat0_Rec, Cat1_Av, Cat1_Rec, Cat2_Av, Cat2_Rec]
      const targetIndices = category === 'all' ? [0, 1, 2, 3, 4, 5] : [parseInt(category)*2, parseInt(category)*2 + 1];

      targetIndices.forEach(idx => {
        const catIdx = Math.floor(idx / 2);
        const isAv = idx % 2 === 0;
        
        if ((field === 'av' && !isAv) || (field === 'rec' && isAv)) return;

        const val = isAv ? student.notes[catIdx] : student.recs[catIdx];

        if (val !== undefined && val !== null && val !== "" && inputs[idx]) {
          inputs[idx].value = val.toString().replace('.', ',');
          ['input', 'change', 'blur'].forEach(t => inputs[idx].dispatchEvent(new Event(t, { bubbles: true })));
        }
      });
      count++;
      row.style.backgroundColor = '#f0fdf4';
      row.style.borderLeft = '4px solid #22c55e';
    }
  });
  alert('Sucesso! ' + count + ' alunos da turma "${selectedTurma}" atualizados.');
})();`;

    navigator.clipboard.writeText(script);
    showSuccess("Script copiado!");
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="max-w-6xl mx-auto space-y-6">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div className="lg:col-span-4 space-y-6">
            <Card className="border-none shadow-sm overflow-hidden">
              <CardHeader className="bg-indigo-600 text-white pb-8">
                <div className="flex items-center gap-3 mb-2">
                  <GraduationCap className="w-6 h-6" />
                  <CardTitle className="text-lg font-bold">Lançador SEGES</CardTitle>
                </div>
                <CardDescription className="text-indigo-100 text-xs">
                  Configure os dados da planilha para gerar o script de lançamento.
                </CardDescription>
              </CardHeader>
              
              <CardContent className="p-5 space-y-6 -mt-6">
                <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 space-y-3">
                  <Label className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">
                    <FileSpreadsheet className="w-3 h-3" /> 1. Carregar Planilha
                  </Label>
                  <div className="relative group">
                    <Input 
                      type="file" 
                      accept=".xlsx,.xls" 
                      onChange={handleFileUpload} 
                      className="cursor-pointer bg-slate-50 border-dashed border-2 border-slate-200 hover:border-indigo-400 transition-colors h-16 text-[10px] pt-6" 
                    />
                    <div className="absolute inset-0 flex items-center justify-center pointer-events-none text-slate-400 text-xs font-medium group-hover:text-indigo-600">
                      {sheetNames.length > 0 ? (
                        <span className="flex items-center gap-2 text-green-600"><CheckCircle2 className="w-4 h-4" /> Planilha Pronta</span>
                      ) : (
                        "Clique para selecionar o arquivo"
                      )}
                    </div>
                  </div>
                </div>

                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-[10px]">2</span>
                    Turma (Abas da Planilha)
                  </Label>
                  <Select value={selectedTurma} onValueChange={setSelectedTurma} disabled={sheetNames.length === 0}>
                    <SelectTrigger className="bg-slate-50 border-slate-100 h-11">
                      <SelectValue placeholder="Selecione a turma..." />
                    </SelectTrigger>
                    <SelectContent>
                      {sheetNames.map(name => <SelectItem key={name} value={name}>{name}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label className="text-xs font-bold text-slate-600 flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-[10px]">3</span>
                    Área de Conhecimento
                  </Label>
                  <div className="grid grid-cols-2 gap-2">
                    {AREAS.map(area => (
                      <button 
                        key={area.id} 
                        onClick={() => setSelectedArea(area.id)} 
                        className={`p-3 rounded-xl text-xs font-bold transition-all border-2 ${
                          selectedArea === area.id 
                          ? 'bg-indigo-600 border-indigo-600 text-white shadow-md' 
                          : 'bg-white border-slate-100 text-slate-500 hover:border-indigo-200'
                        }`}
                      >
                        {area.id}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="grid grid-cols-1 gap-4 pt-2">
                  <div className="space-y-2">
                    <Label className="text-[10px] font-bold text-slate-400 uppercase">Categoria</Label>
                    <Select value={selectedCategory} onValueChange={setSelectedCategory}>
                      <SelectTrigger className="bg-slate-50 border-slate-100 h-10 text-xs"><SelectValue /></SelectTrigger>
                      <SelectContent>
                        {CATEGORIES.map(cat => <SelectItem key={cat.id} value={cat.id}>{cat.name}</SelectItem>)}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-2">
                    <Label className="text-[10px] font-bold text-slate-400 uppercase">Campo de Nota</Label>
                    <Select value={selectedField} onValueChange={setSelectedField}>
                      <SelectTrigger className="bg-slate-50 border-slate-100 h-10 text-xs"><SelectValue /></SelectTrigger>
                      <SelectContent>
                        {FIELDS.map(f => <SelectItem key={f.id} value={f.id}>{f.name}</SelectItem>)}
                      </SelectContent>
                    </Select>
                  </div>
                </div>

                <Button 
                  onClick={generateScript} 
                  disabled={!selectedTurma || !selectedArea} 
                  className="w-full bg-indigo-600 hover:bg-indigo-700 py-7 text-md font-bold rounded-2xl shadow-lg mt-4 transition-all active:scale-95"
                >
                  <ClipboardCopy className="w-5 h-5 mr-2" /> Copiar Script
                </Button>
              </CardContent>
            </Card>
          </div>

          <div className="lg:col-span-8">
            <Card className="border-none shadow-sm h-full min-h-[600px] overflow-hidden flex flex-col">
              <CardHeader className="border-b border-slate-100 flex flex-row items-center justify-between bg-white">
                <div>
                  <CardTitle className="text-lg">Visualização dos Dados</CardTitle>
                  <CardDescription>
                    {selectedTurma ? `Turma: ${selectedTurma}` : 'Carregue uma planilha para começar'}
                  </CardDescription>
                </div>
              </CardHeader>
              <CardContent className="p-0 flex-1 bg-white">
                {filteredData.length > 0 ? (
                  <div className="overflow-auto max-h-[700px]">
                    <Table>
                      <TableHeader className="bg-slate-50/50 sticky top-0 z-10">
                        <TableRow>
                          <TableHead className="font-bold text-slate-700 pl-6">Aluno</TableHead>
                          {selectedArea && (
                            <>
                              <TableHead className="text-center font-bold text-indigo-600">Notas ({selectedArea})</TableHead>
                              <TableHead className="text-center font-bold text-orange-600">Recs (R{selectedArea})</TableHead>
                            </>
                          )}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {filteredData.map((row, i) => (
                          <TableRow key={i} className="hover:bg-slate-50/50 transition-colors">
                            <TableCell className="font-medium text-slate-900 text-xs pl-6">{row._name || '-'}</TableCell>
                            {selectedArea && (
                              <>
                                <TableCell className="text-center font-bold text-indigo-600 text-xs">
                                    {row[selectedArea]?.join(' | ') || '-'}
                                </TableCell>
                                <TableCell className="text-center font-bold text-orange-600 text-xs">
                                    {row[`R${selectedArea}`]?.join(' | ') || '-'}
                                </TableCell>
                              </>
                            )}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-full min-h-[500px] text-slate-400">
                    <Layers className="w-16 h-16 mb-4 opacity-10" />
                    <p className="text-sm font-medium">Aguardando dados da planilha</p>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>
        </div>
        <MadeWithDyad />
      </div>
    </div>
  );
};

export default Index;
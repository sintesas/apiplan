<?php

namespace App\Http\Controllers;

use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Facades\Excel;
use Maatwebsite\Excel\Concerns\FromQuery;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ExcelVariasHojasController implements WithMultipleSheets
{
    public function exportExcel()
    {
        return Excel::download(new ExcelVariasHojasController(), 'cargos.xlsx');
    }

    public function sheets(): array
    {
        $sheets = [];

        $sheets[] = new class implements FromQuery, WithHeadings, WithTitle, WithColumnWidths, WithStyles {
            public function query()
            {
                return DB::table(DB::raw('vw_app_cargos'))
                ->select(
                    'cargo_id',
                    'cargo',
                    'descripcion',
                    'categoria',
                    'clase_cargo',
                    'cargo_ruta_id'
                )
                ->orderBy('cargo_id');
            }

            public function headings(): array
            {
                return [
                    'Cargo Id',
                    'Cargo',
                    'Descripción',
                    'Categoria',
                    'Clase Cargo',
                    'Tipo Ruta'
                ];
            }

            public function title(): string
            {
                return 'Datos Basicos';
            }

            public function columnWidths(): array
            {
                return [
                    'A' => 12,
                    'B' => 50,
                    'C' => 15,
                    'D' => 15,
                    'E' => 15,
                    'F' => 15
                ];
            }

            public function styles(Worksheet $sheet)
            {
                $sheet->getStyle('A1:F1')->applyFromArray([
                    'font' => [
                        'bold' => true,
                        'size' => 12
                    ],
                ]);
            
                foreach (range('A', 'F') as $column) {
                    $sheet->getColumnDimension($column)->setWidth(15);
            
        
                    $sheet->getStyle($column)->getFont()->setSize(12);
            
                    $lastRow = $sheet->getHighestRow();
            
                    for ($row = 2; $row <= $lastRow; $row++) {
                        $cellValue = $sheet->getCell($column . $row)->getValue();
            
                        if ($cellValue === null || $cellValue === '') {
                            $sheet->setCellValue($column . $row, ' ');
                        }
                    }
            
                    
                }
            
                $sheet->setAutoFilter("A1:F$lastRow");
            }
        };

        $sheets[] = new class implements FromQuery, WithHeadings, WithTitle, WithColumnWidths, WithStyles {
            public function query()
            {
                return DB::table(DB::raw('vw_app_cargos_configuracion_excel'))
                ->select(
                    'cargo_configuracion_id',
                    'grado',
                    'cargo',
                    'clasificacion',
                    'requisito_cargo',
                    'cuerpo',
                    'especialidad',
                    'area',
                    'educacion',
                    'conocimiento',
                    'experiencia',
                    'competencia',
                    'observaciones'
                )
                ->orderBy('cargo_configuracion_id');
            }

            public function headings(): array
            {
                return [
                    'Cargo Configuración Id',
                    'Grado',
                    'Cargo',
                    'Clasificación de la Función',
                    'Requisito Cargo',
                    'Cuerpo',
                    'Especialidad',
                    'Area',
                    'Educacion',
                    'Conocimiento',
                    'Experiencia',
                    'Competencia',
                    'Observaciones'
                ];
            }

            public function title(): string
            {
                return 'Configuración';
            }

            public function columnWidths(): array
            {
                return [
                    'A' => 12,
                    'B' => 15,
                    'C' => 15,
                    'D' => 15,
                    'E' => 15,
                    'F' => 15,
                    'G' => 15,
                    'H' => 15,
                    'I' => 15,
                    'J' => 15,
                    'K' => 15,
                    'L' => 15,
                    'M' => 15
                ];
            }

            public function styles(Worksheet $sheet)
            {
                $sheet->getStyle('A1:M1')->applyFromArray([
                    'font' => [
                        'bold' => true,
                        'size' => 12, 
                    ],
                ]);
            
                foreach (range('A', 'M') as $column) {
                    $sheet->getColumnDimension($column)->setWidth(15);            
        
                    $sheet->getStyle($column)->getFont()->setSize(12);
            
                    $lastRow = $sheet->getHighestRow();
            
                    for ($row = 2; $row <= $lastRow; $row++) {
                        $cellValue = $sheet->getCell($column . $row)->getValue();
            
                        if ($cellValue === null || $cellValue === '') {
                            $sheet->setCellValue($column . $row, ' ');
                        }
                    }
                }
            
                $sheet->setAutoFilter("A1:M$lastRow");
            }
        };

        $sheets[] = new class implements FromQuery, WithHeadings, WithTitle, WithColumnWidths, WithStyles {
            public function query()
            {
                return DB::table(DB::raw('vw_app_ubicacion_cargos'))
                ->select(
                    'ubicacion_cargo_id',
                    'cargo_configuracion_id',
                    'grado',
                    'cargo_id',
                    'cargo',
                    'nivel1',
                    'nivel2',
                    'nivel3',
                    'nivel4',
                    'nivel5',
                    'nivel6',
                    'nivel7',
                    'nivel8',
                    'puesto_cantidad',
                    'cargo_jefe_inmediato_id',
                    'cargo_jefe_inmediato'
                )
                ->orderBy('ubicacion_cargo_id');
            }

            public function headings(): array
            {
                return [
                    'Ubicacion Cargo Id',
                    'Cargo Configuracion Id',
                    'Grado',
                    'Cargo Id',
                    'Cargo',                    
                    'Nivel 1',
                    'Nivel 2',
                    'Nivel 3',
                    'Nivel 4',
                    'Nivel 5',
                    'Nivel 6',
                    'Nivel 7',
                    'Nivel 8',
                    'Puesto Cantidad',
                    'Cargo Jefe Inmediato Id',
                    'Cargo Jefe Inmediato'
                ];
            }

            public function title(): string
            {
                return 'Ubicaciones';
            }

            public function columnWidths(): array
            {
                return [
                    'A' => 10,
                    'B' => 10,
                    'C' => 10,
                    'D' => 10,
                    'E' => 15,
                    'F' => 10,
                    'G' => 15,
                    'H' => 10,
                    'I' => 15,
                    'J' => 10,
                    'K' => 15,
                    'L' => 10,
                    'M' => 15,
                    'N' => 15,
                    'O' => 15,
                    'P' => 15
                ];
            }

            public function styles(Worksheet $sheet)
            {
                $sheet->getStyle('A1:P1')->applyFromArray([
                    'font' => [
                        'bold' => true,
                        'size' => 12, 
                    ],
                ]);
            
                foreach (range('A', 'P') as $column) {
                    $sheet->getColumnDimension($column)->setWidth(15);            
        
                    $sheet->getStyle($column)->getFont()->setSize(12);
            
                    $lastRow = $sheet->getHighestRow();
            
                    for ($row = 2; $row <= $lastRow; $row++) {
                        $cellValue = $sheet->getCell($column . $row)->getValue();
            
                        if ($cellValue === null || $cellValue === '') {
                            $sheet->setCellValue($column . $row, ' ');
                        }
                    }
                }
            
                $sheet->setAutoFilter("A1:P$lastRow");
            }
        };

        $sheets[] = new class implements FromQuery, WithHeadings, WithTitle, WithColumnWidths, WithStyles {
            public function query()
            {
                return DB::table(DB::raw('vw_app_cargos_experiencias'))
                ->select(
                    'cargo_experiencia_id',
                    'cargo_configuracion_id',
                    'grado',
                    'cargo',                    
                    'cargo_previo_id',
                    'cargo_previo',
                    'anio',
                    'mes'
                )
                ->orderBy('cargo_experiencia_id');
            }

            public function headings(): array
            {
                return [
                    'Cargo Experiencia Id',
                    'Cargo Configuracion Id',
                    'Grado',
                    'Cargo',                    
                    'Cargo Id',
                    'Cargo Previo',
                    'Año',
                    'Mes'
                ];
            }

            public function title(): string
            {
                return 'Cargos Previos';
            }

            public function columnWidths(): array
            {
                return [
                    'A' => 12,
                    'B' => 15,
                    'C' => 15,
                    'D' => 50,
                    'E' => 15,
                    'F' => 15,
                    'G' => 15,
                    'H' => 15
                ];
            }

            public function styles(Worksheet $sheet)
            {
                $sheet->getStyle('A1:H1')->applyFromArray([
                    'font' => [
                        'bold' => true,
                        'size' => 12, 
                    ],
                ]);
            
                foreach (range('A', 'H') as $column) {
                    $sheet->getColumnDimension($column)->setWidth(15);            
        
                    $sheet->getStyle($column)->getFont()->setSize(12);
            
                    $lastRow = $sheet->getHighestRow();
            
                    for ($row = 2; $row <= $lastRow; $row++) {
                        $cellValue = $sheet->getCell($column . $row)->getValue();
            
                        if ($cellValue === null || $cellValue === '') {
                            $sheet->setCellValue($column . $row, ' ');
                        }
                    }
                }
            
                $sheet->setAutoFilter("A1:H$lastRow");
            }
        };

        return $sheets;
    }
}

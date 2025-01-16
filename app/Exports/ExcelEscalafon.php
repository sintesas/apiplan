<?php

namespace App\Exports;

use Illuminate\Support\Str;
use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Facades\Excel;
use Maatwebsite\Excel\Concerns\FromQuery;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ExcelEscalafon implements FromQuery, WithHeadings, WithTitle, WithColumnWidths, WithStyles
{
    protected $cuerpo_id;
    protected $especialidad_id;
    protected $area_id;

    public function __construct(int $cuerpo_id, int $especialidad_id, int $area_id)
    {
        $this->cuerpo_id = $cuerpo_id;
        $this->especialidad_id = $especialidad_id;
        $this->area_id = $area_id;
    }

    public function query()
    {
        if ($this->area_id == 0) {
            return DB::table(DB::raw('vw_reporte_escalafon'))
                ->select(
                    'cargo',
                    'grado',
                    'cantidad_puesto',
                    'sigla_cuerpo',
                    'sigla_especialidad',
                    'funcion',
                    'clasificacion',
                    'educacion',
                    'conocimiento',
                    'experiencia',
                    'experiencia_cargo',
                    'competencia'
                )
                ->where('cuerpo_id', $this->cuerpo_id)
                ->where('especialidad_id', $this->especialidad_id)
                ->orderBy('cargo');
        }
        else {
            return DB::table(DB::raw('vw_reporte_escalafon'))
            ->select(
                'cargo',
                'grado',
                'cantidad_puesto',
                'sigla_cuerpo',
                'sigla_especialidad',
                'sigla_area',
                'funcion',
                'clasificacion',
                'educacion',
                'conocimiento',
                'experiencia',
                'experiencia_cargo',
                'competencia'
            )
            ->where('cuerpo_id', $this->cuerpo_id)
            ->where('especialidad_id', $this->especialidad_id)
            ->where('area_id', $this->area_id)
            ->orderBy('cargo'); 
        }
    }

    public function headings(): array
    {
        if ($this->area_id == 0) {
            return [
                'CARGO',
                'GRADO',
                'CANTIDAD DE PUESTOS DE TRABAJO',
                'CUERPO',
                'ESPECIALIDAD',
                'FUNCIONES',
                'CLASIFICACIÓN DE LA FUNCIÓN',
                'EDUCACIÓN FORMAL',
                'CONOCIMIENTOS ADICIONALES',
                'EXPERIENCIA EN PROCESO',
                'EXPERIENCIA CARGOS PREVIOS',
                'COMPETENCIAS'
            ];
        }
        else {
            return [
                'CARGO',
                'GRADO',
                'CANTIDAD DE PUESTOS DE TRABAJO',
                'CUERPO',
                'ESPECIALIDAD',
                'ÁREA DE CONOCIMIENTO',
                'FUNCIONES',
                'CLASIFICACIÓN DE LA FUNCIÓN',
                'EDUCACIÓN FORMAL',
                'CONOCIMIENTOS ADICIONALES',
                'EXPERIENCIA EN PROCESO',
                'EXPERIENCIA CARGOS PREVIOS',
                'COMPETENCIAS'
            ];
        }
    }

    public function title(): string
    {
        return 'Escalafón de Cargos';
    }

    public function columnWidths(): array
    {
        if ($this->area_id == 0) {
            return [
                'A' => 20,
                'B' => 20,
                'C' => 20,
                'D' => 20,
                'E' => 20,
                'F' => 20,
                'G' => 20,
                'H' => 20,
                'I' => 20,
                'J' => 20,
                'K' => 20,
                'L' => 20,
            ];
        }
        else {
            return [
                'A' => 20,
                'B' => 20,
                'C' => 20,
                'D' => 20,
                'E' => 20,
                'F' => 20,
                'G' => 20,
                'H' => 20,
                'I' => 20,
                'J' => 20,
                'K' => 20,
                'L' => 20,
                'M' => 20
            ];
        }
    }

    public function styles(Worksheet $sheet)
    {
        if ($this->area_id == 0) {
            $sheet->getStyle('A1:L1')->applyFromArray([
                'font' => [
                    'bold' => true,
                    'size' => 12, 
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        'color' => ['argb' => '000000'],
                    ],
                ],
            ]);

            $lastRow = $sheet->getHighestRow();
            $sheet->getStyle('A1:L' . $lastRow)->applyFromArray([
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        'color' => ['argb' => '000000'],
                    ],
                ],
            ]);
        
            foreach (range('A', 'L') as $column) {
                $sheet->getColumnDimension($column)->setWidth(20);    

                $sheet->getStyle($column)->getFont()->setSize(12);
        
                $lastRow = $sheet->getHighestRow();
        
                for ($row = 2; $row <= $lastRow; $row++) {
                    $cellValue = $sheet->getCell($column . $row)->getValue();
        
                    if ($cellValue === null || $cellValue === '') {
                        $sheet->setCellValue($column . $row, ' ');
                    }
                }
            }
        
            $sheet->setAutoFilter("A1:L$lastRow");
        }
        else {
            $sheet->getStyle('A1:M1')->applyFromArray([
                'font' => [
                    'bold' => true,
                    'size' => 12, 
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        'color' => ['argb' => '000000'],
                    ],
                ],
            ]);

            $lastRow = $sheet->getHighestRow();
            $sheet->getStyle('A1:M' . $lastRow)->applyFromArray([
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                        'color' => ['argb' => '000000'],
                    ],
                ],
            ]);
        
            foreach (range('A', 'M') as $column) {
                $sheet->getColumnDimension($column)->setWidth(20);    

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
    }
}

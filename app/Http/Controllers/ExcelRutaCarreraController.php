<?php

namespace App\Http\Controllers;

use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Facades\Excel;
use Maatwebsite\Excel\Concerns\FromQuery;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ExcelRutaCarreraController extends Controller implements FromQuery, WithHeadings, WithTitle, WithColumnWidths, WithStyles
{
    public function export()
    {
        return Excel::download(new ExcelRutaCarreraController, 'rutas.xlsx');
    }

    public function query()
    {
        return DB::table(DB::raw('vw_app_rutas_excel'))
            ->select(
                'ruta_carrera_id',
                'nombreruta',
                'descripcion',
                'tipo_ruta',
                'siglaCuerpo',
                'cuerpo',
                'siglaEspecialidad',
                'especialidad',
                'siglaArea',
                'area',
                'cargo',
                'cargo_id',
                'descripcionCargo',
                'clase_cargo',
                'tipo_cargo'
            )
            ->orderBy('ruta_carrera_id');
    }

    public function headings(): array
    {
        return [
            'Ruta Carrera Id',
            'Nombre Ruta',
            'Descripción Ruta',
            'Tipo Ruta',
            'Sigla Cuerpo',
            'Cuerpo',
            'Sigla Especialidad',
            'Especialidad',
            'Sigla Área',
            'Área',
            'Cargo',
            'Cargo Id',
            'Descripción Cargo',
            'Clase Cargo',
            'Tipo Cargo'
        ];
    }

    public function title(): string
    {
        return 'Rutas';
    }

    public function columnWidths(): array
    {
        return [
            'A' => 10,
            'B' => 12,
            'C' => 12,
            'D' => 10,
            'E' => 12,
            'F' => 15,
            'G' => 12,
            'H' => 15,
            'I' => 15,
            'J' => 15,
            'K' => 15,
            'L' => 12,
            'M' => 12,
            'N' => 15,
            'O' => 15
        ];
    }

    public function styles(Worksheet $sheet)
    {
        $sheet->getStyle('A1:O1')->applyFromArray([
            'font' => [
                'bold' => true,
                'size' => 12, 
            ]
        ]);
    
        foreach (range('A', 'O') as $column) {
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
    
        $sheet->setAutoFilter("A1:O$lastRow");
    }
}

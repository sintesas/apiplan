<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Facades\Excel;
use Maatwebsite\Excel\Concerns\FromQuery;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

use App\Exports\ExcelEscalafon;

class ExcelEscalafonCargoController extends Controller {

    public function export(Request $request) {
        $cuerpo_id = $request->get('cuerpo_id');
        $especialidad_id = $request->get('especialidad_id');
        $area_id = $request->get('area_id');
        
        return Excel::download(new ExcelEscalafon($cuerpo_id, $especialidad_id, $area_id), 'escalafon.xlsx');
    }

    public function checkReforteEscalafon(Request $request) {
        $cuerpo_id = $request->get('cuerpo_id');
        $especialidad_id = $request->get('especialidad_id');
        $area_id = $request->get('area_id');

        if ($area_id == 0) {
            $db = DB::table(DB::raw('vw_reporte_escalafon'))
                            ->select('*')
                            ->where('cuerpo_id', $cuerpo_id)
                            ->where('especialidad_id', $especialidad_id)
                            ->exists();

            if ($db) {
                $response = json_encode(array('existe' => 1, 'tipo' => 1), JSON_NUMERIC_CHECK);
                $response = json_decode($response);
        
                return response()->json($response, 200);
            }
            else {
                $response = json_encode(array('existe' => 0, 'mensaje' => 'No hay información'), JSON_NUMERIC_CHECK);
                $response = json_decode($response);
        
                return response()->json($response, 200);
            }
        }
        else {
            $db = DB::table(DB::raw('vw_reporte_escalafon'))
                            ->select('*')
                            ->where('cuerpo_id', $cuerpo_id)
                            ->where('especialidad_id', $especialidad_id)
                            ->where('area_id', $area_id)
                            ->exists();

            if ($db) {
                $response = json_encode(array('existe' => 1, 'tipo' => 2), JSON_NUMERIC_CHECK);
                $response = json_decode($response);
        
                return response()->json($response, 200);
            }
            else {
                $response = json_encode(array('existe' => 0, 'mensaje' => 'No hay información'), JSON_NUMERIC_CHECK);
                $response = json_decode($response);
        
                return response()->json($response, 200);
            }
        }
    }
}

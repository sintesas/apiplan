<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Str;
use Illuminate\Support\Facades\DB;
use Barryvdh\DomPDF\Facade\Pdf;

class PdfController extends Controller
{
    public function getInformesPreview($id) {
        set_time_limit(0);
        
        $db = DB::select("exec pr_get_app_cargos_configuracion_by_cargo_grado_id ?", array($id));
        $cargo_configuracion_id = $db[0]->cargo_configuracion_id;
        $cargo_id = $db[0]->cargo_id;
        $cargo_grado_id = $db[0]->cargo_grado_id;
        $cargo = $db[0]->cargo;
        $grado = $db[0]->grado;
        $cuerpo = $db[0]->cuerpo;
        $especialidad = $db[0]->especialidad;
        $area = $db[0]->area;
        $educacion = $db[0]->educacion;
        $conocimiento = $db[0]->conocimiento;
        $experiencia = $db[0]->experiencia;
        $competencia = $db[0]->competencia;
        $requisito_cargo = $db[0]->requisito_cargo;

        $ubicacion = DB::select('exec pr_get_app_ubicaciones_cargos_by_id ?', array($cargo_configuracion_id));
        $expCar = DB::select("exec pr_get_app_cargos_experiencias_by_id ?", array($cargo_configuracion_id));

        $data = [
            'cargo' => $cargo,
            'grado' => $grado,
            'ubicacion' => $ubicacion,
            'cuerpo' => $cuerpo,
            'especialidad' => $especialidad,
            'area' => $area,
            'educacion' => $educacion,
            'conocimiento' => $conocimiento,
            'experiencia' => $experiencia,
            'competencia' => $competencia,
            'expCar' => $expCar,
            'requisito_cargo' => $requisito_cargo,
        ];

        $pdf = Pdf::loadView('perfil-cargo', compact('data'));
        return $pdf->stream();
    }

    public function getInformes($id) {
        set_time_limit(0);
        
        $db = DB::select("exec pr_get_app_cargos_configuracion_by_cargo_grado_id ?", array($id));
        $cargo_configuracion_id = $db[0]->cargo_configuracion_id;
        $cargo_id = $db[0]->cargo_id;
        $cargo_grado_id = $db[0]->cargo_grado_id;
        $cargo = $db[0]->cargo;
        $grado = $db[0]->grado;
        $cuerpo = $db[0]->cuerpo;
        $especialidad = $db[0]->especialidad;
        $area = $db[0]->area;
        $educacion = $db[0]->educacion;
        $conocimiento = $db[0]->conocimiento;
        $experiencia = $db[0]->experiencia;
        $competencia = $db[0]->competencia;
        $requisito_cargo = $db[0]->requisito_cargo;

        $ubicacion = DB::select('exec pr_get_app_ubicaciones_cargos_by_id ?', array($cargo_configuracion_id));
        $expCar = DB::select("exec pr_get_app_cargos_experiencias_by_id ?", array($cargo_configuracion_id));

        $data = [
            'cargo' => $cargo,
            'grado' => $grado,
            'ubicacion' => $ubicacion,
            'cuerpo' => $cuerpo,
            'especialidad' => $especialidad,
            'area' => $area,
            'educacion' => $educacion,
            'conocimiento' => $conocimiento,
            'experiencia' => $experiencia,
            'competencia' => $competencia,
            'expCar' => $expCar,
            'requisito_cargo' => $requisito_cargo,
        ];

        $pdf = Pdf::loadView('perfil-cargo', compact('data'));
        $uuid = Str::uuid();
        return $pdf->download($uuid . '.pdf');
    }
}

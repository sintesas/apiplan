<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class RutaCarrera extends Model
{
    use HasFactory;

    protected $table = 'tb_app_ruta_carrera';

    protected $primaryKey = 'ruta_carrera_id';

    protected $fillable = [
        'cuerpo_id,cuerpo,especialidad_id,especialidad,area_id,area,descripcion,tipo_categoria_id,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion,nombreruta'
    ];

    public function crud_ruta_carrera(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_ruta_carrera ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('ruta_carrera_id'),
                            $request->input('nombreruta'),
                            $request->input('cuerpo_id'),
                            $request->input('cuerpo'),
                            $request->input('especialidad_id'),
                            $request->input('especialidad'),
                            $request->input('area_id'),
                            $request->input('area'),
                            $request->input('descripcion'),
                            $request->input('tipo_categoria_id'),
                            $request->input('tipo_ruta_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')                            
                        ]);
        return $db;
    }

    public function get_ruta_carrera(Request $request) {
        $db = DB::select("exec pr_get_app_ruta_carrera ?,?", 
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function get_ruta_carrera_by_id(Request $request) {
        $db = DB::select("exec pr_get_app_ruta_carrera_by_id ?,?",
                        [
                            $request->input('especialidad_id'),
                            $request->input('tipo_ruta_id')
                        ]);
        return $db;
    }

    public function get_ruta_carrera_by_cargo(Request $request) {
        $db = DB::select("exec pr_get_ruta_carrera_by_cargo ?,?",
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function get_grados_by_especialidad(Request $request) {
        $db = DB::select("exec pr_get_app_grados_piramide_by_especialidad ?", array($request->input('especialidad_id')));
        return $db;
    }

    public function get_grados_detalle_requerimiento_piramide(Request $request) {
        $db = DB::select('exec pr_get_grados_detalle_requerimiento_piramide ?,?', array($request->input('especialidad_id'), $request->input('grado_id')));
        return $db;
    }

    public function get_grados_detalle_requisito_ley_piramide(Request $request) {
        $db = DB::select('exec pr_get_grados_detalle_requisito_ley_piramide ?,?', array($request->input('especialidad_id'), $request->input('grado_id')));
        return $db;
    }

    public function get_grados_detalle_by_especialidad(Request $request) {
        $db = DB::select("exec pr_get_grados_detalle_piramide_by_especialidad ?,?,?", array($request->input('especialidad_id'), $request->input('tipo_ruta_id'), $request->input('ruta_carrera_id')));
        return $db;
    }

    public function get_grados_detalle_cargos_piramide_by_ruta_carrera(Request $request) {
        $db = DB::select('exec pr_get_grados_detalle_cargos_piramide_by_ruta_carrera ?,?', array($request->input('ruta_carrera_id'), $request->input('grado_id')));
        return $db;
    }

    public function get_detalle_cargo_ruta_carrera(Request $request) {
        $db = DB::select('exec pr_get_detalle_cargo_ruta_carrera ?,?',
                        [
                            $request->input('cargo_id'),
                            $request->input('grado_id')
                        ]);
        return $db;
    }

    public function get_cuerpos_especialidades_areas_ruta_carrera() {
        $db = DB::select('exec pr_get_app_cuerpos_especialidades_areas_ruta_carrera');
        return $db;
    }

    public function get_app_especialidades_rutas() {
        $db = DB::select('exec pr_get_app_especialidades_rutas');
        return $db;
    }

    public function get_ruta_carrera_activos() {
        $db = DB::select('exec pr_get_ruta_carrera_activos');
        return $db;
    }

    public function get_width_by_rutas(Request $request) {
        $db = DB::select('exec pr_get_width_by_rutas ?', array($request->input('ruta_carrera_id')));

        return $db;
    }

    public function eliminar_ruta_carrera(Request $request) {
        $db = DB::select("exec pr_eliminar_ruta_carrera ?", array($request->input('ruta_carrera_id')));
        return $db;
    }
}

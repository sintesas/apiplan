<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class CargoConfiguracion extends Model
{
    use HasFactory;

    protected $table = 'tb_app_cargos_configuracion';

    protected $primaryKey = 'cargo_configuracion_id';

    protected $fillable = [
        'cargo_grado_id,puesto_cantidad,cargo_jefe_inmediato_id,cargo_jefe_inmediato,nivel1,nivel2,nivel3,nivel,nivel5,duracion,requisito_cargo,cuerpo,cuerpo_id,especialidad,especialidad_id,area,area_id,educacion,educacion_id,conocimiento,conocimiento_id,experiencia1,experiencia2,experiencia3,experiencia4,experiencia5,competencia1,competencia2,competencia3,competencia4,competencia5,observaciones,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_cargos_configuracion(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_cargos_configuracion ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('cargo_configuracion_id'),
                            $request->input('cargo_grado_id'),
                            $request->input('requisito_cargo'),
                            $request->input('cuerpo'),
                            $request->input('cuerpo_id'),
                            $request->input('especialidad'),
                            $request->input('especialidad_id'),
                            $request->input('area'),
                            $request->input('area_id'),
                            $request->input('educacion'),
                            $request->input('educacion_id'),
                            $request->input('conocimiento'),
                            $request->input('conocimiento_id'),
                            $request->input('experiencia'),
                            $request->input('experiencia_id'),
                            $request->input('competencia'),
                            $request->input('competencia_id'),
                            $request->input('observaciones'),
                            $request->input('clasificacion_id'),
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_cargos_configuracion_by_cargo_grado_id(Request $request) {
        $db = DB::select("exec pr_get_app_cargos_configuracion_by_cargo_grado_id ?", array($request->input('cargo_grado_id')));
        return $db;
    }

    public function get_cargos_configuracion_by_id(Request $request) {
        $db = DB::select("exec pr_get_app_cargos_configuracion_by_id ?", array($request->input('cargo_configuracion_id')));
        return $db;
    }
}

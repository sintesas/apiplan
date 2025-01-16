<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class Ruta extends Model
{
    use HasFactory;

    protected $table = 'tb_app_rutas';

    protected $primaryKey = 'ruta_id';

    protected $fillable = [
        'ruta_carrera_id,cargo_id,cargo_prev_id,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_rutas(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_rutas ?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('ruta_id'),
                            $request->input('ruta_carrera_id'),
                            $request->input('cargo_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_rutas_by_ruta_carrera_id(Request $request) {
        $db = DB::select("exec pr_get_app_rutas_by_ruta_carrera_id ?", array($request->input('ruta_carrera_id')));
        return $db;
    }

    public function get_rutas_ruta_carrera_id($ruta_carrera_id) {
        $db = DB::select("exec pr_get_app_rutas_by_ruta_carrera_id ?", array($ruta_carrera_id));
        return $db;
    }

    public function get_cargos_by_rutas(Request $request) {
        $db = DB::select("exec pr_get_cargos_by_rutas ?,?,?,?,?",
                        [
                            $request->input('tipo_ruta_id'),
                            $request->input('tipo_categoria_id'),
                            $request->input('especialidad_id'),
                            $request->input('cargo_ruta_id'),
                            $request->input('ruta_carrera_id')
                        ]);
        return $db;
    }

    public function get_app_rutas_full() {
        $db = DB::select("exec pr_get_app_rutas_full");
        return $db;
    }

}

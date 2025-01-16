<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class CargoExperiencia extends Model
{
    use HasFactory;

    protected $table = 'tb_app_cargos_experiencias';

    protected $primaryKey = 'cargo_experiencia_id';

    protected $fillable = [
        'cargo_experiencia_id,cargo_configuracion_id,cargo_previo_id,anio,mes,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_cargos_experiencias(Request $request, $evento)  {
        $db = DB::select('exec pr_crud_app_cargos_experiencias ?,?,?,?,?,?,?,?',
                        [
                            $evento,
                            $request->input('cargo_experiencia_id'),
                            $request->input('cargo_configuracion_id'),
                            $request->input('cargo_previo_id'),
                            $request->input('anio'),
                            $request->input('mes'),
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_app_cargos_experiencias_by_id(Request $request) {
        $db = DB::select("exec pr_get_app_cargos_experiencias_by_id ?", array($request->input('cargo_configuracion_id')));
        return $db;
    }

    public function eliminar_cargos_experiencias_by_id(Request $request) {
        $db = DB::select("exec pr_eliminar_cargos_experiencias_by_id ?", array($request->input('cargo_experiencia_id')));
        return $db;
    }
}

<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class CargoGrado extends Model
{
    use HasFactory;

    protected $table = 'tb_app_cargos_grados';

    protected $primaryKey = 'cargo_grado_id';

    protected $fillable = [
        'cargo_id,grado_id,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_cargos_grados(Request $request, $evento)  {
        $db = DB::select("exec pr_crud_app_cargos_grados ?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('cargo_grado_id'),
                            $request->input('cargo_id'),
                            $request->input('grado_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_cargos_grados_by_cargo_id(Request $request) {
        $db = DB::select("exec pr_get_app_cargos_grados_by_cargo_id ?", array($request->input('cargo_id')));
        return $db;
    }

    public function eliminar_cargos_grados_by_id(Request $request) {
        $db = DB::select("exec pr_eliminar_cargos_grados_by_id ?", array($request->input('cargo_grado_id')));
        return $db;
    }
}

<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class Cargo extends Model
{
    use HasFactory;

    protected $table = 'tb_app_cargos';

    protected $primaryKey = 'cargo_id';

    protected $fillable = [
        'cargo,descripcion,categoria_id,clase_categoria_id,cargo_ruta_id,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_cargo(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_cargos ?,?,?,?,?,?,?,?,?,?", 
                        [
                            $evento,
                            $request->input('cargo_id'),
                            $request->input('cargo'),
                            $request->input('descripcion'),
                            $request->input('categoria_id'),
                            $request->input('clase_cargo_id'),
                            $request->input('cargo_ruta_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_cargos(Request $request) {
        $db = DB::select("exec pr_get_app_cargos ?,?",
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 2000
                        ]);
        return $db;
    }

    public function get_cargos_full() {
        $db = DB::select("exec pr_get_app_cargos_full");
        return $db;
    }

    public function get_detalle_cargos(Request $request) {
        $db = DB::select("exec pr_get_detalle_cargos ?", array($request->input('cargo_id')));
        return $db;
    }

    public function get_cargos_by_id(Request $request) {
        $db = DB::select('exec pr_get_cargos_by_id ?', array($request->input('cargo_id')));
        return $db;
    }

    public function get_listas_niveles() {
        $db = DB::select("exec pr_get_listas_niveles");
        return $db;
    }

    public function eliminar_cargo(Request $request) {
        $db = DB::select("exec pr_eliminar_cargo ?", array($request->input('cargo_id')));
        return $db;
    }
}

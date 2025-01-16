<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class Grado extends Model
{
    use HasFactory;

    protected $table = 'tb_app_grados';

    protected $primaryKey = 'grado_id';

    protected $fillable = [
        'grado,descripcion,duracion,nivel_id,grado_previo_id,categoria_id,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_grados(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_grados ?,?,?,?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('grado_id'),
                            $request->input('grado'),
                            $request->input('descripcion'),
                            $request->input('duracion'),
                            $request->input('nivel_id'),
                            $request->input('grado_previo_id'),
                            $request->input('categoria_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        
        return $db;
    }

    public function get_grados(Request $request) {
        $db = DB::select("exec pr_get_app_grados ?,?",
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function get_grados_full() {
        $db = DB::select("exec pr_get_app_grados_full");
        return $db;
    }

    public function get_detalle_grados(Request $request) {
        $db = DB::select("exec pr_get_detalle_grados ?", array($request->input('grado_id')));
        return $db;
    }
}

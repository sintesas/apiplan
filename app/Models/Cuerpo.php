<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class Cuerpo extends Model
{
    use HasFactory;

    protected $table = 'tb_app_cuerpos';

    protected $primaryKey = 'cuerpo_id';

    protected $fillable = [
        'tipo_categoria_id,sigla,cuerpo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_cuerpos(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_cuerpos ?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('cuerpo_id'),
                            $request->input('tipo_categoria_id'),
                            $request->input('sigla'),
                            $request->input('cuerpo'),
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_cuerpos(Request $request) {
        $db = DB::select("exec pr_get_app_cuerpos ?,?", 
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function get_cuerpos_full() {
        $db = DB::select("exec pr_get_app_cuerpos_full");
        return $db;
    }

    public function get_cuerpos_by_categoria(Request $request) {
        $db = DB::select("exec pr_get_app_cuerpos_by_categoria ?", array($request->input('tipo_categoria_id')));
        return $db;
    }
}

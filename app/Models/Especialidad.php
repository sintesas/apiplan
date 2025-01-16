<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class Especialidad extends Model
{
    use HasFactory;

    protected $table = 'tb_app_especialidades';

    protected $primaryKey = 'especialidad_id';

    protected $fillable = [
        'tipo_categoria_id,sigla,especialidad,cuerpo_id,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_especialidades(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_especialidades ?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('especialidad_id'),
                            $request->input('tipo_categoria_id'),
                            $request->input('sigla'),
                            $request->input('especialidad'),
                            $request->input('cuerpo_id'),
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_especialidades(Request $request) {
        $db = DB::select("exec pr_get_app_especialidades ?,?", 
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function get_especialidades_full() {
        $db = DB::select("exec pr_get_app_especialidades_full");
        return $db;
    }

    public function get_especialidades_by_categoria_cuerpo(Request $request) {
        $db = DB::select("exec pr_get_app_especialidades_by_categoria_cuerpo ?,?",
                        [
                            $request->input('tipo_categoria_id'),
                            $request->input('cuerpo_id')
                        ]);
        return $db;
    }

    public function get_especialidades_by_categoria(Request $request) {
        $db = DB::select("exec pr_get_app_especialidades_by_categoria ?",
                        [
                            $request->input('tipo_categoria_id')
                        ]);
        return $db;
    }
}

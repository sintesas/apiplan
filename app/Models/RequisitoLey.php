<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class RequisitoLey extends Model
{
    use HasFactory;

    protected $table = 'tb_app_requisitos_ley';

    protected $primaryKey = 'requisito_ley_id';

    protected $fillable = [
        'requisito_ley,categoria_id,cuerpo_id,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_requisitos_ley(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_requisitos_ley ?,?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('requisito_ley_id'),
                            $request->input('requisito_ley'),
                            $request->input('descripcion'),
                            $request->input('categoria_id'),
                            $request->input('cuerpo_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_requisitos_ley(Request $request) {
        $db = DB::select("exec pr_get_app_requisitos_ley ?,?", 
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function crud_requisitos_especialidades(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_requisitos_especialidades ?,?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('requisito_especialidad_id'),
                            $request->input('requisito_ley_id'),
                            $request->input('categoria_id'),
                            $request->input('cuerpo_id'),
                            $request->input('especialidad_id'),
                            $request->input('checked') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_requisitos_especialidades_by_id(Request $request) {
        $db = DB::select("exec pr_get_app_requisitos_especialidades_by_id ?", 
                        [
                            $request->input('requisito_ley_id'),
                        ]);
        return $db;
    }

    public function crud_requisitos_grados(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_requisitos_grados ?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('requisito_grado_id'),
                            $request->input('requisito_ley_id'),
                            $request->input('categoria_id'),
                            $request->input('grado_id'),
                            $request->input('checked') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_requisitos_grados_by_id(Request $request) {
        $db = DB::select("exec pr_get_app_requisitos_grados_by_id ?", 
                        [
                            $request->input('requisito_ley_id'),
                        ]);
        return $db;
    }
}

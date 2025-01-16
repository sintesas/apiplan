<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class PerfilUsuario extends Model
{
    use HasFactory;

    protected $table = 'tb_app_perfiles_usuarios';

    protected $primaryKey = 'perfil_usuario_id';

    protected $fillable = [
        'perfil_id,usuario_id,tipo_perfil_id,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_perfiles_usuarios(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_perfiles_usuarios ?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('perfil_usuario_id'),
                            $request->input('perfil_id'),
                            $request->input('usuario_id'),
                            $request->input('tipo_perfil_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_perfiles_usuarios(Request $request) {
        $db = DB::select("exec pr_get_app_perfiles_usuarios ?,?", 
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }
}

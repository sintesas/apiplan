<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class UsuarioRol extends Model
{
    use HasFactory;

    protected $table = 'tb_app_usuarios_roles';

    protected $primaryKey = 'usuario_rol_id';

    protected $fillable = [
        "usuario_rol_id,usuario_id,rol_id,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion"
    ];

    public $timestamps = false;

    public function crud_usuarios_roles(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_usuarios_roles ?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('usuario_rol_id'),
                            $request->input('usuario_id'),
                            $request->input('rol_id'),
                            $request->input('rol_privilegio_id'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function get_usuarios_roles_by_usuario_id(Request $request) {
        $db = DB::select("exec pr_get_app_usuarios_roles_by_usuario_id ?", array($request->input('usuario_id')));

        return $db;
    }

    public function get_rol_privilegios_pantalla() {
        $db = DB::select("exec pr_get_rol_privilegios_pantalla");

        return $db;
    }

    public function eliminar_usuarios_roles_by_id(Request $request) {
        $db = DB::select("exec pr_eliminar_usuarios_roles_by_id ?", array($request->input('usuario_rol_id')));

        return $db;
    }
}

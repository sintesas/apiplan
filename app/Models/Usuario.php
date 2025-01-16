<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Support\Facades\DB;
use Illuminate\Foundation\Auth\User as Authenticatable;
use Illuminate\Notifications\Notifiable;
use Illuminate\Http\Request;
use Tymon\JWTAuth\Contracts\JWTSubject;
use PDO;
use PDOException;

class Usuario extends Authenticatable implements JWTSubject
{
    use HasFactory, Notifiable;

    protected $table = 'tb_app_usuarios';

    protected $primaryKey = 'usuario_id';

    protected $fillable = [
        'usuario,nombre_completo,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public $timestamps = false;

    public function getJWTIdentifier() {
        return $this->getKey();
    }

    public function getJWTCustomClaims() {
        return [];
    }

    public function crud_usuarios(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_usuarios ?,?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('usuario_id'),
                            $request->input('usuario'),
                            $request->input('nombre_completo'),
                            $request->input('email'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);
        return $db;
    }

    public function checkUsuario($usuario) {
        $result = DB::select("exec pr_check_usuario ?", array($usuario));
        if (count($result) > 0)
            return true;
        else return false;
    }

    public function getLoginUsuario($usuario) {
        $result = DB::select("exec pr_get_login_usuario ?", array($usuario));

        return $result;
    }

    public function get_usuarios(Request $request) {
        $db = DB::select("exec pr_get_app_usuarios ?,?",
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function get_usuarios_full() {
        $db = DB::select("exec pr_get_app_usuarios_full");
        return $db;
    }

    public function get_permisos_by_user(Request $request) {
        $db = DB::select("exec pr_get_permisos_by_usuario ?,?",
                        [
                            $request->input('usuario'),
                            $request->input('cod_modulo')
                        ]);
        return $db;
    }

    public function get_roles_by_usuario_id(Request $request) {
        $db = DB::select('exec pr_get_usuarios_roles_by_usuario_id ?', array($request->get('usuario_id')));
        return $db;
    }
}

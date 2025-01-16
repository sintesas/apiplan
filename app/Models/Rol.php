<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class Rol extends Model
{
    use HasFactory;

    protected $table = 'tb_app_roles';

    protected $primaryKey = 'rol_id';

    protected $fillable = [
        'rol_id,rol,descripcion,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_roles(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_roles ?,?,?,?,?,?,?",
                        [
                            $evento,
                            $request->input('rol_id'),
                            $request->input('rol'),
                            $request->input('descripcion'),
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador'),
                        ]);
        return $db;
    }

    public function get_roles(Request $request) {
        $db = DB::select("exec pr_get_app_roles ?,?", 
                        [
                            $request->input('filtro'),
                            $request->input('filtro') + 200
                        ]);
        return $db;
    }

    public function get_roles_activos() {
        $db = DB::select("exec pr_get_roles_activos");

        return $db;
    }
}

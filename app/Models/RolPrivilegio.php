<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Support\Facades\DB;
use Illuminate\Http\Request;

class RolPrivilegio extends Model
{
    use HasFactory;

    protected $table = 'tb_app_roles_privilegios';

    protected $primaryKey = 'rol_privilegio_id';

    protected $fillable = [
        'rol_privilegio_id,rol_id,num_pantalla,nombre_pantalla,consultar,crear,actualizar,eliminar,activo,usuario_creador,fecha_creacion,usuario_modificador,fecha_modificacion'
    ];

    public function crud_roles_privilegios(Request $request, $evento) {
        $db = DB::select("exec pr_crud_app_roles_privilegios ?,?,?,?,?,?,?,?,?,?,?,?,?", 
                        [
                            $evento,
                            $request->input('rol_privilegio_id'),
                            $request->input('rol_id'),
                            $request->input('num_pantalla'),
                            $request->input('modulo'),
                            $request->input('nombre_pantalla'),
                            $request->input('consultar') == true ? 'S' : 'N',
                            $request->input('crear') == true ? 'S' : 'N',
                            $request->input('actualizar') == true ? 'S' : 'N',
                            $request->input('eliminar') == true ? 'S' : 'N',
                            $request->input('activo') == true ? 'S' : 'N',
                            $request->input('usuario_creador'),
                            $request->input('usuario_modificador')
                        ]);

        return $db;
    }

    public function get_roles_privilegios_by_rol_id(Request $request) {
        $db = DB::select("exec pr_get_app_roles_privilegios_by_rol_id ?",  array($request->input('rol_id')));
        
        return $db;
    }

    public function eliminar_roles_privilegios_by_id(Request $request) {
        $db = DB::select("exec pr_eliminar_roles_privilegios_by_id ?",  array($request->input('rol_privilegio_id')));
        
        return $db;
    }

    public function get_modulos() {
        $db = DB::select("exec pr_get_modulos");
        return $db;
    }
}
